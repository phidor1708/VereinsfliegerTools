from __future__ import annotations

import argparse
import csv
import json
import re
import subprocess
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from tempfile import NamedTemporaryFile

import pandas as pd

APP_ROOT = Path(__file__).resolve().parent
UP_DIR = APP_ROOT / "up"
DEFAULT_MAEHPLAN_FILE = APP_ROOT / "ExcelAusschnittMaehplan2026.txt"
DEFAULT_MAEHPLAN_JSON = UP_DIR / "MaehplanStatic.json"
DEFAULT_SPRING_CLEANING_NOTE = "Alle: Frühjahrsputz am Modellfluggelände von 9:00 - 12:00"

CUSTOM_FIELD_ITEM_SEP = "§§§"
CUSTOM_FIELD_KEYVAL_SEP = "¤¤¤"

MERGED_ACTIVE_LABEL = "MFG-Aktive"
MERGED_ACTIVE_KEYS = {"aktiv", "aktiv_probe"}

MERGED_FOERDER_LABEL = "MFG-FörderUndEhrMitglieder"
MERGED_FOERDER_KEYS = {"ehrenmitglied", "foerdernd", "fordernd"}

SKIP_KEYS = {"ausgeschieden", "intern"}

EXCLUDED_EXTRA_COLUMNS_EXACT = {
    "Geburtsjahr",
    "Geburtstag",
}
EXCLUDED_EXTRA_COLUMNS_REGEX = [
    re.compile(r"rabatt", re.I),
]

GERMAN_MONTHS = {
    "januar": 1,
    "februar": 2,
    "maerz": 3,
    "märz": 3,
    "april": 4,
    "mai": 5,
    "juni": 6,
    "juli": 7,
    "august": 8,
    "september": 9,
    "oktober": 10,
    "november": 11,
    "dezember": 12,
}


@dataclass(frozen=True)
class MaehMemberInfo:
    partner: str
    term_nr_text: str
    term_dates_iso: tuple[str, ...]


@dataclass(frozen=True)
class MaehPlanInfo:
    term_0_date: str  # ISO date of the second April Saturday (Frühjahrsputz)
    spring_cleaning_note: str
    members: dict[str, MaehMemberInfo]


@dataclass(frozen=True)
class ContactRow:
    first_name: str
    last_name: str
    email: str
    status_key: str
    salutation: str
    title: str
    phone_home: str
    mobile_home: str
    phone_work: str
    mobile_work: str
    street: str
    postal_code: str
    city: str
    country: str
    birthday: str
    member_number: str
    entry_date: str
    exit_date: str
    functions: str
    remark: str
    custom_labeled_fields: tuple[tuple[str, str], ...]


def detect_delimiter(csv_path: Path) -> str:
    sample = csv_path.read_text(encoding="utf-8-sig", errors="ignore")[:4096]
    try:
        return csv.Sniffer().sniff(sample, delimiters=";,\t,").delimiter
    except csv.Error:
        return ";"


def find_latest_csv(input_csv: str | None) -> Path:
    if input_csv:
        path = Path(input_csv)
        if not path.is_absolute():
            path = APP_ROOT / path
        if not path.exists():
            raise FileNotFoundError(f"Input CSV not found: {path}")
        return path

    candidates = sorted(UP_DIR.glob("*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise FileNotFoundError("No CSV files found in ./up. Run download first.")
    return candidates[0]


def normalize_status(raw_status: str) -> str | None:
    value = (raw_status or "").strip().lower()
    value = value.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    value = re.sub(r"\s+", " ", value)
    if not value:
        return None

    compact = re.sub(r"[\s()]+", "", value)
    if compact == "aktivprobe":
        return "aktiv_probe"

    normalized = re.sub(r"[^a-z0-9]+", "_", value).strip("_")
    return normalized or None


def normalize_status_label(raw_status: str) -> str:
    return re.sub(r"\s+", " ", (raw_status or "").strip())


def split_name(name: str) -> tuple[str, str]:
    cleaned = (name or "").strip()
    if not cleaned:
        return "", ""

    if "," in cleaned:
        last, first = cleaned.split(",", 1)
        return first.strip(), last.strip()

    parts = cleaned.split()
    if len(parts) == 1:
        return "", parts[0]
    return " ".join(parts[:-1]), parts[-1]


def column_or_empty(df: pd.DataFrame, column_name: str) -> pd.Series:
    if column_name in df.columns:
        return df[column_name].astype(str).str.strip()
    return pd.Series([""] * len(df), index=df.index, dtype="string")


def normalize_date_value(raw_value: str) -> str:
    value = (raw_value or "").strip()
    if not value:
        return ""

    for fmt in (
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%d.%m.%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y/%m/%d",
    ):
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue

    return ""


def normalize_name_key(raw_name: str) -> str:
    value = (raw_name or "").strip().lower()
    value = value.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def build_contact_name_keys(first_name: str, last_name: str) -> list[str]:
    first = (first_name or "").strip()
    last = (last_name or "").strip()
    keys: list[str] = []
    if first and last:
        keys.append(normalize_name_key(f"{last} {first}"))
        keys.append(normalize_name_key(f"{first} {last}"))
    elif first:
        keys.append(normalize_name_key(first))
    elif last:
        keys.append(normalize_name_key(last))

    deduped: list[str] = []
    for key in keys:
        if key and key not in deduped:
            deduped.append(key)
    return deduped


def parse_german_date(raw_text: str) -> date | None:
    text = (raw_text or "").strip()
    if not text:
        return None

    match = re.search(r"(\d{1,2})\.\s*([A-Za-zÄÖÜäöüß]+)\s*(\d{4})", text)
    if not match:
        return None

    day = int(match.group(1))
    month_name = match.group(2).strip().lower()
    month_name = month_name.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
    year = int(match.group(3))

    month = GERMAN_MONTHS.get(month_name)
    if not month:
        return None

    try:
        return date(year, month, day)
    except ValueError:
        return None


def second_april_weekend(year: int) -> tuple[date, date]:
    first_april = date(year, 4, 1)
    saturday_offset = (5 - first_april.weekday()) % 7
    first_saturday = first_april + timedelta(days=saturday_offset)
    second_saturday = first_saturday + timedelta(days=7)
    second_sunday = second_saturday + timedelta(days=1)
    return second_saturday, second_sunday


def format_april_second_weekend(year: int) -> str:
    saturday, sunday = second_april_weekend(year)
    return f"Frühjahrsputz: {saturday.strftime('%d.%m.%Y')}–{sunday.strftime('%d.%m.%Y')}"


def second_april_saturday_iso(year: int) -> str:
    saturday, _ = second_april_weekend(year)
    return saturday.isoformat()


def load_maehplan_info(maehplan_file: Path | None, fallback_year: int) -> MaehPlanInfo | None:
    if not maehplan_file or not maehplan_file.exists():
        return None

    lines = maehplan_file.read_text(encoding="utf-8", errors="ignore").splitlines()
    by_member: dict[str, list[tuple[str, date | None, int]]] = defaultdict(list)

    term_nr = 0
    for raw_line in lines:
        if "Samstag" not in raw_line:
            continue

        columns = [part.strip() for part in raw_line.split("\t")]
        if not columns:
            continue

        current_date = parse_german_date(columns[0])
        normalized_line = normalize_name_key(raw_line)
        if "fruehjahrsputz" in normalized_line or "fuehjahrsputz" in normalized_line:
            continue

        name1 = columns[1] if len(columns) > 1 else ""
        name2 = columns[3] if len(columns) > 3 else ""
        if not name1 or not name2:
            continue

        term_nr += 1
        key1 = normalize_name_key(name1)
        key2 = normalize_name_key(name2)
        if key1 and key2:
            by_member[key1].append((name2, current_date, term_nr))
            by_member[key2].append((name1, current_date, term_nr))

    resolved_members: dict[str, MaehMemberInfo] = {}
    for key, assignments in by_member.items():
        partner_counter = Counter(partner.strip() for partner, _, _ in assignments if partner.strip())
        partner = partner_counter.most_common(1)[0][0] if partner_counter else ""

        unique_terms: dict[int, date | None] = {}
        for _, term_date, number in assignments:
            unique_terms[number] = term_date

        sorted_numbers = sorted(unique_terms.keys())
        term_nr_text = ", ".join(str(number) for number in sorted_numbers)
        term_dates_iso: list[str] = []
        for number in sorted_numbers:
            term_date = unique_terms[number]
            if term_date:
                term_dates_iso.append(term_date.isoformat())

        resolved_members[key] = MaehMemberInfo(
            partner=partner,
            term_nr_text=term_nr_text,
            term_dates_iso=tuple(term_dates_iso),
        )

    return MaehPlanInfo(
        term_0_date=second_april_saturday_iso(fallback_year),
        spring_cleaning_note=DEFAULT_SPRING_CLEANING_NOTE,
        members=resolved_members,
    )


def save_maehplan_json(info: MaehPlanInfo, path: Path) -> None:
    data: dict = {
        "Mähtermin_0": info.term_0_date,
        "FrühjahrsputzInfo": info.spring_cleaning_note,
    }
    data.update({
        key: {
            "Mähpartner": member.partner,
            "MähterminNr": member.term_nr_text,
        }
        for key, member in info.members.items()
    })
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_maehplan_json(path: Path) -> MaehPlanInfo | None:
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    term_0_str = data.get("Mähtermin_0", "")
    try:
        term_0 = date.fromisoformat(term_0_str) if term_0_str else None
    except ValueError:
        term_0 = None
    spring_cleaning_note = str(data.get("FrühjahrsputzInfo", DEFAULT_SPRING_CLEANING_NOTE)).strip()
    members: dict[str, MaehMemberInfo] = {}
    for key, member_data in data.items():
        if key in {"Mähtermin_0", "FrühjahrsputzInfo"}:
            continue
        term_nr_text = member_data.get("MähterminNr", "")
        term_dates_iso: tuple[str, ...] = ()
        if term_0:
            dates = []
            for part in term_nr_text.split(","):
                part = part.strip()
                if part.isdigit():
                    d = term_0 + timedelta(weeks=int(part))
                    dates.append(d.isoformat())
            term_dates_iso = tuple(dates)
        members[key] = MaehMemberInfo(
            partner=member_data.get("Mähpartner", ""),
            term_nr_text=term_nr_text,
            term_dates_iso=term_dates_iso,
        )
    return MaehPlanInfo(
        term_0_date=term_0_str,
        spring_cleaning_note=spring_cleaning_note,
        members=members,
    )


def sanitize_custom_field_part(value: str) -> str:
    text = (value or "").replace("\t", " ").replace("\n", " ").strip()
    text = text.replace(CUSTOM_FIELD_ITEM_SEP, " ").replace(CUSTOM_FIELD_KEYVAL_SEP, " ")
    return text


def encode_custom_labeled_fields(fields: tuple[tuple[str, str], ...]) -> str:
    encoded_parts: list[str] = []
    for label, value in fields:
        clean_label = sanitize_custom_field_part(label)
        clean_value = sanitize_custom_field_part(value)
        if not clean_label or not clean_value:
            continue
        encoded_parts.append(f"{clean_label}{CUSTOM_FIELD_KEYVAL_SEP}{clean_value}")
    return CUSTOM_FIELD_ITEM_SEP.join(encoded_parts)


def custom_labeled_fields_to_text(fields: tuple[tuple[str, str], ...]) -> str:
    return " | ".join(f"{label}: {value}" for label, value in fields)


def prioritize_custom_labeled_fields(fields: list[tuple[str, str]]) -> list[tuple[str, str]]:
    priority = {
        "Mähtermin_1": 0,
        "Mähtermin_2": 1,
        "Eintrittsdatum": 2,
        "Austrittsdatum": 3,
    }
    indexed = list(enumerate(fields))
    indexed.sort(key=lambda item: (priority.get(item[1][0], 100), item[0]))
    return [field for _, field in indexed]


def should_include_extra_column(column_name: str) -> bool:
    name = (column_name or "").strip()
    if not name:
        return False
    if name in EXCLUDED_EXTRA_COLUMNS_EXACT:
        return False
    return not any(pattern.search(name) for pattern in EXCLUDED_EXTRA_COLUMNS_REGEX)


def build_custom_labeled_fields(row: pd.Series, extra_columns: list[str]) -> tuple[tuple[str, str], ...]:
    fields: list[tuple[str, str]] = []

    entry_iso = str(row.get("entry_date_iso", "")).strip()
    exit_iso = str(row.get("exit_date_iso", "")).strip()
    if entry_iso:
        fields.append(("Eintrittsdatum", entry_iso))
    elif str(row.get("entry_date", "")).strip():
        fields.append(("Eintrittsdatum", str(row.get("entry_date", "")).strip()))
    if exit_iso:
        fields.append(("Austrittsdatum", exit_iso))
    elif str(row.get("exit_date", "")).strip():
        fields.append(("Austrittsdatum", str(row.get("exit_date", "")).strip()))

    for column_name in extra_columns:
        value = str(row.get(column_name, "")).strip()
        if not value:
            continue
        fields.append((column_name, value))

    return tuple(fields)


def build_contact_note(contact: ContactRow, label: str) -> str:
    parts = [f"Vereinsflieger Verteiler: {label}"]
    if contact.member_number:
        parts.append(f"MitgliedsNr: {contact.member_number}")
    if contact.remark:
        parts.append(f"Bemerkung: {contact.remark}")
    return " | ".join(parts)


def maybe_download_member_csv(ask_download: bool) -> None:
    if not ask_download or not sys.stdin.isatty():
        return

    answer = input("Neue Mitgliederdaten jetzt herunterladen? [j/N]: ").strip().lower()
    if answer not in {"j", "ja", "y", "yes"}:
        return

    download_script = APP_ROOT / "download_mitglieder_csv.py"
    if not download_script.exists():
        raise FileNotFoundError(f"Download script not found: {download_script}")

    subprocess.run(
        [sys.executable, str(download_script), "--headless", "--timeout-ms", "90000"],
        cwd=str(APP_ROOT),
        check=True,
    )


def load_contacts_dataframe(csv_path: Path) -> pd.DataFrame:
    delimiter = detect_delimiter(csv_path)
    df = pd.read_csv(csv_path, sep=delimiter, dtype=str, encoding="utf-8-sig").fillna("")
    df.columns = [str(col).strip() for col in df.columns]
    source_columns = list(df.columns)

    required = ["Mailadresse", "Mitgliedsstatus"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing expected columns: {missing}. Available: {list(df.columns)}")

    if not ({"Vorname", "Nachname"}.issubset(df.columns) or "Name" in df.columns):
        raise ValueError(
            "Expected name columns not found. Need either 'Vorname' and 'Nachname' or 'Name'."
        )

    name_series = column_or_empty(df, "Name")
    first_name_series = column_or_empty(df, "Vorname")
    last_name_series = column_or_empty(df, "Nachname")
    split_from_name = name_series.apply(split_name)

    df["first_name"] = first_name_series.where(first_name_series != "", split_from_name.apply(lambda x: x[0]))
    df["last_name"] = last_name_series.where(last_name_series != "", split_from_name.apply(lambda x: x[1]))

    mail_primary = column_or_empty(df, "Mailadresse")
    mail_alt_roundmail = column_or_empty(df, "abw. Rundmailadresse")
    mail_alt_invoice = column_or_empty(df, "Abw. Rechnungsemail")

    df["email"] = mail_primary
    for alt in (mail_alt_roundmail, mail_alt_invoice):
        has_valid_email = df["email"].str.contains("@", na=False)
        df["email"] = df["email"].where(has_valid_email, alt)

    df["salutation"] = column_or_empty(df, "Anrede")
    df["title"] = column_or_empty(df, "Titel")
    df["phone_home"] = column_or_empty(df, "Telefon (privat)")
    df["mobile_home"] = column_or_empty(df, "Mobil (privat)")
    df["phone_work"] = column_or_empty(df, "Telefon (geschäftl.)")
    df["mobile_work"] = column_or_empty(df, "Mobil (geschäftl.)")
    df["street"] = column_or_empty(df, "Straße")
    df["postal_code"] = column_or_empty(df, "Plz")
    df["city"] = column_or_empty(df, "Ort")
    df["country"] = column_or_empty(df, "Land")
    df["birthday"] = column_or_empty(df, "Geburtsdatum")
    df["member_number"] = column_or_empty(df, "MitgliedsNr")
    df["entry_date"] = column_or_empty(df, "Eintritt")
    df["exit_date"] = column_or_empty(df, "Austritt")
    df["functions"] = column_or_empty(df, "Funktionen")
    df["remark"] = column_or_empty(df, "Bemerkung")

    df["birthday_iso"] = df["birthday"].apply(normalize_date_value)
    df["entry_date_iso"] = df["entry_date"].apply(normalize_date_value)
    df["exit_date_iso"] = df["exit_date"].apply(normalize_date_value)

    mapped_source_columns = {
        "Name",
        "Vorname",
        "Nachname",
        "Mailadresse",
        "abw. Rundmailadresse",
        "Abw. Rechnungsemail",
        "Mitgliedsstatus",
        "Anrede",
        "Titel",
        "Telefon (privat)",
        "Mobil (privat)",
        "Telefon (geschäftl.)",
        "Mobil (geschäftl.)",
        "Straße",
        "Plz",
        "Ort",
        "Land",
        "Geburtsdatum",
        "MitgliedsNr",
        "Eintritt",
        "Austritt",
        "Funktionen",
        "Bemerkung",
    }
    extra_columns = [
        col
        for col in source_columns
        if col not in mapped_source_columns and should_include_extra_column(col)
    ]
    df["custom_labeled_fields"] = df.apply(
        lambda row: build_custom_labeled_fields(row, extra_columns),
        axis=1,
    )

    df["Mitgliedsstatus"] = df["Mitgliedsstatus"].astype(str).str.strip()

    df["status_key"] = df["Mitgliedsstatus"].apply(normalize_status)
    df["status_label"] = df["Mitgliedsstatus"].apply(normalize_status_label)
    df = df[df["status_key"].notna()].copy()
    df = df[df["email"].str.contains("@", na=False)].copy()

    return df


def to_contact_rows(df: pd.DataFrame, maehplan_info: MaehPlanInfo | None = None) -> dict[str, list[ContactRow]]:
    grouped: dict[str, list[ContactRow]] = {}
    for _, row in df.iterrows():
        key = row["status_key"]
        if key in SKIP_KEYS:
            continue
        original_label = str(row["status_label"]).strip()
        if key in MERGED_ACTIVE_KEYS:
            target_label = MERGED_ACTIVE_LABEL
        elif key in MERGED_FOERDER_KEYS:
            target_label = MERGED_FOERDER_LABEL
        else:
            target_label = original_label

        first_name = str(row["first_name"]).strip()
        last_name = str(row["last_name"]).strip()
        custom_labeled_fields = list(tuple(row["custom_labeled_fields"]))

        if maehplan_info:
            member_info: MaehMemberInfo | None = None
            for name_key in build_contact_name_keys(first_name, last_name):
                if name_key in maehplan_info.members:
                    member_info = maehplan_info.members[name_key]
                    break

            if member_info:
                if member_info.partner:
                    custom_labeled_fields.append(("Mähpartner", member_info.partner))
                if member_info.term_nr_text:
                    custom_labeled_fields.append(("MähterminNr", member_info.term_nr_text))
                if member_info.term_dates_iso:
                    custom_labeled_fields.append(("Mähtermin_1", member_info.term_dates_iso[0]))
                if len(member_info.term_dates_iso) > 1:
                    custom_labeled_fields.append(("Mähtermin_2", member_info.term_dates_iso[1]))

            custom_labeled_fields = prioritize_custom_labeled_fields(custom_labeled_fields)

        grouped.setdefault(target_label, []).append(
            ContactRow(
                first_name=first_name,
                last_name=last_name,
                email=str(row["email"]).strip(),
                status_key=key,
                salutation=str(row["salutation"]).strip(),
                title=str(row["title"]).strip(),
                phone_home=str(row["phone_home"]).strip(),
                mobile_home=str(row["mobile_home"]).strip(),
                phone_work=str(row["phone_work"]).strip(),
                mobile_work=str(row["mobile_work"]).strip(),
                street=str(row["street"]).strip(),
                postal_code=str(row["postal_code"]).strip(),
                city=str(row["city"]).strip(),
                country=str(row["country"]).strip(),
                birthday=str(row["birthday_iso"]).strip(),
                member_number=str(row["member_number"]).strip(),
                entry_date=str(row["entry_date"]).strip(),
                exit_date=str(row["exit_date"]).strip(),
                functions=str(row["functions"]).strip(),
                remark=str(row["remark"]).strip(),
                custom_labeled_fields=tuple(custom_labeled_fields),
            )
        )

    ordered_labels = sorted(grouped.keys(), key=lambda label: (0 if label == MERGED_ACTIVE_LABEL else 1 if label == MERGED_FOERDER_LABEL else 2, label.lower()))
    return {
        label: sorted(
            grouped[label],
            key=lambda c: (c.last_name.lower(), c.first_name.lower(), c.email.lower()),
        )
        for label in ordered_labels
    }


def slugify(text: str) -> str:
    value = text.lower().replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    value = re.sub(r"[^a-z0-9]+", "_", value).strip("_")
    return value


def vcard_escape(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
    return escaped.replace("\n", "\\n")


def write_exports(grouped: dict[str, list[ContactRow]], output_dir: Path) -> dict[str, dict[str, Path | int]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary: dict[str, dict[str, Path | int]] = {}

    for label, contacts in grouped.items():
        base = slugify(label)
        csv_path = output_dir / f"contacts_{base}.csv"
        vcf_path = output_dir / f"contacts_{base}.vcf"

        export_rows = [
            {
                "First Name": c.first_name,
                "Last Name": c.last_name,
                "E-mail Address": c.email,
                "Phone (Home)": c.phone_home,
                "Mobile (Home)": c.mobile_home,
                "Phone (Work)": c.phone_work,
                "Mobile (Work)": c.mobile_work,
                "Street": c.street,
                "Postal Code": c.postal_code,
                "City": c.city,
                "Country": c.country,
                "Birthday": c.birthday,
                "Title": c.title,
                "Salutation": c.salutation,
                "Member Number": c.member_number,
                "Entry Date": c.entry_date,
                "Exit Date": c.exit_date,
                "Functions": c.functions,
                "Custom Labels": custom_labeled_fields_to_text(c.custom_labeled_fields),
                "Notes": build_contact_note(c, label),
            }
            for c in contacts
        ]
        pd.DataFrame(export_rows).to_csv(csv_path, index=False, encoding="utf-8")

        with vcf_path.open("w", encoding="utf-8") as f:
            for c in contacts:
                full_name = " ".join(part for part in [c.first_name, c.last_name] if part).strip() or c.email
                f.write("BEGIN:VCARD\n")
                f.write("VERSION:3.0\n")
                f.write(f"N:{vcard_escape(c.last_name)};{vcard_escape(c.first_name)};;;\n")
                f.write(f"FN:{vcard_escape(full_name)}\n")
                f.write(f"EMAIL;TYPE=INTERNET;TYPE=WORK:{vcard_escape(c.email)}\n")

                if c.phone_home:
                    f.write(f"TEL;TYPE=HOME,VOICE:{vcard_escape(c.phone_home)}\n")
                if c.mobile_home:
                    f.write(f"TEL;TYPE=CELL:{vcard_escape(c.mobile_home)}\n")
                if c.phone_work:
                    f.write(f"TEL;TYPE=WORK,VOICE:{vcard_escape(c.phone_work)}\n")
                if c.mobile_work:
                    f.write(f"TEL;TYPE=WORK,CELL:{vcard_escape(c.mobile_work)}\n")

                if c.street or c.city or c.postal_code or c.country:
                    f.write(
                        "ADR;TYPE=HOME:;;"
                        f"{vcard_escape(c.street)};"
                        f"{vcard_escape(c.city)};;"
                        f"{vcard_escape(c.postal_code)};"
                        f"{vcard_escape(c.country)}\n"
                    )

                if c.birthday:
                    f.write(f"BDAY:{vcard_escape(c.birthday)}\n")

                if c.title:
                    f.write(f"TITLE:{vcard_escape(c.title)}\n")

                if c.functions:
                    f.write(f"ROLE:{vcard_escape(c.functions)}\n")

                f.write(f"NOTE:{vcard_escape(build_contact_note(c, label))}\n")
                f.write("END:VCARD\n")

            summary[label] = {"csv": csv_path, "vcf": vcf_path, "count": len(contacts)}

    return summary


def run_osascript(script: str, args: list[str]) -> str:
    result = subprocess.run(["osascript", "-e", script, *args], capture_output=True, text=True)
    if result.returncode != 0:
        stderr = (result.stderr or "").strip()
        raise RuntimeError(stderr or "AppleScript call failed")
    return (result.stdout or "").strip()


def delete_existing_vf_import_groups(prefix: str) -> list[str]:
    """Delete all Contacts groups whose name contains 'VF Import' (identifies auto-created lists)."""
    find_and_delete_script = """
    on run argv
      set marker to item 1 of argv
      set deletedNames to {}
      tell application \"Contacts\"
        set allGroups to every group
        repeat with g in allGroups
          set gName to name of g
          if gName contains marker then
            set end of deletedNames to gName
            delete g
          end if
        end repeat
        save
      end tell
      if (count of deletedNames) is 0 then
        return \"\"
      end if
      set AppleScript's text item delimiters to \"||\"
      set joined to deletedNames as text
      set AppleScript's text item delimiters to \"\"
      return joined
    end run
    """
    output = run_osascript(find_and_delete_script, ["VF Import"])
    if output:
        return [name for name in output.split("||") if name]
    return []


def create_contacts_group(group_name: str, contacts: list[ContactRow], label: str) -> None:
    if not contacts:
        return

    with NamedTemporaryFile("w", encoding="utf-8", suffix=".tsv", delete=False) as tmp:
        tmp_path = Path(tmp.name)
        for c in contacts:
            first = c.first_name.replace("\t", " ").replace("\n", " ").strip()
            last = c.last_name.replace("\t", " ").replace("\n", " ").strip()
            email = c.email.replace("\t", " ").replace("\n", " ").strip()
            phone_home = c.phone_home.replace("\t", " ").replace("\n", " ").strip()
            mobile_home = c.mobile_home.replace("\t", " ").replace("\n", " ").strip()
            phone_work = c.phone_work.replace("\t", " ").replace("\n", " ").strip()
            mobile_work = c.mobile_work.replace("\t", " ").replace("\n", " ").strip()
            street = c.street.replace("\t", " ").replace("\n", " ").strip()
            postal_code = c.postal_code.replace("\t", " ").replace("\n", " ").strip()
            city = c.city.replace("\t", " ").replace("\n", " ").strip()
            country = c.country.replace("\t", " ").replace("\n", " ").strip()
            birthday = c.birthday.replace("\t", " ").replace("\n", " ").strip()
            function_value = c.functions.replace("\t", " ").replace("\n", " ").strip()
            note = build_contact_note(c, label).replace("\t", " ").replace("\n", " ").strip()
            custom_labeled = encode_custom_labeled_fields(c.custom_labeled_fields)
            tmp.write(
                f"{first}\t{last}\t{email}\t{phone_home}\t{mobile_home}\t{phone_work}\t"
                f"{mobile_work}\t{street}\t{postal_code}\t{city}\t{country}\t{birthday}\t{function_value}\t{note}\t{custom_labeled}\n"
            )

    create_script = """
        on parseDateValue(rawText)
            try
                set cleaned to rawText as text
                set cleaned to my trimText(cleaned)
                if cleaned is "" then
                    return missing value
                end if

                if cleaned contains "/" then
                    set oldDelims to AppleScript's text item delimiters
                    set AppleScript's text item delimiters to "/"
                    set parts to text items of cleaned
                    set AppleScript's text item delimiters to oldDelims
                else if cleaned contains "." then
                    set oldDelims to AppleScript's text item delimiters
                    set AppleScript's text item delimiters to "."
                    set parts to text items of cleaned
                    set AppleScript's text item delimiters to oldDelims
                else
                    set oldDelims to AppleScript's text item delimiters
                    set AppleScript's text item delimiters to "-"
                    set parts to text items of cleaned
                    set AppleScript's text item delimiters to oldDelims
                end if

                if (count of parts) is not 3 then
                    return missing value
                end if

                set firstPart to item 1 of parts as text
                set secondPart to item 2 of parts as text
                set thirdPart to item 3 of parts as text

                set y to 0
                set m to 0
                set d to 0
                if (length of firstPart) is 4 then
                    set y to firstPart as integer
                    set m to secondPart as integer
                    set d to thirdPart as integer
                else
                    set d to firstPart as integer
                    set m to secondPart as integer
                    set y to thirdPart as integer
                    if y < 100 then
                        set y to y + 2000
                    end if
                end if

                set parsedDate to current date
                set year of parsedDate to y
                set month of parsedDate to m
                set day of parsedDate to d
                set time of parsedDate to 0
                return parsedDate
            on error
                return missing value
            end try
        end parseDateValue

        on trimText(valueText)
            set txt to valueText as text
            repeat while txt begins with " "
                set txt to text 2 thru -1 of txt
            end repeat
            repeat while txt ends with " "
                set txt to text 1 thru -2 of txt
            end repeat
            return txt
        end trimText

    on run argv
      set groupName to item 1 of argv
      set tsvPath to item 2 of argv

      tell application \"Contacts\"
        if (exists group groupName) then
          error \"Group already exists: \" & groupName
        end if

        set newGroup to make new group with properties {name:groupName}

                set fileRef to open for access (POSIX file tsvPath)
                try
                    set txt to read fileRef as «class utf8»
                on error number -39
                    set txt to ""
                end try
                close access fileRef

        set oldDelims to AppleScript's text item delimiters
        set AppleScript's text item delimiters to linefeed
        set allLines to text items of txt
        set AppleScript's text item delimiters to oldDelims

        repeat with lineText in allLines
          set lineText to contents of lineText
          if lineText is not \"\" then
            set oldDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to tab
            set fields to text items of lineText
            set AppleScript's text item delimiters to oldDelims

                        if (count of fields) is greater than or equal to 15 then
              set firstName to item 1 of fields
              set lastName to item 2 of fields
              set emailValue to item 3 of fields
                            set homePhone to item 4 of fields
                            set homeMobile to item 5 of fields
                            set workPhone to item 6 of fields
                            set workMobile to item 7 of fields
                            set streetValue to item 8 of fields
                            set zipValue to item 9 of fields
                            set cityValue to item 10 of fields
                            set countryValue to item 11 of fields
                            set birthdayValue to item 12 of fields
                            set functionValue to item 13 of fields
                            set noteValue to item 14 of fields
                            set customLabeledValue to item 15 of fields

              set newPerson to make new person with properties {first name:firstName, last name:lastName}

                            if emailValue is not "" then
                                make new email at end of emails of newPerson with properties {label:\"work\", value:emailValue}
                            end if

                            if homePhone is not "" then
                                make new phone at end of phones of newPerson with properties {label:\"home\", value:homePhone}
                            end if
                            if homeMobile is not "" then
                                make new phone at end of phones of newPerson with properties {label:\"mobile\", value:homeMobile}
                            end if
                            if workPhone is not "" then
                                make new phone at end of phones of newPerson with properties {label:\"work\", value:workPhone}
                            end if
                            if workMobile is not "" then
                                make new phone at end of phones of newPerson with properties {label:\"work mobile\", value:workMobile}
                            end if

                            if streetValue is not "" or zipValue is not "" or cityValue is not "" or countryValue is not "" then
                                make new address at end of addresses of newPerson with properties {label:\"home\", street:streetValue, zip:zipValue, city:cityValue, country:countryValue}
                            end if

                            if birthdayValue is not "" then
                                set parsedBirthDate to my parseDateValue(birthdayValue)
                                if parsedBirthDate is not missing value then
                                    try
                                        set birth date of newPerson to parsedBirthDate
                                    end try
                                end if
                            end if

                            if functionValue is not "" then
                                try
                                    set job title of newPerson to functionValue
                                end try
                            end if

                            if noteValue is not "" then
                                set note of newPerson to noteValue
                            end if

                            if customLabeledValue is not "" then
                                set oldDelims to AppleScript's text item delimiters
                                set AppleScript's text item delimiters to "§§§"
                                set customItems to text items of customLabeledValue
                                set AppleScript's text item delimiters to oldDelims

                                set customDateCount to 0

                                repeat with customItem in customItems
                                    set customItem to contents of customItem
                                    if customItem is not "" then
                                        set oldDelims to AppleScript's text item delimiters
                                        set AppleScript's text item delimiters to "¤¤¤"
                                        set customParts to text items of customItem
                                        set AppleScript's text item delimiters to oldDelims

                                        if (count of customParts) is greater than or equal to 2 then
                                            set customLabel to item 1 of customParts
                                            set customValue to item 2 of customParts
                                            if customLabel is not "" and customValue is not "" then
                                                set parsedCustomDate to my parseDateValue(customValue)
                                                if parsedCustomDate is not missing value then
                                                    if customDateCount < 2 then
                                                        try
                                                            make new custom date at end of custom dates of newPerson with properties {label:customLabel, value:parsedCustomDate}
                                                            save
                                                            set customDateCount to customDateCount + 1
                                                        on error
                                                            make new related name at end of related names of newPerson with properties {label:customLabel, value:customValue}
                                                        end try
                                                    else
                                                        make new related name at end of related names of newPerson with properties {label:customLabel, value:customValue}
                                                    end if
                                                else
                                                    make new related name at end of related names of newPerson with properties {label:customLabel, value:customValue}
                                                end if
                                            end if
                                        end if
                                    end if
                                end repeat
                            end if

              add newPerson to newGroup
            end if
          end if
        end repeat

        save
      end tell
      return \"ok\"
    end run
    """

    try:
        run_osascript(create_script, [group_name, str(tmp_path)])
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass


def apply_to_contacts(grouped: dict[str, list[ContactRow]], prefix: str) -> list[str]:
    run_stamp = datetime.now().strftime("%Y-%m-%d %H-%M")
    non_empty_labels = [label for label, contacts in grouped.items() if contacts]
    # Format: "{label} VF Import {timestamp}"
    group_names = [f"{label} VF Import {run_stamp}".strip() for label in non_empty_labels]

    deleted = delete_existing_vf_import_groups(prefix)
    if deleted:
        print(f"Deleted {len(deleted)} existing VF Import list(s): {', '.join(deleted)}")

    created: list[str] = []
    for label, group_name in zip(non_empty_labels, group_names):
        contacts = grouped[label]
        create_contacts_group(group_name, contacts, label)
        created.append(group_name)
    return created


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Create contact export lists from latest Vereinsflieger CSV and optionally create "
            "new Apple Contacts lists (without touching existing lists)."
        )
    )
    parser.add_argument("--input-csv", default=None, help="Input CSV path. Defaults to newest CSV in ./up.")
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Directory for generated list files. Defaults to ./up/contacts_exports/<timestamp>.",
    )
    parser.add_argument(
        "--apply-contacts",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Create new lists and contacts in Apple Contacts (default: true).",
    )
    parser.add_argument(
        "--contacts-prefix",
        default="VF Import",
        help="Prefix for new Apple Contacts list names.",
    )
    parser.add_argument(
        "--pull-data-from-vereinsflieger",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="Download fresh member CSV from Vereinsflieger before processing (default: false).",
    )
    parser.add_argument(
        "--maehplan-file",
        default=None,
        help="Path to mowing plan text file for Mähpartner/Mähtermin fields. Parsed result is saved to up/maehplan.json.",
    )
    args = parser.parse_args()

    if not args.input_csv:
        maybe_download_member_csv(args.pull_data_from_vereinsflieger)

    input_csv = find_latest_csv(args.input_csv)

    fallback_year = datetime.now().year
    maehplan_json_path = DEFAULT_MAEHPLAN_JSON

    # Resolve txt source: explicit arg, or default file if it exists
    maehplan_txt: Path | None = None
    if args.maehplan_file:
        maehplan_txt = Path(args.maehplan_file)
        if not maehplan_txt.is_absolute():
            maehplan_txt = APP_ROOT / maehplan_txt
    elif DEFAULT_MAEHPLAN_FILE.exists():
        maehplan_txt = DEFAULT_MAEHPLAN_FILE

    txt_info = load_maehplan_info(maehplan_txt, fallback_year)
    if txt_info is not None:
        save_maehplan_json(txt_info, maehplan_json_path)
        print(f"Mähplan data saved to {maehplan_json_path}")
        maehplan_info = txt_info
    else:
        maehplan_info = load_maehplan_json(maehplan_json_path)
    if args.output_dir:
        output_dir = Path(args.output_dir)
        if not output_dir.is_absolute():
            output_dir = APP_ROOT / output_dir
    else:
        output_dir = UP_DIR / "contacts_exports" / datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    df = load_contacts_dataframe(input_csv)
    grouped = to_contact_rows(df, maehplan_info=maehplan_info)
    summary = write_exports(grouped, output_dir)

    print(f"Input CSV: {input_csv}")
    print(f"Output directory: {output_dir}")
    for label, info in summary.items():
        print(
            f"- {label}: {info['count']} contacts -> {info['csv']} and {info['vcf']}"
        )

    if args.apply_contacts:
        created_groups = apply_to_contacts(grouped, args.contacts_prefix)
        print("Created Apple Contacts lists:")
        for name in created_groups:
            print(f"- {name}")
    else:
        print("Apple Contacts was not modified. Use --apply-contacts / omit --no-apply-contacts to create new lists.")


if __name__ == "__main__":
    main()
