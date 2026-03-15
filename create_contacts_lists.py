from __future__ import annotations

import argparse
import csv
import re
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

import pandas as pd

APP_ROOT = Path(__file__).resolve().parent
UP_DIR = APP_ROOT / "up"

MERGED_ACTIVE_LABEL = "MFG-Aktive"
MERGED_ACTIVE_KEYS = {"aktiv", "aktiv_probe"}


@dataclass(frozen=True)
class ContactRow:
    first_name: str
    last_name: str
    email: str
    status_key: str


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


def load_contacts_dataframe(csv_path: Path) -> pd.DataFrame:
    delimiter = detect_delimiter(csv_path)
    df = pd.read_csv(csv_path, sep=delimiter, dtype=str, encoding="utf-8-sig").fillna("")
    df.columns = [str(col).strip() for col in df.columns]

    required = ["Name", "Mailadresse", "Mitgliedsstatus"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing expected columns: {missing}. Available: {list(df.columns)}")

    df["Name"] = df["Name"].astype(str).str.strip()
    df["Mailadresse"] = df["Mailadresse"].astype(str).str.strip()
    df["Mitgliedsstatus"] = df["Mitgliedsstatus"].astype(str).str.strip()

    df["status_key"] = df["Mitgliedsstatus"].apply(normalize_status)
    df["status_label"] = df["Mitgliedsstatus"].apply(normalize_status_label)
    df = df[df["status_key"].notna()].copy()
    df = df[df["Mailadresse"].str.contains("@", na=False)].copy()

    first_last = df["Name"].apply(split_name)
    df["first_name"] = first_last.apply(lambda x: x[0])
    df["last_name"] = first_last.apply(lambda x: x[1])

    return df


def to_contact_rows(df: pd.DataFrame) -> dict[str, list[ContactRow]]:
    grouped: dict[str, list[ContactRow]] = {}
    for _, row in df.iterrows():
        key = row["status_key"]
        original_label = str(row["status_label"]).strip()
        if key in MERGED_ACTIVE_KEYS:
            target_label = MERGED_ACTIVE_LABEL
        else:
            target_label = original_label

        grouped.setdefault(target_label, []).append(
            ContactRow(
                first_name=str(row["first_name"]).strip(),
                last_name=str(row["last_name"]).strip(),
                email=str(row["Mailadresse"]).strip(),
                status_key=key,
            )
        )

    ordered_labels = sorted(grouped.keys(), key=lambda label: (0 if label == MERGED_ACTIVE_LABEL else 1, label.lower()))
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
                "Notes": f"Vereinsflieger Verteiler: {label}",
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
                f.write(f"NOTE:{vcard_escape('Vereinsflieger Verteiler: ' + label)}\n")
                f.write("END:VCARD\n")

            summary[label] = {"csv": csv_path, "vcf": vcf_path, "count": len(contacts)}

    return summary


def run_osascript(script: str, args: list[str]) -> str:
    result = subprocess.run(["osascript", "-e", script, *args], capture_output=True, text=True)
    if result.returncode != 0:
        stderr = (result.stderr or "").strip()
        raise RuntimeError(stderr or "AppleScript call failed")
    return (result.stdout or "").strip()


def ensure_groups_do_not_exist(group_names: list[str]) -> None:
    check_script = """
    on run argv
      set existingNames to {}
      tell application \"Contacts\"
        repeat with groupName in argv
          if (exists group (contents of groupName)) then
            set end of existingNames to (contents of groupName)
          end if
        end repeat
      end tell
      if (count of existingNames) is 0 then
        return \"\"
      end if
      set AppleScript's text item delimiters to \"||\"
      set joined to existingNames as text
      set AppleScript's text item delimiters to \"\"
      return joined
    end run
    """
    output = run_osascript(check_script, group_names)
    if output:
        existing = [name for name in output.split("||") if name]
        raise RuntimeError(
            "Refusing to touch existing Contacts lists. Already present: " + ", ".join(existing)
        )


def create_contacts_group(group_name: str, contacts: list[ContactRow]) -> None:
    if not contacts:
        return

    with NamedTemporaryFile("w", encoding="utf-8", suffix=".tsv", delete=False) as tmp:
        tmp_path = Path(tmp.name)
        for c in contacts:
            first = c.first_name.replace("\t", " ").replace("\n", " ").strip()
            last = c.last_name.replace("\t", " ").replace("\n", " ").strip()
            email = c.email.replace("\t", " ").replace("\n", " ").strip()
            tmp.write(f"{first}\t{last}\t{email}\n")

    create_script = """
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
                    set txt to read fileRef
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

            if (count of fields) is greater than or equal to 3 then
              set firstName to item 1 of fields
              set lastName to item 2 of fields
              set emailValue to item 3 of fields

              set newPerson to make new person with properties {first name:firstName, last name:lastName}
              make new email at end of emails of newPerson with properties {label:\"work\", value:emailValue}
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
    group_names = [f"{prefix} {label} {run_stamp}".strip() for label in non_empty_labels]

    ensure_groups_do_not_exist(group_names)

    created: list[str] = []
    for label, group_name in zip(non_empty_labels, group_names):
        contacts = grouped[label]
        create_contacts_group(group_name, contacts)
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
        action="store_true",
        help="Create new lists and contacts in Apple Contacts (safe mode: fails if list names exist).",
    )
    parser.add_argument(
        "--contacts-prefix",
        default="VF Import",
        help="Prefix for new Apple Contacts list names.",
    )
    args = parser.parse_args()

    input_csv = find_latest_csv(args.input_csv)
    if args.output_dir:
        output_dir = Path(args.output_dir)
        if not output_dir.is_absolute():
            output_dir = APP_ROOT / output_dir
    else:
        output_dir = UP_DIR / "contacts_exports" / datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    df = load_contacts_dataframe(input_csv)
    grouped = to_contact_rows(df)
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
        print("Apple Contacts was not modified. Use --apply-contacts to create new lists.")


if __name__ == "__main__":
    main()
