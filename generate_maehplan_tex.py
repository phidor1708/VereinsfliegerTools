#!/usr/bin/env python3
"""Generate a LaTeX Mähplan from up/MaehplanStatic.json.

Output: up/{year}-MFG-Maehplan.tex
Compile: pdflatex -output-directory up up/{year}-MFG-Maehplan.tex
"""
from __future__ import annotations

import csv
import json
import re
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent
UP_DIR = APP_ROOT / "up"
JSON_PATH = UP_DIR / "MaehplanStatic.json"
INVISIBLE_CONTROL_RE = re.compile(r"[\u200e\u200f\u202a-\u202e\u2066-\u2069]")
DEFAULT_SPRING_CLEANING_NOTE = "Alle: Frühjahrsputz am Modellfluggelände von 9:00 - 12:00"


def normalize_key(name: str) -> str:
    value = (name or "").strip().lower()
    value = value.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def clean_text(value: str) -> str:
    return INVISIBLE_CONTROL_RE.sub("", (value or "")).strip()


def format_phone(raw_phone: str) -> str:
    phone = clean_text(raw_phone)
    if not phone:
        return ""

    phone = re.sub(r"[^\d+]", "", phone)
    if phone.startswith("00"):
        phone = "+" + phone[2:]

    if phone.count("+") > 1:
        phone = "+" + phone.replace("+", "")
    elif "+" in phone and not phone.startswith("+"):
        phone = "+" + phone.replace("+", "")

    return phone


_TEX_TABLE = str.maketrans(
    {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
)


def tex(value: str) -> str:
    return clean_text(value).translate(_TEX_TABLE)


def load_phone_lookup() -> dict[str, str]:
    candidates = sorted(UP_DIR.glob("mitglieder_*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        candidates = sorted(UP_DIR.glob("*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        return {}

    csv_path = candidates[0]
    lookup: dict[str, str] = {}

    try:
        with csv_path.open(encoding="utf-8-sig", newline="") as fh:
            sample = fh.read(4096)
        delimiter = ";"
        try:
            delimiter = csv.Sniffer().sniff(sample, delimiters=";,\t").delimiter
        except Exception:
            pass

        with csv_path.open(encoding="utf-8-sig", newline="") as fh:
            reader = csv.DictReader(fh, delimiter=delimiter)
            for row in reader:
                raw_name = clean_text(row.get("Name") or "")
                first_name = clean_text(row.get("Vorname") or "")
                last_name = clean_text(row.get("Nachname") or "")

                if not first_name and not last_name:
                    if "," in raw_name:
                        last_name, first_name = raw_name.split(",", 1)
                        last_name, first_name = last_name.strip(), first_name.strip()
                    else:
                        parts = raw_name.split()
                        first_name = " ".join(parts[:-1]) if len(parts) > 1 else ""
                        last_name = parts[-1] if parts else ""

                phone = (
                    clean_text(row.get("Mobil (privat)") or "")
                    or clean_text(row.get("Mobil (geschäftl.)") or "")
                    or clean_text(row.get("Telefon (privat)") or "")
                    or clean_text(row.get("Telefon (geschäftl.)") or "")
                )
                phone = format_phone(phone)

                keys: list[str] = []
                if last_name:
                    keys.append(normalize_key(f"{last_name} {first_name}"))
                if first_name:
                    keys.append(normalize_key(f"{first_name} {last_name}"))
                if raw_name:
                    keys.append(normalize_key(raw_name))

                for key in keys:
                    if key:
                        lookup[key] = phone
    except Exception:
        pass

    return lookup


TEMPLATE = r"""\documentclass[10pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[ngerman]{babel}
\usepackage[a4paper,landscape,top=1.0cm,bottom=1.0cm,left=1.0cm,right=1.0cm]{geometry}
\usepackage{tabularx}
\usepackage{array}
\usepackage[table]{xcolor}
\usepackage{helvet}
\renewcommand{\familydefault}{\sfdefault}
\setlength{\parindent}{0pt}
\setlength{\tabcolsep}{5pt}
\setlength{\extrarowheight}{0.8pt}
\renewcommand{\arraystretch}{1.34}
\newcolumntype{T}{>{\small\raggedright\arraybackslash}m{3.2cm}}

\definecolor{mfgblue}{rgb}{0.13,0.33,0.58}
\definecolor{headerbg}{RGB}{226,237,252}

\begin{document}
\pagestyle{empty}

% ── Page 1 Header ─────────────────────────────────────────────────────────────
\noindent{\color{mfgblue}\rule{\textwidth}{1.4pt}}
{\centering
    {\normalsize\bfseries M\"ahplan <<YEAR>>\par}
    {\normalsize\bfseries\color{mfgblue} Modellfliegergruppe H\"ochstadt\,/\,Aisch\,e.\,V.\par}
    {\small Meldung zum freiwilligen M\"ahdienst\par}
}
\vspace{0.05em}
\noindent{\color{mfgblue}\rule{\textwidth}{0.8pt}}

\vspace{0.16em}

% ── Schedule table page 1 (Termin 0..14) ─────────────────────────────────────
\noindent
\begin{tabularx}{\textwidth}{|>{\small}l|>{\small}l|T|>{\small}l|T|>{\small\raggedright\arraybackslash}X|}
  \hline
  \rowcolor{headerbg}{\bfseries Datum} & {\bfseries M\"aher 1} & {\bfseries Tel. Nr. M\"aher 1} & {\bfseries M\"aher 2} & {\bfseries Tel. Nr. M\"aher 2} & {\bfseries Anmerkungen} \\
  \hline
<<ROWS_PAGE_1>>
\end{tabularx}

\vfill
{\footnotesize\color{gray}\ensuremath{\bullet} Der mit dem Punkt gekennzeichnete M\"aher ist f\"ur die Durchf\"uhrung verantwortlich.}\par
{\footnotesize\color{gray}Stand M\"ahplan: <<PLAN_STATUS>>\hfill Seite 1 von 2}

\newpage

% ── Page 2 Header ─────────────────────────────────────────────────────────────
\noindent{\color{mfgblue}\rule{\textwidth}{1.4pt}}
{\centering
    {\normalsize\bfseries M\"ahplan <<YEAR>>\par}
    {\normalsize\bfseries\color{mfgblue} Modellfliegergruppe H\"ochstadt\,/\,Aisch\,e.\,V.\par}
    {\small Meldung zum freiwilligen M\"ahdienst\par}
}
\vspace{0.05em}
\noindent{\color{mfgblue}\rule{\textwidth}{0.8pt}}

\vspace{0.16em}

% ── Schedule table page 2 (Termin 15..Ende) ──────────────────────────────────
\noindent
\begin{tabularx}{\textwidth}{|>{\small}l|>{\small}l|T|>{\small}l|T|>{\small\raggedright\arraybackslash}X|}
  \hline
  \rowcolor{headerbg}{\bfseries Datum} & {\bfseries M\"aher 1} & {\bfseries Tel. Nr. M\"aher 1} & {\bfseries M\"aher 2} & {\bfseries Tel. Nr. M\"aher 2} & {\bfseries Anmerkungen} \\
  \hline
<<ROWS_PAGE_2>>
\end{tabularx}

\vfill
{\footnotesize\color{gray}\ensuremath{\bullet} Der mit dem Punkt gekennzeichnete M\"aher ist f\"ur die Durchf\"uhrung verantwortlich.}\par
{\footnotesize\color{gray}Stand M\"ahplan: <<PLAN_STATUS>>\hfill Seite 2 von 2}

\end{document}
"""


def build_rows(
    schedule: list[tuple[int, date, str, str, str, str, str, int]],
    include_spring_row: bool,
    spring_date: date,
    spring_note: str,
    row_height_ex: float,
) -> str:
    lines: list[str] = []

    if include_spring_row:
        frpz_date = spring_date.strftime("%d.%m")
        lines.append(
            f"  \\rule{{0pt}}{{{row_height_ex:.2f}ex}}{frpz_date} & \\multicolumn{{5}}{{c|}}{{\\textbf{{{tex(spring_note)}}}}} \\\\" 
        )
        lines.append("  \\hline")

    for _, current_date, p1, tel1, p2, tel2, remark, responsible_idx in schedule:
        date_str = current_date.strftime("%d.%m")
        p1_cell = tex(p1)
        p2_cell = tex(p2)
        if responsible_idx == 1:
            p1_cell = f"{p1_cell} \\ensuremath{{\\bullet}}"
        else:
            p2_cell = f"{p2_cell} \\ensuremath{{\\bullet}}"
        lines.append(
            f"  \\rule{{0pt}}{{{row_height_ex:.2f}ex}}{date_str} & {p1_cell} & {tex(tel1)} & {p2_cell} & {tex(tel2)} & {tex(remark)} \\\\"
        )
        lines.append("  \\hline")

    return "\n".join(lines)


def compile_latex(tex_path: Path) -> bool:
    command = [
        "pdflatex",
        "-interaction=nonstopmode",
        "-halt-on-error",
        "-output-directory",
        str(UP_DIR),
        str(tex_path),
    ]
    try:
        result = subprocess.run(command, cwd=APP_ROOT, check=False)
    except FileNotFoundError:
        print("Warning: pdflatex not found. Skipped PDF compilation.", file=sys.stderr)
        return False

    if result.returncode != 0:
        print("Warning: pdflatex failed. Please check the LaTeX log in up/.", file=sys.stderr)
        return False

    return True


def main() -> None:
    if not JSON_PATH.exists():
        print(f"Error: {JSON_PATH} not found. Run create_contacts_lists.py first.", file=sys.stderr)
        sys.exit(1)

    raw: dict = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    term_0 = date.fromisoformat(raw["Mähtermin_0"])
    spring_note = str(raw.get("FrühjahrsputzInfo", DEFAULT_SPRING_CLEANING_NOTE)).strip() or DEFAULT_SPRING_CLEANING_NOTE
    year = term_0.year
    phones = load_phone_lookup()
    plan_status = datetime.fromtimestamp(JSON_PATH.stat().st_mtime).strftime("%d.%m.%Y %H:%M")

    member_data = {
        key: value
        for key, value in raw.items()
        if isinstance(value, dict) and "Mähpartner" in value and "MähterminNr" in value
    }

    display: dict[str, str] = {}
    for key, info in member_data.items():
        partner_key = normalize_key(info["Mähpartner"])
        if partner_key in member_data:
            display[key] = member_data[partner_key]["Mähpartner"]

    seen: set[frozenset[int]] = set()
    schedule: list[tuple[int, date, str, str, str, str, str, int]] = []

    for key, info in member_data.items():
        nrs = [int(x.strip()) for x in info["MähterminNr"].split(",") if x.strip().isdigit()]
        marker = frozenset(nrs)
        if marker in seen:
            continue
        seen.add(marker)

        p1_display = display.get(key, key.title())
        partner_key = normalize_key(info["Mähpartner"])
        p2_display = display.get(partner_key, info["Mähpartner"])

        tel1 = phones.get(normalize_key(p1_display), "")
        tel2 = phones.get(normalize_key(p2_display), "")
        remark = str(info.get("Anmerkung", "")).strip()

        for idx, nr in enumerate(sorted(nrs)):
            current_date = term_0 + timedelta(weeks=nr)
            responsible_idx = 1 if idx % 2 == 0 else 2
            schedule.append((nr, current_date, p1_display, tel1, p2_display, tel2, remark, responsible_idx))

    schedule.sort(key=lambda item: item[0])

    missing_phone_names = sorted(
        {p1 for _, _, p1, tel1, _, _, _, _ in schedule if not tel1}
        | {p2 for _, _, _, _, p2, tel2, _, _ in schedule if not tel2}
    )
    if missing_phone_names:
        print("Warning: Missing phone number for the following Mäher:")
        for name in missing_phone_names:
            print(f"- {name}")

    page_1_schedule = [item for item in schedule if item[0] <= 14]
    page_2_schedule = [item for item in schedule if item[0] > 14]

    def compute_row_height_ex(row_count: int) -> float:
        if row_count <= 0:
            return 3.2
        return max(3.1, min(5.0, 58.0 / row_count))

    page_1_row_count = len(page_1_schedule) + 1
    page_2_row_count = len(page_2_schedule)
    page_1_row_height_ex = compute_row_height_ex(page_1_row_count)
    page_2_row_height_ex = compute_row_height_ex(page_2_row_count)

    rows_page_1 = build_rows(
        schedule=page_1_schedule,
        include_spring_row=True,
        spring_date=term_0,
        spring_note=spring_note,
        row_height_ex=page_1_row_height_ex,
    )
    rows_page_2 = build_rows(
        schedule=page_2_schedule,
        include_spring_row=False,
        spring_date=term_0,
        spring_note=spring_note,
        row_height_ex=page_2_row_height_ex,
    )

    result = (
        TEMPLATE
        .replace("<<YEAR>>", str(year))
        .replace("<<ROWS_PAGE_1>>", rows_page_1)
        .replace("<<ROWS_PAGE_2>>", rows_page_2)
        .replace("<<PLAN_STATUS>>", plan_status)
    )

    out_path = UP_DIR / f"{year}-MFG-Maehplan.tex"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(result, encoding="utf-8")
    print(f"Generated:  {out_path}")
    if compile_latex(out_path):
        print(f"Compiled:   {UP_DIR / f'{year}-MFG-Maehplan.pdf'}")
    else:
        print(f"To compile manually: pdflatex -interaction=nonstopmode -halt-on-error -output-directory up {out_path.name}")


if __name__ == "__main__":
    main()
