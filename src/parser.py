"""
src/parser.py

Parses a CSV member list exported from vereinsflieger.de and returns a
list of member dicts that contain at least a first name, last name and
e-mail address.

Vereinsflieger.de CSV files use semicolons as separators and may be
encoded in Latin-1 or UTF-8.  The parser tries both encodings.
"""

from __future__ import annotations

import csv
import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

# Candidate column names for each field (lowercase, stripped).
# The first match in the CSV header wins.
_EMAIL_CANDIDATES = [
    "e-mail",
    "email",
    "e_mail",
    "mail",
    "emailadresse",
    "e-mailadresse",
]
_FIRSTNAME_CANDIDATES = [
    "vorname",
    "firstname",
    "first name",
    "first_name",
    "givenname",
]
_LASTNAME_CANDIDATES = [
    "nachname",
    "name",
    "lastname",
    "last name",
    "last_name",
    "familienname",
    "surname",
]

# Some vereinsflieger exports have a combined "name" column formatted as
# "Nachname, Vorname".  This flag is set when only such a combined column
# is found.
_COMBINED_NAME_CANDIDATES = ["name", "mitglied", "mitgliedername"]


def _detect_column(header: list[str], candidates: list[str]) -> str | None:
    """Return the first header field that matches one of *candidates*."""
    header_lower = [h.lower().strip() for h in header]
    for candidate in candidates:
        if candidate in header_lower:
            return header[header_lower.index(candidate)]
    return None


def _read_csv(path: Path) -> tuple[list[dict[str, str]], str]:
    """Try to read a CSV file, attempting UTF-8 then Latin-1 encoding."""
    for encoding in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            with path.open(newline="", encoding=encoding) as fh:
                # Vereinsflieger uses ";" as delimiter
                sample = fh.read(4096)
                fh.seek(0)
                delimiter = ";" if sample.count(";") > sample.count(",") else ","
                reader = csv.DictReader(fh, delimiter=delimiter)
                rows = list(reader)
                if rows:
                    logger.debug(
                        "CSV read with encoding=%s, delimiter='%s', rows=%d",
                        encoding,
                        delimiter,
                        len(rows),
                    )
                    return rows, encoding
        except (UnicodeDecodeError, csv.Error):
            continue
    return [], "utf-8"


def parse_member_csv(csv_path: Path | str) -> list[dict[str, Any]]:
    """Parse a vereinsflieger.de member CSV and return a list of members.

    Each member dict contains:

    * ``first_name`` – given name
    * ``last_name``  – family name
    * ``email``      – e-mail address (always present; rows without an
                       e-mail address are skipped)
    * ``raw``        – the original CSV row dict for access to extra fields

    Parameters
    ----------
    csv_path:
        Path to the CSV file exported from vereinsflieger.de.

    Returns
    -------
    list[dict]
        Members that have at least one e-mail address.
    """
    csv_path = Path(csv_path)
    rows, _ = _read_csv(csv_path)

    if not rows:
        logger.warning("CSV file is empty or could not be read: %s", csv_path)
        return []

    header = list(rows[0].keys())
    logger.debug("CSV header: %s", header)

    email_col = _detect_column(header, _EMAIL_CANDIDATES)
    first_col = _detect_column(header, _FIRSTNAME_CANDIDATES)
    last_col = _detect_column(header, _LASTNAME_CANDIDATES)
    combined_col = _detect_column(header, _COMBINED_NAME_CANDIDATES)

    if email_col is None:
        logger.error(
            "Could not find an e-mail column in the CSV.  "
            "Available columns: %s",
            header,
        )
        return []

    members: list[dict[str, Any]] = []
    for row in rows:
        email = row.get(email_col, "").strip()
        if not email or "@" not in email:
            continue  # skip rows without a valid e-mail address

        # Resolve first/last name
        if first_col and last_col:
            first_name = row.get(first_col, "").strip()
            last_name = row.get(last_col, "").strip()
        elif combined_col:
            combined = row.get(combined_col, "").strip()
            if "," in combined:
                last_name, _, first_name = combined.partition(",")
                last_name = last_name.strip()
                first_name = first_name.strip()
            else:
                parts = combined.split()
                first_name = parts[0] if parts else ""
                last_name = " ".join(parts[1:]) if len(parts) > 1 else ""
        else:
            first_name = ""
            last_name = ""

        members.append(
            {
                "first_name": first_name,
                "last_name": last_name,
                "email": email,
                "raw": row,
            }
        )

    return members
