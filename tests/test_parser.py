"""
tests/test_parser.py

Tests for src/parser.py
"""

import csv
import io
from pathlib import Path

import pytest

from src.parser import parse_member_csv, _detect_column


# ─── helpers ─────────────────────────────────────────────────────────────────

def _write_csv(tmp_path: Path, rows: list[dict], delimiter: str = ";") -> Path:
    """Write *rows* as a CSV file and return the path."""
    p = tmp_path / "members.csv"
    fieldnames = list(rows[0].keys()) if rows else []
    with p.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, delimiter=delimiter)
        writer.writeheader()
        writer.writerows(rows)
    return p


# ─── _detect_column ──────────────────────────────────────────────────────────

def test_detect_column_exact_match():
    header = ["Vorname", "Nachname", "E-Mail"]
    assert _detect_column(header, ["e-mail"]) == "E-Mail"


def test_detect_column_case_insensitive():
    header = ["vorname", "NACHNAME", "EMAIL"]
    assert _detect_column(header, ["email"]) == "EMAIL"


def test_detect_column_first_candidate_wins():
    header = ["email", "mail"]
    assert _detect_column(header, ["email", "mail"]) == "email"


def test_detect_column_no_match():
    header = ["foo", "bar"]
    assert _detect_column(header, ["email", "e-mail"]) is None


# ─── parse_member_csv ────────────────────────────────────────────────────────

def test_parse_standard_columns(tmp_path):
    rows = [
        {"Vorname": "Max", "Nachname": "Mustermann", "E-Mail": "max@example.com"},
        {"Vorname": "Erika", "Nachname": "Musterfrau", "E-Mail": "erika@example.com"},
    ]
    csv_file = _write_csv(tmp_path, rows)
    members = parse_member_csv(csv_file)
    assert len(members) == 2
    assert members[0]["first_name"] == "Max"
    assert members[0]["last_name"] == "Mustermann"
    assert members[0]["email"] == "max@example.com"


def test_parse_skips_rows_without_email(tmp_path):
    rows = [
        {"Vorname": "Max", "Nachname": "Mustermann", "E-Mail": "max@example.com"},
        {"Vorname": "Ghost", "Nachname": "Member", "E-Mail": ""},
        {"Vorname": "Invalid", "Nachname": "Email", "E-Mail": "not-an-email"},
    ]
    csv_file = _write_csv(tmp_path, rows)
    members = parse_member_csv(csv_file)
    assert len(members) == 1
    assert members[0]["first_name"] == "Max"


def test_parse_combined_name_column_comma_separated(tmp_path):
    rows = [
        {"Name": "Mustermann, Max", "E-Mail": "max@example.com"},
    ]
    csv_file = _write_csv(tmp_path, rows)
    members = parse_member_csv(csv_file)
    assert len(members) == 1
    assert members[0]["first_name"] == "Max"
    assert members[0]["last_name"] == "Mustermann"


def test_parse_combined_name_column_space_separated(tmp_path):
    rows = [
        {"Name": "Max Mustermann", "E-Mail": "max@example.com"},
    ]
    csv_file = _write_csv(tmp_path, rows)
    members = parse_member_csv(csv_file)
    assert len(members) == 1
    assert members[0]["first_name"] == "Max"
    assert members[0]["last_name"] == "Mustermann"


def test_parse_comma_delimiter(tmp_path):
    """CSV files with comma delimiter should also be parsed correctly."""
    rows = [
        {"Vorname": "Max", "Nachname": "Mustermann", "E-Mail": "max@example.com"},
    ]
    csv_file = _write_csv(tmp_path, rows, delimiter=",")
    members = parse_member_csv(csv_file)
    assert len(members) == 1


def test_parse_raw_field_preserved(tmp_path):
    rows = [
        {"Vorname": "Max", "Nachname": "Mustermann", "E-Mail": "max@example.com", "Ort": "Höchstadt"},
    ]
    csv_file = _write_csv(tmp_path, rows)
    members = parse_member_csv(csv_file)
    assert members[0]["raw"]["Ort"] == "Höchstadt"


def test_parse_empty_csv(tmp_path):
    csv_file = tmp_path / "empty.csv"
    csv_file.write_text("", encoding="utf-8")
    members = parse_member_csv(csv_file)
    assert members == []


def test_parse_no_email_column(tmp_path):
    rows = [{"Vorname": "Max", "Nachname": "Mustermann"}]
    csv_file = _write_csv(tmp_path, rows)
    members = parse_member_csv(csv_file)
    assert members == []


def test_parse_latin1_encoded(tmp_path):
    """CSV files encoded in Latin-1 (common from older Windows systems) are read correctly."""
    csv_file = tmp_path / "latin1.csv"
    content = "Vorname;Nachname;E-Mail\nJörg;Müller;joerg@example.com\n"
    csv_file.write_bytes(content.encode("latin-1"))
    members = parse_member_csv(csv_file)
    assert len(members) == 1
    assert members[0]["first_name"] == "Jörg"
    assert members[0]["last_name"] == "Müller"
