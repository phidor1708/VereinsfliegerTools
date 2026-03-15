"""
tests/test_exporter.py

Tests for src/exporter.py
"""

from pathlib import Path
import re

import pytest

from src.exporter import export_vcards, export_applescript, _member_to_vcard, _applescript_string


# ─── sample data ─────────────────────────────────────────────────────────────

MEMBERS = [
    {"first_name": "Max", "last_name": "Mustermann", "email": "max@example.com", "raw": {}},
    {"first_name": "Erika", "last_name": "Musterfrau", "email": "erika@example.com", "raw": {}},
]

MEMBER_SPECIAL = {
    "first_name": 'O"Brien',
    "last_name": "O'Connor",
    "email": "obrien@example.com",
    "raw": {},
}


# ─── _member_to_vcard ────────────────────────────────────────────────────────

def test_vcard_contains_required_fields():
    card = _member_to_vcard(MEMBERS[0])
    assert "BEGIN:VCARD" in card
    assert "END:VCARD" in card
    assert "VERSION:3.0" in card
    assert "max@example.com" in card
    assert "Mustermann" in card
    assert "Max" in card


def test_vcard_uid_is_uuid():
    card = _member_to_vcard(MEMBERS[0])
    uid_line = next(l for l in card.splitlines() if l.startswith("UID:"))
    uid_value = uid_line[4:]
    # Should be a valid UUID (all uppercase hex + dashes)
    assert re.match(r"[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}", uid_value)


def test_vcard_escapes_semicolons_in_name():
    member = {"first_name": "Hans;Klaus", "last_name": "Müller", "email": "hk@example.com", "raw": {}}
    card = _member_to_vcard(member)
    # Semicolons in name values must be escaped
    assert "Hans\\;Klaus" in card


# ─── export_vcards ───────────────────────────────────────────────────────────

def test_export_vcards_creates_file(tmp_path):
    out = tmp_path / "members.vcf"
    result = export_vcards(MEMBERS, out)
    assert result.exists()


def test_export_vcards_contains_all_members(tmp_path):
    out = tmp_path / "members.vcf"
    export_vcards(MEMBERS, out)
    content = out.read_text(encoding="utf-8")
    assert content.count("BEGIN:VCARD") == 2
    assert "max@example.com" in content
    assert "erika@example.com" in content


def test_export_vcards_utf8_encoding(tmp_path):
    members = [{"first_name": "Jörg", "last_name": "Müller", "email": "joerg@example.com", "raw": {}}]
    out = tmp_path / "umlauts.vcf"
    export_vcards(members, out)
    content = out.read_bytes().decode("utf-8")
    assert "Jörg" in content
    assert "Müller" in content


def test_export_vcards_creates_parent_dir(tmp_path):
    out = tmp_path / "sub" / "dir" / "members.vcf"
    export_vcards(MEMBERS, out)
    assert out.exists()


# ─── _applescript_string ─────────────────────────────────────────────────────

def test_applescript_string_wraps_in_quotes():
    assert _applescript_string("hello") == '"hello"'


def test_applescript_string_escapes_double_quotes():
    assert _applescript_string('say "hi"') == '"say \\"hi\\""'


def test_applescript_string_escapes_backslash():
    assert _applescript_string("C:\\path") == '"C:\\\\path"'


# ─── export_applescript ──────────────────────────────────────────────────────

def test_export_applescript_creates_file(tmp_path):
    out = tmp_path / "update.applescript"
    result = export_applescript(MEMBERS, "TestGroup", out)
    assert result.exists()


def test_export_applescript_contains_group_name(tmp_path):
    out = tmp_path / "update.applescript"
    export_applescript(MEMBERS, "Mein Verein", out)
    content = out.read_text(encoding="utf-8")
    assert "Mein Verein" in content


def test_export_applescript_contains_all_emails(tmp_path):
    out = tmp_path / "update.applescript"
    export_applescript(MEMBERS, "TestGroup", out)
    content = out.read_text(encoding="utf-8")
    assert "max@example.com" in content
    assert "erika@example.com" in content


def test_export_applescript_member_count_in_dialog(tmp_path):
    out = tmp_path / "update.applescript"
    export_applescript(MEMBERS, "TestGroup", out)
    content = out.read_text(encoding="utf-8")
    assert str(len(MEMBERS)) in content


def test_export_applescript_tell_contacts_block(tmp_path):
    out = tmp_path / "update.applescript"
    export_applescript(MEMBERS, "TestGroup", out)
    content = out.read_text(encoding="utf-8")
    assert 'tell application "Contacts"' in content
    assert "end tell" in content
    assert "save" in content
