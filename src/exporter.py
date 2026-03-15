"""
src/exporter.py

Exports a list of member dicts (as produced by src/parser.py) into two
formats suitable for Apple Mail on macOS:

1. A vCard (.vcf) file that can be imported into Apple Contacts.
2. An AppleScript (.applescript) that creates / updates a contact group
   in Apple Contacts.  The group is then available as a distribution
   list when composing a new message in Apple Mail.
"""

from __future__ import annotations

import logging
import uuid
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# vCard export
# ─────────────────────────────────────────────────────────────────────────────

def _escape_vcard(value: str) -> str:
    """Escape special characters for vCard property values."""
    return (
        value.replace("\\", "\\\\")
        .replace(",", "\\,")
        .replace(";", "\\;")
        .replace("\n", "\\n")
    )


def _member_to_vcard(member: dict[str, Any]) -> str:
    """Convert a single member dict to a vCard 3.0 string."""
    first = _escape_vcard(member.get("first_name", ""))
    last = _escape_vcard(member.get("last_name", ""))
    email = member.get("email", "").strip()
    uid = str(uuid.uuid4()).upper()

    lines = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        f"UID:{uid}",
        f"N:{last};{first};;;",
        f"FN:{first} {last}".strip(),
        f"EMAIL;TYPE=INTERNET:{email}",
        "END:VCARD",
    ]
    return "\r\n".join(lines) + "\r\n"


def export_vcards(
    members: list[dict[str, Any]],
    output_path: Path | str,
) -> Path:
    """Write all members as vCards to a single .vcf file.

    The file can be imported into Apple Contacts by double-clicking it or
    via *File > Import* in the Contacts app.

    Parameters
    ----------
    members:
        List of member dicts produced by :func:`src.parser.parse_member_csv`.
    output_path:
        Destination file path (should end with ``.vcf``).

    Returns
    -------
    Path
        Absolute path of the written file.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    content = "".join(_member_to_vcard(m) for m in members)
    output_path.write_text(content, encoding="utf-8")
    logger.debug("Wrote %d vCards to %s", len(members), output_path)
    return output_path.resolve()


# ─────────────────────────────────────────────────────────────────────────────
# AppleScript export
# ─────────────────────────────────────────────────────────────────────────────

def _applescript_string(value: str) -> str:
    """Wrap a Python string as a properly escaped AppleScript string literal."""
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def export_applescript(
    members: list[dict[str, Any]],
    group_name: str,
    output_path: Path | str,
) -> Path:
    """Generate an AppleScript that creates / updates an Apple Contacts group.

    The script:

    * Creates the group *group_name* if it does not exist yet.
    * For every member: creates a new contact (or finds an existing one by
      e-mail) and adds it to the group.
    * Saves the Contacts database at the end.

    Run the script on your Mac with::

        osascript output/update_contacts_group.applescript

    Parameters
    ----------
    members:
        List of member dicts.
    group_name:
        Name for the Apple Contacts group / Mac Mail distribution list.
    output_path:
        Destination file path (should end with ``.applescript``).

    Returns
    -------
    Path
        Absolute path of the written file.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    group_name_as = _applescript_string(group_name)

    lines: list[str] = [
        "-- VereinsfliegerTools: update Apple Contacts group",
        "-- Generated automatically – run with: osascript update_contacts_group.applescript",
        "",
        'tell application "Contacts"',
        "    -- Create the group if it does not exist yet",
        f"    set groupName to {group_name_as}",
        "    set matchingGroups to (every group whose name is groupName)",
        "    if (count of matchingGroups) > 0 then",
        "        set theGroup to item 1 of matchingGroups",
        "    else",
        "        set theGroup to make new group with properties {name:groupName}",
        "    end if",
        "",
    ]

    for m in members:
        first = _applescript_string(m.get("first_name", ""))
        last = _applescript_string(m.get("last_name", ""))
        email = _applescript_string(m.get("email", ""))
        lines += [
            f"    -- Member: {m.get('first_name', '')} {m.get('last_name', '')} <{m.get('email', '')}>",
            f"    set memberEmail to {email}",
            f"    set memberFirst to {first}",
            f"    set memberLast to {last}",
            "    -- Find existing contact by e-mail or create a new one",
            "    set matchingPeople to (every person whose value of emails contains memberEmail)",
            "    if (count of matchingPeople) > 0 then",
            "        set thePerson to item 1 of matchingPeople",
            "    else",
            "        set thePerson to make new person with properties {first name:memberFirst, last name:memberLast}",
            '        make new email at end of emails of thePerson with properties {label:"work", value:memberEmail}',
            "    end if",
            "    add thePerson to theGroup",
            "",
        ]

    lines += [
        "    save",
        "end tell",
        "",
        f'display dialog "Gruppe " & {group_name_as} & " wurde mit {len(members)} Mitglied(ern) aktualisiert." buttons {{"OK"}} default button "OK"',
        "",
    ]

    output_path.write_text("\n".join(lines), encoding="utf-8")
    logger.debug("Wrote AppleScript to %s", output_path)
    return output_path.resolve()
