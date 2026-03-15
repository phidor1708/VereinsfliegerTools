"""
VereinsfliegerTools – main entry point.

Usage:
    python main.py [--config config.yaml] [--csv path/to/members.csv]

If --csv is provided the download step is skipped and the given file is
parsed directly (useful for testing without internet access).
"""

import argparse
import logging
import sys
from pathlib import Path

import yaml

from src.downloader import download_member_csv
from src.parser import parse_member_csv
from src.exporter import export_vcards, export_applescript

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def load_config(config_path: str) -> dict:
    path = Path(config_path)
    if not path.exists():
        logger.error(
            "Config file not found: %s\n"
            "Copy config.example.yaml to config.yaml and fill in your credentials.",
            config_path,
        )
        sys.exit(1)
    with path.open(encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Download Vereinsflieger member list and create Mac Mail distribution lists."
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to YAML config file (default: config.yaml)",
    )
    parser.add_argument(
        "--csv",
        metavar="FILE",
        help="Skip download and use an existing CSV file instead",
    )
    parser.add_argument(
        "--output-dir",
        metavar="DIR",
        help="Override the output directory from the config",
    )
    args = parser.parse_args()

    config = load_config(args.config)

    output_dir = Path(
        args.output_dir or config.get("export", {}).get("output_dir", "./output")
    )
    output_dir.mkdir(parents=True, exist_ok=True)

    group_name: str = config.get("export", {}).get(
        "group_name", "Modellflugverein"
    )

    # ── Step 1: obtain CSV ────────────────────────────────────────────────
    if args.csv:
        csv_path = Path(args.csv)
        logger.info("Using existing CSV file: %s", csv_path)
    else:
        vf_cfg = config.get("vereinsflieger", {})
        username = vf_cfg.get("username", "")
        password = vf_cfg.get("password", "")
        if not username or not password:
            logger.error(
                "No credentials found in config. "
                "Set vereinsflieger.username and vereinsflieger.password."
            )
            sys.exit(1)
        csv_path = output_dir / "members.csv"
        logger.info("Downloading member list from vereinsflieger.de …")
        download_member_csv(
            username=username,
            password=password,
            club_id=vf_cfg.get("club_id", ""),
            output_path=csv_path,
        )
        logger.info("Member list saved to %s", csv_path)

    # ── Step 2: parse CSV ────────────────────────────────────────────────
    members = parse_member_csv(csv_path)
    logger.info("Parsed %d members with e-mail address.", len(members))

    if not members:
        logger.warning(
            "No members with e-mail addresses found – no output written."
        )
        return

    # ── Step 3: export ────────────────────────────────────────────────────
    vcf_path = output_dir / "members.vcf"
    export_vcards(members, vcf_path)
    logger.info("vCard file written to %s", vcf_path)

    script_path = output_dir / "update_contacts_group.applescript"
    export_applescript(members, group_name=group_name, output_path=script_path)
    logger.info("AppleScript written to %s", script_path)

    print(
        "\n"
        "─────────────────────────────────────────────────────────────────\n"
        "Done! Next steps on your Mac:\n"
        "\n"
        f"  1. Open '{vcf_path}' to import all members into Apple Contacts.\n"
        "     (double-click the file or drag it onto the Contacts app)\n"
        "\n"
        f"  2. Run the AppleScript to create / update the '{group_name}'\n"
        f"     group in Apple Contacts:\n"
        f"       osascript '{script_path}'\n"
        "     The group then appears as a distribution list in Apple Mail.\n"
        "─────────────────────────────────────────────────────────────────\n"
    )


if __name__ == "__main__":
    main()
