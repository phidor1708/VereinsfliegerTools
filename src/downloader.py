"""
src/downloader.py

Logs into vereinsflieger.de using a browser (Playwright) and downloads
the member list as a CSV file.

Vereinsflieger.de does not offer CSV export through its REST API on the
free plan, so we automate the web UI instead.
"""

from __future__ import annotations

import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# ── URL constants (vereinsflieger.de web UI) ──────────────────────────────────
_BASE_URL = "https://www.vereinsflieger.de"
_LOGIN_URL = f"{_BASE_URL}/member/login/index"
# The member-list export URL – vereinsflieger.de sends a CSV when this page is
# requested with the right parameters.
_EXPORT_URL = f"{_BASE_URL}/member/member/export"


def download_member_csv(
    username: str,
    password: str,
    club_id: str = "",
    output_path: Path | str = "output/members.csv",
) -> Path:
    """Log into vereinsflieger.de and download the member list as CSV.

    Parameters
    ----------
    username:
        The e-mail / username used to log in.
    password:
        The account password.
    club_id:
        Optional club ID (Vereinsnummer).  Leave empty when the login
        form does not show a separate club-ID field.
    output_path:
        Where to write the downloaded CSV file.

    Returns
    -------
    Path
        Absolute path to the written CSV file.
    """
    # Import here so the rest of the module can be tested without playwright.
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
    except ImportError as exc:
        raise ImportError(
            "playwright is required for downloading.  "
            "Install it with: pip install playwright && playwright install chromium"
        ) from exc

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # ── 1. Login ──────────────────────────────────────────────────────
        logger.debug("Navigating to login page: %s", _LOGIN_URL)
        page.goto(_LOGIN_URL, wait_until="networkidle")

        # Fill in username
        page.fill('input[name="email"]', username)
        # Fill in password
        page.fill('input[name="password"]', password)

        # Optionally fill club ID if the field is present and a value was given
        if club_id:
            try:
                cid_field = page.locator('input[name="cid"]')
                if cid_field.count() > 0:
                    cid_field.fill(club_id)
            except PWTimeoutError:
                logger.debug("Club-ID field not found or timed out – skipping.")

        # Submit the login form
        logger.debug("Submitting login form …")
        with page.expect_navigation(wait_until="networkidle"):
            page.click('button[type="submit"], input[type="submit"]')

        # Verify we are logged in by checking for a sign of the member area
        if "login" in page.url.lower():
            raise RuntimeError(
                "Login failed – still on login page after submitting credentials.  "
                "Please check your username and password in config.yaml."
            )
        logger.debug("Login successful. Current URL: %s", page.url)

        # ── 2. Download the member list CSV ───────────────────────────────
        logger.debug("Navigating to member export …")
        with page.expect_download() as download_info:
            page.goto(_EXPORT_URL, wait_until="networkidle")

        download = download_info.value
        download.save_as(output_path)
        logger.debug("Download saved to %s", output_path)

        context.close()
        browser.close()

    return output_path.resolve()
