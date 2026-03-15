from __future__ import annotations

import argparse
from datetime import datetime
import os
import re
import signal
import subprocess
import time
from pathlib import Path

from playwright.sync_api import Locator, Page, sync_playwright


APP_ROOT = Path(__file__).resolve().parent


def list_other_script_pids() -> list[int]:
    current_pid = os.getpid()
    output = subprocess.check_output(["ps", "-axo", "pid=,command="], text=True)

    pids: list[int] = []
    for line in output.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split(maxsplit=1)
        if len(parts) != 2:
            continue
        pid_text, command = parts
        if "download_mitglieder_csv.py" not in command:
            continue
        try:
            pid = int(pid_text)
        except ValueError:
            continue
        if pid == current_pid:
            continue
        pids.append(pid)

    return sorted(set(pids))


def process_exists(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def terminate_pid(pid: int, grace_seconds: float = 4.0) -> None:
    if not process_exists(pid):
        return

    os.kill(pid, signal.SIGTERM)
    deadline = time.time() + grace_seconds
    while time.time() < deadline:
        if not process_exists(pid):
            return
        time.sleep(0.2)

    if process_exists(pid):
        os.kill(pid, signal.SIGKILL)


def ensure_no_other_instances(kill_other_runs: bool) -> None:
    other_pids = list_other_script_pids()
    if not other_pids:
        return

    if not kill_other_runs:
        pids = ", ".join(str(pid) for pid in other_pids)
        raise RuntimeError(
            "Another script instance is already running "
            f"(PID(s): {pids}). Use --kill-other-runs to close stale instances."
        )

    for pid in other_pids:
        terminate_pid(pid)
    print(f"Closed previous script instance(s): {', '.join(str(pid) for pid in other_pids)}")


def load_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        values[key] = value
    return values


def first_clickable(locator: Locator, timeout_ms: int = 3000) -> Locator | None:
    try:
        locator.first.wait_for(state="visible", timeout=timeout_ms)
    except Exception:
        pass
    try:
        count = locator.count()
    except Exception:
        return None
    for index in range(count):
        candidate = locator.nth(index)
        try:
            candidate.wait_for(state="visible", timeout=timeout_ms)
            return candidate
        except Exception:
            continue
    return None


def click_by_candidates(page: Page, candidates: list[Locator], timeout_ms: int = 3000) -> bool:
    for locator in candidates:
        target = first_clickable(locator, timeout_ms=timeout_ms)
        if not target:
            continue
        try:
            target.click(timeout=timeout_ms)
            return True
        except Exception:
            continue
    return False


def get_first_input(page: Page, selectors: list[str], timeout_ms: int = 3000) -> Locator | None:
    for selector in selectors:
        locator = page.locator(selector)
        target = first_clickable(locator, timeout_ms=timeout_ms)
        if target:
            return target
    return None


def dismiss_common_popups(page: Page) -> None:
    click_by_candidates(
        page,
        [
            page.get_by_role("button", name=re.compile(r"alle akzeptieren|akzeptieren|zustimmen", re.I)),
            page.locator("button:has-text('Alle akzeptieren')"),
            page.locator("button:has-text('Akzeptieren')"),
        ],
        timeout_ms=1200,
    )


def login(page: Page, email: str, password: str, timeout_ms: int) -> None:
    page.goto("https://www.vereinsflieger.de", wait_until="domcontentloaded", timeout=timeout_ms)
    dismiss_common_popups(page)

    email_input = get_first_input(
        page,
        [
            "input[name='user'][placeholder='Benutzer oder E-Mail']",
            "input[name='user']",
            "input[placeholder='Benutzer oder E-Mail']",
        ],
        timeout_ms=timeout_ms,
    )
    password_input = get_first_input(
        page,
        [
            "input[name='pwinput'][type='password']",
            "input[name='pwinput']",
            "input[id='pwinput']",
            "input[placeholder='Passwort'][type='password']",
        ],
        timeout_ms=timeout_ms,
    )

    if not email_input or not password_input:
        raise RuntimeError("Login fields not found. Open page structure may have changed.")

    email_input.click(timeout=timeout_ms)
    email_input.fill(email)
    password_input.click(timeout=timeout_ms)
    password_input.fill(password)

    logged_in = click_by_candidates(
        page,
        [
            page.locator("button:has-text('Anmelden')"),
            page.get_by_role("button", name=re.compile(r"^anmelden$", re.I)),
            page.locator("input[type='submit'][value='Anmelden']"),
            page.locator("button[type='submit']"),
            page.get_by_role("button", name=re.compile(r"login|anmelden", re.I)),
        ],
        timeout_ms=timeout_ms,
    )

    if not logged_in:
        password_input.press("Enter")

    page.wait_for_load_state("networkidle", timeout=timeout_ms)


def ensure_logged_in(page: Page) -> None:
    if "/member/" in page.url:
        return

    indicators = [
        page.get_by_role("link", name=re.compile(r"Abmelden", re.I)),
        page.get_by_text(re.compile(r"Abmelden", re.I)),
        page.get_by_text(re.compile(r"Übersicht", re.I)),
    ]
    for indicator in indicators:
        try:
            if indicator.first.is_visible(timeout=1200):
                return
        except Exception:
            continue

    raise RuntimeError(
        "Login appears unsuccessful (no member area indicator found). "
        "Automation stopped before further clicks."
    )


def open_club(page: Page, timeout_ms: int) -> None:
    if "/member/community/" in page.url:
        return

    club_opened = click_by_candidates(
        page,
        [
            page.locator("#topnavi a[href='/member/community/main']"),
            page.locator("#topnavi a:has-text('MFG Höchstadt/Aisch')"),
            page.get_by_role("link", name=re.compile(r"^MFG Höchstadt/Aisch$", re.I)),
        ],
        timeout_ms=timeout_ms,
    )
    if not club_opened:
        page.goto("https://www.vereinsflieger.de/member/community/main", wait_until="networkidle", timeout=timeout_ms)

    page.wait_for_url(re.compile(r"/member/community/"), timeout=timeout_ms)


def open_mitglieder_left(page: Page, timeout_ms: int) -> None:
    if "/member/community/users.php" in page.url:
        return

    page.wait_for_timeout(500)

    members_opened = click_by_candidates(
        page,
        [
            page.locator("#leftnavi a[href='/member/community/users.php']"),
            page.locator("#leftnavi a:has-text('Mitglieder')"),
            page.get_by_role("link", name=re.compile(r"^Mitglieder$", re.I)),
        ],
        timeout_ms=timeout_ms,
    )

    if not members_opened:
        page.goto("https://www.vereinsflieger.de/member/community/users.php", wait_until="networkidle", timeout=timeout_ms)

    page.wait_for_url(re.compile(r"/member/community/users\.php"), timeout=timeout_ms)
    page.wait_for_load_state("networkidle", timeout=timeout_ms)


def select_table_quick_config(page: Page, config_label: str, timeout_ms: int) -> None:
    selected = page.evaluate(
        """
        ([wantedLabel]) => {
            const selects = Array.from(document.querySelectorAll("select"));
            const select = selects.find((candidate) => {
                if (candidate.name === "lst_configavailable") {
                    return true;
                }
                if ((candidate.getAttribute("onchange") || "").includes("onChangeTableListQuickConfig")) {
                    return true;
                }
                const containerText = candidate.parentElement?.textContent || "";
                return containerText.includes("Konfiguration laden");
            });
            if (!select) {
                return false;
            }

            const wantedOption = Array.from(select.options).find(
                (option) => !option.disabled && option.textContent?.trim() === wantedLabel
            );
            if (!wantedOption) {
                return false;
            }

            if (select.value === wantedOption.value) {
                return true;
            }

            select.value = wantedOption.value;
            select.dispatchEvent(new Event("change", { bubbles: true }));
            return true;
        }
        """,
        [config_label],
    )
    if not selected:
        raise RuntimeError(f"Could not find quick configuration option {config_label!r} on Mitglieder page.")

    page.wait_for_load_state("networkidle", timeout=timeout_ms)


def stop_after_step(current_step: str, stop_after: str, keep_open_seconds: int) -> bool:
    if current_step != stop_after:
        return False

    print(f"Stopped after step: {current_step}")
    if keep_open_seconds > 0:
        print(f"Keeping browser open for {keep_open_seconds} seconds for debugging...")
        time.sleep(keep_open_seconds)
    return True


def clear_table_filters(page: Page, timeout_ms: int) -> None:
    status_header = page.locator("th:has-text('Mitgliedsstatus')").first
    status_header.wait_for(state="visible", timeout=timeout_ms)
    status_header.click(timeout=timeout_ms)

    filter_menu = page.locator("#menu_col11")
    filter_menu.wait_for(state="visible", timeout=timeout_ms)

    select_all_checkbox = page.locator("#menu_col11 input#col11_all").first
    select_all_checkbox.wait_for(state="visible", timeout=timeout_ms)
    if not select_all_checkbox.is_checked(timeout=timeout_ms):
        select_all_checkbox.check(timeout=timeout_ms, force=True)

    apply_button = page.locator("#menu_col11 input[type='submit'][name='submitcol11']").first
    apply_button.wait_for(state="visible", timeout=timeout_ms)
    apply_button.click(timeout=timeout_ms)

    page.wait_for_load_state("networkidle", timeout=timeout_ms)

    status_header.click(timeout=timeout_ms)
    filter_menu.wait_for(state="visible", timeout=timeout_ms)
    all_selected = page.evaluate(
        """
        () => {
            const menu = document.querySelector('#menu_col11');
            if (!menu) {
                return false;
            }
            const statusCheckboxes = Array.from(menu.querySelectorAll("input[type='checkbox'][id^='col11_']"));
            const valueCheckboxes = statusCheckboxes.filter((cb) => cb.id !== 'col11_all');
            return valueCheckboxes.length > 0 && valueCheckboxes.every((cb) => cb.checked === true);
        }
        """
    )
    if not all_selected:
        raise RuntimeError("Mitgliedsstatus filter is still active. 'Alle auswählen' could not be enforced.")

    page.keyboard.press("Escape")


def export_csv(page: Page, output_path: Path, timeout_ms: int) -> None:
    export_trigger = page.locator("a.tableheader-icon.icon-download-cloud[title*='Daten exportieren']").first
    if export_trigger.count() == 0:
        export_trigger = page.locator("a.tableheader-icon.icon-download-cloud").first
    export_trigger.wait_for(state="visible", timeout=timeout_ms)
    export_trigger.click(timeout=timeout_ms)

    export_menu = page.locator("div[id^='menu_exportlist_']").first
    export_menu.wait_for(state="visible", timeout=timeout_ms)

    with page.expect_download(timeout=timeout_ms) as download_info:
        csv_clicked = click_by_candidates(
            page,
            [
                page.locator("div[id^='menu_exportlist_'] a[href*='output=csv']"),
                page.locator("div[id^='menu_exportlist_'] a:has-text('CSV-Format')"),
                page.locator("div[id^='menu_exportlist_'] a:has-text('CSV herunterladen')"),
            ],
            timeout_ms=timeout_ms,
        )
        if not csv_clicked:
            raise RuntimeError("Could not find CSV export option after opening cloud menu.")

    download = download_info.value
    output_path.parent.mkdir(parents=True, exist_ok=True)
    download.save_as(str(output_path))


def main() -> None:
    parser = argparse.ArgumentParser(description="Download Mitglieder CSV from Vereinsflieger.")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode.")
    parser.add_argument("--timeout-ms", type=int, default=20000, help="Timeout in milliseconds for UI steps.")
    parser.add_argument(
        "--kill-other-runs",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Close older running script instances before starting a new run (default: true).",
    )
    parser.add_argument(
        "--stop-after",
        choices=["login", "club", "mitglieder", "filters", "export"],
        default="export",
        help="Stop after this step for debugging.",
    )
    parser.add_argument(
        "--keep-open-seconds",
        type=int,
        default=0,
        help="When stopping, keep browser open for this many seconds before exit.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output CSV path (relative to app folder if not absolute).",
    )
    parser.add_argument(
        "--quick-config",
        default="Alles",
        help="Quick configuration name to select on the Mitglieder page before export (default: Alles).",
    )
    args = parser.parse_args()

    env = load_env_file(APP_ROOT / ".env.local")
    email = env.get("VEREINSFLIEGER_EMAIL")
    password = env.get("VEREINSFLIEGER_PASSWORD")
    if not email or not password:
        raise RuntimeError("Missing VEREINSFLIEGER_EMAIL or VEREINSFLIEGER_PASSWORD in .env.local")

    ensure_no_other_instances(kill_other_runs=args.kill_other_runs)

    run_stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    default_output = Path("up") / f"mitglieder_{run_stamp}.csv"
    output_path = Path(args.output) if args.output else default_output
    if not output_path.is_absolute():
        output_path = APP_ROOT / output_path

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        context = browser.new_context(accept_downloads=True, locale="de-DE")
        page = context.new_page()

        try:
            login(page, email=email, password=password, timeout_ms=args.timeout_ms)
            ensure_logged_in(page)
            if stop_after_step("login", args.stop_after, args.keep_open_seconds):
                return

            open_club(page, timeout_ms=args.timeout_ms)
            if stop_after_step("club", args.stop_after, args.keep_open_seconds):
                return

            open_mitglieder_left(page, timeout_ms=args.timeout_ms)
            if stop_after_step("mitglieder", args.stop_after, args.keep_open_seconds):
                return

            select_table_quick_config(page, config_label=args.quick_config, timeout_ms=args.timeout_ms)
            clear_table_filters(page, timeout_ms=args.timeout_ms)
            if stop_after_step("filters", args.stop_after, args.keep_open_seconds):
                return

            export_csv(page, output_path=output_path, timeout_ms=args.timeout_ms)
            if stop_after_step("export", args.stop_after, args.keep_open_seconds):
                return

            print(f"CSV downloaded to: {output_path}")
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()