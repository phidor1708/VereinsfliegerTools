# VereinsfliegerTools

## Projektdokumentation

- Vollständige Doku: [docs/PROJEKTDOKUMENTATION.md](docs/PROJEKTDOKUMENTATION.md)

## Mitglieder CSV Download

### 1) Install dependencies

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

### 2) Store credentials locally

Use `.env.local` in the project root:

```env
VEREINSFLIEGER_EMAIL=your-email
VEREINSFLIEGER_PASSWORD=your-password
```

`.env.local` is git-ignored.

### 3) Run automation

```bash
python download_mitglieder_csv.py
```

The script logs in, opens `MFG Höchstadt/Aisch`, navigates to `Mitglieder`, clears table filters, exports CSV, and saves it as `up/mitglieder_YYYY-MM-DD.csv`.

The `up/` folder is local-only (`.gitignore`) and is not synchronized with git.

Optional:

```bash
python download_mitglieder_csv.py --headless --output up/mitglieder_export.csv --timeout-ms 30000
```

### 4) Step-by-step debugging

The script can stop after each stage:

```bash
python download_mitglieder_csv.py --stop-after login --keep-open-seconds 300
python download_mitglieder_csv.py --stop-after club --keep-open-seconds 300
python download_mitglieder_csv.py --stop-after mitglieder --keep-open-seconds 300
python download_mitglieder_csv.py --stop-after filters --keep-open-seconds 300
python download_mitglieder_csv.py --stop-after export --keep-open-seconds 300
```

Instance handling:

- By default, `--kill-other-runs` is active, so older running script instances are closed automatically.
- Use `--no-kill-other-runs` if you want strict single-instance blocking instead of auto-close.

## Create Apple Contacts Lists from CSV

New script: `create_contacts_lists.py`

What it does:

- Loads the latest CSV from `up/` (or a file passed via `--input-csv`)
- Builds a pandas DataFrame
- Merges `aktiv` and `aktiv (Probe)` into one list named `MFG-Aktive`
- Creates contact exports for every other membership status group that is not empty
- Writes per-status `.csv` and `.vcf` files to `up/contacts_exports/<timestamp>/`

Run export only (safe, no Contacts changes):

```bash
python create_contacts_lists.py
```

Optional: create new Apple Contacts lists from the same data:

```bash
python create_contacts_lists.py --apply-contacts
```

Safety behavior for Apple Contacts:

- The script **only creates new lists**.
- It refuses to run if a target list name already exists.
- It does not rename, delete, or modify existing lists.
