# Projektdokumentation VereinsfliegerTools

## 1) Ziel des Projekts

Dieses Projekt automatisiert drei zusammenhängende Aufgaben für den Verein:

1. Mitgliederdaten aus Vereinsflieger als CSV laden.
2. Daraus Kontaktlisten (CSV/VCF) erzeugen und optional in Apple Kontakte importieren.
3. Einen druckfähigen Mähplan als LaTeX/PDF erzeugen.

Die Skripte sind so gebaut, dass sie lokal auf macOS laufen und sensible Daten nicht im Git landen.

---

## 2) Projektbausteine

### `download_mitglieder_csv.py`

- Loggt sich per Playwright in Vereinsflieger ein.
- Navigiert zur Mitglieder-Tabelle.
- Setzt die gewünschte Tabellenkonfiguration (Standard: `Alles`).
- Exportiert CSV nach `up/mitglieder_YYYY-MM-DD_HH-MM.csv`.

Wichtige Parameter:

- `--headless`
- `--timeout-ms`
- `--kill-other-runs` / `--no-kill-other-runs`
- `--stop-after login|club|mitglieder|filters|export`
- `--keep-open-seconds`
- `--output`
- `--quick-config`

Benötigt in `.env.local`:

- `VEREINSFLIEGER_EMAIL`
- `VEREINSFLIEGER_PASSWORD`

### `create_contacts_lists.py`

- Liest die neueste Mitglieder-CSV (oder `--input-csv`).
- Normalisiert Felder (u. a. Datumswerte).
- Gruppiert nach Mitgliedsstatus inkl. definierter Merge-Regeln:
  - Aktiv + Aktiv (Probe) -> `MFG-Aktive`
  - Fördernd + Ehrenmitglied -> `MFG-FörderUndEhrMitglieder`
- Exportiert pro Gruppe:
  - `contacts_<gruppe>.csv`
  - `contacts_<gruppe>.vcf`
- Optionaler Apple-Kontakte-Import per AppleScript.

Wichtige Defaults:

- `--apply-contacts` ist standardmäßig `true`.
- `--pull-data-from-vereinsflieger` ist standardmäßig `false`.

Mähplan-Integration:

- Liest optional Textquelle (`--maehplan-file` oder `ExcelAusschnittMaehplan2026.txt`).
- Persistiert statische Zuordnung nach `up/MaehplanStatic.json`.
- Überträgt Mähplan-Felder in Kontakte (u. a. `Mähpartner`, `MähterminNr`, `Mähtermin_1`, `Mähtermin_2`).

### `generate_maehplan_tex.py`

- Liest `up/MaehplanStatic.json`.
- Liest Telefonnummern aus neuester Mitglieder-CSV (Mobil priorisiert).
- Erzeugt `up/<YEAR>-MFG-Maehplan.tex`.
- Kompiliert automatisch per `pdflatex` zu `up/<YEAR>-MFG-Maehplan.pdf`.

Aktuelles PDF-Layout:

- 2 Seiten (Split: Termin `<= 14` auf Seite 1, Rest auf Seite 2).
- Kompakter Header ohne Logo.
- Tabellen mit Telefonspalten und Anmerkungen.
- Frühjahrsputz-Zeile als eigene, zentrierte Zeile.
- Seitenfuß mit Stand + Seitenzahl (`Seite 1 von 2`, `Seite 2 von 2`).
- Verantwortlicher pro Termin ist mit gefülltem Punkt markiert (`\bullet`).

Verantwortlichkeitslogik:

- Pro Team und Termin ist genau eine Person verantwortlich.
- Team-interne Alternation:
  - 1. Termin des Teams: erste Person
  - 2. Termin des Teams: zweite Person
  - danach abwechselnd

---

## 3) Datenstruktur

### `up/MaehplanStatic.json`

Top-Level-Felder:

- `Mähtermin_0`: ISO-Datum des Frühjahrsputz-Samstags (Referenz für Terminwochen)
- `FrühjahrsputzInfo`: Anzeigetext für Frühjahrsputz
- pro Mitgliedsschlüssel ein Objekt mit:
  - `Mähpartner`
  - `MähterminNr` (z. B. `"1, 28"`)

Beispielschema:

```json
{
  "Mähtermin_0": "2026-04-11",
  "FrühjahrsputzInfo": "Alle: Frühjahrsputz ...",
  "nachname vorname": {
    "Mähpartner": "Partner Name",
    "MähterminNr": "1, 28"
  }
}
```

---

## 4) Empfohlener Ablauf (operativ)

1. CSV aus Vereinsflieger ziehen:
   - `python download_mitglieder_csv.py`
2. Kontakte erzeugen und optional importieren:
   - mit Import: `python create_contacts_lists.py`
   - ohne Import: `python create_contacts_lists.py --no-apply-contacts`
3. Mähplan erzeugen (inkl. Auto-PDF-Compile):
   - `python generate_maehplan_tex.py`

---

## 5) Fehlerbehebung (Kurz)

- `pdflatex not found`:
  - TeX-Distribution installieren (z. B. MacTeX/TeX Live).
- Kein CSV gefunden:
  - zuerst `download_mitglieder_csv.py` ausführen oder `--input-csv` angeben.
- Apple Kontakte Import schlägt fehl:
  - macOS-Automation/Rechte für Terminal bzw. Python prüfen.
- Fehlende Telefonnummern im Mähplan:
  - CSV enthält keine verwertbaren Werte in Mobil/Telefon-Spalten.

---

## 6) KI-Kontext für zukünftige Unterstützung

Dieser Abschnitt hält fest, was eine KI bei späteren Änderungen zwingend berücksichtigen soll.

### Produkt-/Fachregeln

- Mähplan-PDF bleibt 2-seitig mit Split bei Termin 14.
- Kein Logo im Header (bewusst entfernt).
- Seitenzahlen im Footer beibehalten (`Seite X von 2`).
- Verantwortlicher pro Termin über Punktmarkierung am Namen.
- Verantwortlichkeit je Team alternierend (1. Termin Person 1, 2. Termin Person 2, ...).
- Frühjahrsputz-Hinweis und Verantwortlichkeits-Hinweis im Footer beibehalten.

### Technische Regeln

- `generate_maehplan_tex.py` kompiliert PDF automatisch beim Skriptlauf.
- `create_contacts_lists.py` Defaults nicht unbegründet ändern:
  - `apply_contacts=True`
  - `pull_data_from_vereinsflieger=False`
- `up/` ist lokale Laufzeit-/Exportzone und nicht Git-Quelle.

### Änderungsstrategie für künftige KI-Arbeit

- Bei Layout-Änderungen zuerst nur im LaTeX-Template ändern, dann sofort PDF bauen.
- Bei größeren Template-Eingriffen kleine, isolierte Patches verwenden.
- Nach jeder Änderung an `generate_maehplan_tex.py` mindestens ausführen:
  - `python generate_maehplan_tex.py`
- Bei Mähplan-Datenproblemen zuerst `up/MaehplanStatic.json` validieren.

---

## 7) Ablageorte der Ergebnisse

- Mitglieder-CSV: `up/mitglieder_*.csv`
- Kontaktexporte: `up/contacts_exports/<timestamp>/...`
- Mähplan TEX: `up/<YEAR>-MFG-Maehplan.tex`
- Mähplan PDF: `up/<YEAR>-MFG-Maehplan.pdf`
