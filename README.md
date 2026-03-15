# VereinsfliegerTools

Python-Werkzeug, das automatisch die Mitgliederliste von **vereinsflieger.de** herunterlädt und daraus aktualisierte **E-Mail-Verteilerlisten für Apple Mail auf dem Mac** erzeugt.

Da das kostenlose Abo von vereinsflieger.de keinen API-Zugang bietet, wird die Webseite mit einem Browser-Automationstool (Playwright) bedient.

---

## Voraussetzungen

* Python ≥ 3.10
* macOS (für den abschließenden AppleScript-Schritt; Download und Parsing funktionieren auch unter Linux/Windows)

---

## Installation

```bash
# 1. Repository klonen
git clone https://github.com/phidor1708/VereinsfliegerTools.git
cd VereinsfliegerTools

# 2. Abhängigkeiten installieren
pip install -r requirements.txt

# 3. Chromium-Browser für Playwright installieren
playwright install chromium
```

---

## Konfiguration

```bash
cp config.example.yaml config.yaml
```

`config.yaml` öffnen und die eigenen Zugangsdaten eintragen:

```yaml
vereinsflieger:
  username: "deine@email.de"
  password: "deinPasswort"
  club_id: ""           # optional: Vereinsnummer

export:
  group_name: "Modellflugverein Höchstadt"   # Name der Gruppe in Apple Contacts
  output_dir: "./output"
```

> **Hinweis:** `config.yaml` ist in `.gitignore` eingetragen und wird **nicht** ins Repository committet.

---

## Verwendung

### Normaler Ablauf (Download + Export)

```bash
python main.py
```

Das Werkzeug:
1. Loggt sich automatisch auf vereinsflieger.de ein.
2. Lädt die Mitgliederliste als CSV herunter (`output/members.csv`).
3. Erzeugt eine vCard-Datei (`output/members.vcf`).
4. Erzeugt ein AppleScript (`output/update_contacts_group.applescript`).

### Vorhandene CSV verwenden (ohne Download)

Wenn die CSV-Datei bereits vorliegt (z. B. manuell exportiert):

```bash
python main.py --csv pfad/zur/members.csv
```

### Alle Optionen

```
python main.py --help

optional arguments:
  --config FILE   Pfad zur YAML-Konfigurationsdatei (Standard: config.yaml)
  --csv FILE      Vorhandene CSV-Datei verwenden (Download wird übersprungen)
  --output-dir DIR  Ausgabeverzeichnis (überschreibt den Wert aus config.yaml)
```

---

## Mac-Mail-Verteilerliste einrichten

Nach dem Ausführen von `python main.py` erscheint folgende Ausgabe:

```
Done! Next steps on your Mac:

  1. Open 'output/members.vcf' to import all members into Apple Contacts.
     (double-click the file or drag it onto the Contacts app)

  2. Run the AppleScript to create / update the 'Modellflugverein Höchstadt'
     group in Apple Contacts:
       osascript output/update_contacts_group.applescript
     The group then appears as a distribution list in Apple Mail.
```

**Schritt 1 – Kontakte importieren:**  
`output/members.vcf` per Doppelklick in der Contacts-App öffnen.

**Schritt 2 – Gruppe anlegen/aktualisieren:**  
```bash
osascript output/update_contacts_group.applescript
```
Dieses Script legt die Gruppe *Modellflugverein Höchstadt* in Apple Contacts an (falls nicht vorhanden) und fügt alle Mitglieder hinzu. In Apple Mail steht die Gruppe dann als Verteilerliste zur Verfügung – einfach den Gruppennamen in das An-Feld schreiben.

> **Tipp:** Das Script kann nach jedem Mitgliederupdate erneut ausgeführt werden. Bereits vorhandene Kontakte werden wiederverwendet; neue Mitglieder werden hinzugefügt.

---

## Projektstruktur

```
VereinsfliegerTools/
├── main.py                          # CLI-Einstiegspunkt
├── config.example.yaml              # Beispielkonfiguration
├── requirements.txt                 # Python-Abhängigkeiten
├── src/
│   ├── downloader.py                # Login und CSV-Download (Playwright)
│   ├── parser.py                    # CSV-Parsing
│   └── exporter.py                  # vCard- und AppleScript-Export
└── tests/
    ├── test_parser.py
    └── test_exporter.py
```

---

## Tests ausführen

```bash
python -m pytest tests/ -v
```

---

## Häufige Probleme

| Problem | Lösung |
|---|---|
| `Login failed` | Benutzername/Passwort in `config.yaml` prüfen |
| `playwright is required` | `pip install playwright && playwright install chromium` |
| Keine E-Mail-Spalte gefunden | CSV manuell öffnen und Spaltennamen in `src/parser.py` (`_EMAIL_CANDIDATES`) ergänzen |
| AppleScript fragt nach Zugriffsrechten | In *Systemeinstellungen → Datenschutz → Automatisierung* den Zugriff erlauben |
