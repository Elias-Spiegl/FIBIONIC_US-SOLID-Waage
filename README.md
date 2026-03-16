# FIBIONIC US-Solid Waage

Desktop-App fuer die US-Solid Waage `USS-DBS28-50`, die einen RS232-Datenstrom einliest, stabile Gewichte erkennt und den finalen Zahlenwert in eine Excel-Datei schreibt.

Die UI basiert auf `PySide6`, damit das Tool sowohl auf macOS als auch spaeter als Windows-`.exe` sauber laeuft.

## Was die erste Version schon kann

- RS232-Quelle mit den Standardwerten aus dem Handbuch nutzen:
  - `9600 Baud`
  - `8N1`
  - kontinuierlicher Ausgabemodus auf der Waage
- Live-Datenstrom lesen und Rohdaten anzeigen
- Zielgewicht plus Erfassungsfenster definieren
- Schwankungen ueber eine Stabilitaetslogik abfangen
- stabile Messungen automatisch oder nach Bestaetigung in Excel schreiben
- Excel-Datei, Sheet, Spalte, Startzeile und aktuelle Zeile im UI setzen
- nach jedem Write automatisch zur naechsten Zeile springen
- Excel `Auto`, `Datei` und `Live` Modus
- Simulationsmodus fuer Tests auf dem Mac ohne echte Waage
- letzte Einstellungen lokal speichern

## Projektstruktur

```text
src/fibionic_scale_app/
  app.py            # Desktop-UI
  serial_io.py      # serielle Quelle + Simulation
  parsing.py        # Parser fuer die Scale-Ausgabe
  stability.py      # Stabilitaets- und Capture-Logik
  excel_writer.py   # Schreiben nach Excel
  settings_store.py # lokale Settings
tests/
  test_excel_writer.py
  test_parsing.py
  test_stability.py
```

## Lokaler Start

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
PYTHONPATH=src python -m fibionic_scale_app
```

Wenn du das Projekt lieber als Paket im Entwicklungsmodus installieren willst, geht auch:

```bash
pip install -e .
fibionic-scale
```

Wenn du zuerst nur auf dem Mac testen willst:

1. App starten
2. `Simulationsmodus` aktiv lassen
3. Excel-Datei auswaehlen
4. Zielgewicht setzen
5. Quelle starten

Dann siehst du im UI, wie ein stabiler Wert erkannt und in Excel geschrieben wird.

## Excel-Modi

Die App unterstuetzt drei Schreibmodi:

- `Auto`
  Auf macOS und Windows versucht die App zuerst den Live-Writer ueber die lokal installierte Excel-App. Wenn das nicht klappt, faellt sie automatisch auf den Datei-Modus zurueck.
- `Datei-Modus`
  Die `.xlsx`-Datei wird direkt per `openpyxl` geschrieben. Das ist plattformneutral, aber Aenderungen sind in einer bereits geoeffneten Excel-Datei meist nicht sofort sichtbar.
- `Live-Modus`
  Die App schreibt direkt in die lokal laufende Excel-Anwendung ueber `xlwings`. Das ist fuer deinen OneDrive/Excel-Workflow der beste Modus, wenn die Datei waehrend des Loggens offen bleiben soll.

Hinweise zum Live-Modus:

- funktioniert auf `macOS` und `Windows`
- benoetigt eine lokal installierte Desktop-Version von Microsoft Excel
- funktioniert nicht fuer reine Excel-Online-Sessions ohne lokale Excel-App
- die Datei sollte auf demselben Rechner liegen bzw. von demselben Rechner aus in Excel geoeffnet werden

## Echte Waage anschliessen

Fuer die `USS-DBS28-50` gehen wir aktuell von dem Format aus, das du im Handbuchfoto gezeigt hast:

- 9-polige RS232-Verbindung
- Standard-Baudrate `9600 bps`
- 1 Startbit, 8 Datenbits, 1 Stopbit
- kontinuierliche Ausgabe auf der Waage aktivieren (`C5-0`)

Die Parser-Logik ist absichtlich tolerant gebaut und extrahiert den Zahlenwert auch dann, wenn die Waage Leerzeichen und Einheit mitsendet.

Wichtig fuer die aktuelle Konfiguration: Die Waage sendet bei dir momentan Werte in `g`. Das bedeutet:

- `Zielgewicht` ist in Gramm einzugeben
- `Fenster +/-` ist in Gramm einzugeben
- `Stabilitaets-Toleranz`, `Reset-Schwelle` und `Mindestgewicht` sind ebenfalls in Gramm zu verstehen
- in Excel wird nur der nackte Zahlenwert geschrieben, also z. B. `44.00`

In der UI sind diese Felder deshalb auch mit `(g)` gekennzeichnet.

## Excel-Workflow

Das Tool schreibt immer nur den reinen Zahlenwert in die konfigurierte Zelle.

Beispiel:

- Datei: `Messwerte.xlsx`
- Sheet: `Produktion`
- Spalte: `F`
- aktuelle Zeile: `12`

Dann landet die naechste stabile Messung in `Produktion!F12`. Wenn `Auto-Advance` aktiv ist, springt die App anschliessend auf `F13`.

Wenn du waehrenddessen in Excel direkt sehen willst, wie der Wert erscheint, stelle den Excel-Modus im UI auf `Live` oder `Auto`.

## Windows `.exe` bauen

Auf dem Windows-Rechner kannst du spaeter mit `PyInstaller` eine einzelne `.exe` erzeugen:

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm --windowed --name FIBIONIC-Waage -F -p src src/fibionic_scale_app/__main__.py
```

Danach liegt die fertige Datei unter `dist/FIBIONIC-Waage.exe`.

## Tests

Die kleinen Backend-Tests laufen ohne echte Waage:

```bash
PYTHONPATH=src python -m unittest discover -s tests
```

## Naechste sinnvolle Schritte

- echte Rohdaten der Waage einmal mitschneiden und Parser feinjustieren
- optionalen COM-Port-Testbutton einbauen
- optionalen manuellen "Naechste Zeile"-Button ergaenzen
- optional Excel-Datei waehrend des Schreibens gegen paralleles Oeffnen absichern
