# fibionic Gewichtslogging

Desktop-App für die US-Solid Waage `USS-DBS28-50`.  
Die Anwendung liest den Datenstrom der Waage ein, erkennt stabile Messwerte und schreibt den Zahlenwert automatisch in eine Excel-Datei.

Die App läuft lokal auf `macOS` und kann später als Windows-`.exe` für den Produktions-PC gebaut werden.

## App-Icon

Für die App liegt jetzt ein eigenes Icon-Set im Projekt:

- `logo/fibionic_app_icon.svg`
- `logo/fibionic_app_icon.ico`
- `logo/fibionic_app_icon.icns`

Beim normalen Start auf dem Mac wird das Icon direkt in die Qt-App geladen.  
Für gebaute Bundles und `.exe`-Dateien wird es ebenfalls verwendet.

## Zweck der App

Das Tool ist für einen einfachen Produktionsablauf gedacht:

1. Bauteil auf die Waage legen
2. auf stabilen Messwert warten
3. Wert automatisch in Excel schreiben
4. nächstes Bauteil wiegen

Die App kümmert sich dabei um:

- Verbindung zur Waage
- Stabilitätserkennung
- Prüfung des Zielbereichs
- Ermittlung der nächsten freien Excel-Zelle
- Schreiben des Messwerts in Excel

## Voraussetzungen

Für den Betrieb mit echter Waage:

- US-Solid Waage `USS-DBS28-50`
- RS232-Verbindung zur Waage
- bei macOS oder Windows in der Regel ein USB-zu-RS232-Adapter
- Microsoft Excel lokal installiert, wenn Live-Schreiben genutzt werden soll

Für die App:

- Python-Umgebung mit den Paketen aus `requirements.txt`
- oder später eine gebaute Windows-`.exe`

## Inbetriebnahme

### 1. Waage einstellen

Die Waage sollte für dieses Projekt so konfiguriert sein:

- `9600 Baud`
- `8N1`
- kontinuierliche Ausgabe aktiv
- Einheit auf `g` stellen

Wichtig: Die App erwartet Messwerte in Gramm.  
Wenn die Waage z. B. in `kg` sendet, erscheint eine Fehlermeldung und du musst die Einheit an der Waage auf `g` umstellen.

### 2. Projekt lokal starten

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
PYTHONPATH=src python -m fibionic_scale_app
```

Alternativ im Entwicklungsmodus:

```bash
pip install -e .
fibionic-scale
```

### App-Icon am Mac testen

```bash
PYTHONPATH=src .venv/bin/python -m fibionic_scale_app
```

Danach solltest du das neue Icon in der laufenden App bzw. im Dock sehen.

## Bedienungsanleitung

### 1. Quelle einrichten

Im Bereich `Quelle` arbeitet die App immer mit der echten Waage.

Wenn eine Waage angeschlossen ist, sucht die App den seriellen Port automatisch.  
Wenn nötig, kannst du den Port manuell auswählen.

### 2. Excel-Ziel einstellen

Im Bereich `Excel` legst du fest:

- Excel-Datei
- `Sheet`
- `Spalte`
- `Zeile`
- `Logging-Format`

`Logging-Format` bedeutet:

- `Oben nach unten`
- `Links nach rechts`

Die App sucht von dieser Startposition aus immer selbst die nächste freie Zelle.

### 3. Messwerte einstellen

Im Bereich `Messwerte` setzt du:

- `Zielgewicht (g)`
- `Abweichung +/- (g)`

Beides immer in Gramm.

### 4. Quelle starten

Mit `Quelle starten` beginnt der Messbetrieb.

Sobald die Quelle läuft:

- werden `Messwerte` und `Excel` links ausgeblendet
- im Quellen-Widget wird oben die aktive Quelle angezeigt
- rechts steht der große Statusbereich im Fokus

### 5. Während des Betriebs

Die drei großen Karten zeigen:

- `Live-Wert`
- `Stabiler Messwert`
- `Nächste Zelle`

Darunter stehen die kleineren Live-Informationen:

- `Zielbereich`
- `Logging-Format`

Wenn ein Messwert erfolgreich gespeichert wurde, wird der Statusbereich kurz visuell hervorgehoben.

### 6. Logging pausieren

Während die Quelle läuft, gibt es zwei Zustände:

- `Logging pausieren`
- `Logging fortsetzen`

`Logging pausieren` bedeutet:

- die Quelle bleibt verbunden
- es wird nichts in Excel geschrieben
- die Bereiche `Messwerte` und `Excel` werden wieder eingeblendet
- Einstellungen können angepasst werden

`Logging fortsetzen` setzt den normalen Messablauf wieder fort.

### 7. Stopp

Mit `Stopp` wird die Quelle beendet.

Danach:

- springt die App zurück in den Startzustand
- alle Einstellungsbereiche sind wieder sichtbar
- die Quelle kann neu gestartet werden

## Excel-Verhalten

Die App schreibt nur den reinen Zahlenwert in Excel.  
Die Einheit wird nicht mitgeschrieben.

Beispiel:

- Datei: `Messwerte.xlsx`
- Sheet: `Produktion`
- Spalte: `F`
- Zeile: `12`
- Logging-Format: `Oben nach unten`

Dann schreibt die App z. B. nach:

- `Produktion!F12`
- `Produktion!F13`
- `Produktion!F14`

Bei `Links nach rechts` entsprechend:

- `Produktion!F12`
- `Produktion!G12`
- `Produktion!H12`

## Fehlermeldungen und typische Fälle

### Keine Waage gefunden

Wenn keine Waage automatisch erkannt wird:

- `Automatisch erkennen`
- bei Bedarf `Port manuell wählen`
- Verkabelung und USB-RS232-Adapter prüfen

### Falsche Einheit

Wenn die Waage nicht in Gramm sendet:

- erscheint eine Warnung
- das Logging wird angehalten
- die Einheit an der Waage muss auf `g` gestellt werden

### Excel-Datei kann nicht beschrieben werden

Prüfen:

- ist die richtige Datei ausgewählt?
- ist Excel lokal installiert?
- ist die Datei lokal erreichbar?
- ist die Datei im lokalen Desktop-Excel geöffnet, wenn Live-Schreiben gewünscht ist?

### OneDrive

Bitte keine `.xlsx` direkt aus einem OneDrive-Ordner verwenden.

Empfohlener Ablauf:

- lokal außerhalb von OneDrive loggen
- Datei nach dem Logging oder in einem separaten Schritt synchronisieren

Die App blockiert OneDrive-Dateien bewusst mit einer klaren Meldung, weil das Schreiben dort in der Praxis oft von OneDrive oder Excel gesperrt wird.

## Windows-`.exe` bauen

Für den späteren Einsatz auf Windows:

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller fibionic-gewichtslogging.spec
```

Die fertige Datei liegt danach unter:

```text
dist/fibionic-gewichtslogging.exe
```

Auf dem Mac kannst du mit derselben Spec-Datei auch ein `.app`-Bundle bauen:

```bash
pyinstaller fibionic-gewichtslogging.spec
```

Das Bundle liegt danach unter:

```text
dist/fibionic-gewichtslogging.app
```

## Tests

Die vorhandenen Tests laufen ohne echte Waage:

```bash
PYTHONPATH=src python -m unittest discover -s tests
```

## Projektstruktur

```text
src/fibionic_scale_app/
  app.py            # Desktop-UI
  serial_io.py      # serielle Waagen-Anbindung
  parsing.py        # Parser für die Waagendaten
  stability.py      # Stabilitätslogik
  excel_writer.py   # Excel-Anbindung
  settings_store.py # lokale Speicherung der UI-Einstellungen
tests/
  test_excel_writer.py
  test_parsing.py
  test_scale_sources.py
  test_stability.py
```
