Web-Frontend für `online.py` Verarbeitung

Voraussetzungen

- Python 3.10+ (Windows)
- Virtuelle Umgebung empfohlen

Installation

1. In das Projektverzeichnis wechseln:

   ```bash
   cd buchhaltung
   ```

2. Virtuelle Umgebung erstellen und aktivieren (Windows PowerShell):

   ```bash
   python -m venv .venv
   .venv\Scripts\Activate.ps1
   ```

3. Abhängigkeiten installieren:

   ```bash
   pip install -r requirements.txt
   ```

Entwicklung starten

```bash
python -m app.main
```

Die App läuft unter `http://localhost:5000`. Öffnen Sie die Seite im Browser, laden Sie eine Quelldatei hoch, und der Download startet nach Verarbeitung automatisch.

Produktion (Windows, ohne Docker)

```bash
python -m waitress --listen=0.0.0.0:8000 app.main:create_app
```

Rufen Sie dann die App unter `http://<server>:8000` auf.

Hinweise

- CSV-Ausgaben verwenden das Semikolon `;` als Trennzeichen.
- Der Endpunkt „Sale Ausgangsrechnungen“ gibt die Datei im Encoding `cp1252` aus, entsprechend dem bisherigen Verhalten.
- Die alte Desktop-GUI in `online.py` bleibt unverändert; die Web-App erfüllt die gleichen Aufgaben im Browser.


