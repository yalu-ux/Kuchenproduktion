#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_and_plan.py
-----------------
Prüft ob Lagerbestand.xlsx seit dem letzten Planer-Durchlauf geändert wurde.
Falls ja: erstellt automatisch einen neuen Wochenplan + Tagespläne.
Falls nein: beendet sich ohne Aktion.

Wird morgens mehrmals ausgeführt (z.B. 6:00, 7:00, 8:00, 9:00 Uhr).
"""
import sys, subprocess
from pathlib import Path
from datetime import datetime

BASE        = Path(__file__).parent.parent
LAGER_DATEI = BASE / "Lagerbestand.xlsx"
PLANNER     = Path(__file__).parent / "planner.py"
MTIME_FILE  = Path(__file__).parent / ".last_lager_mtime"  # gespeicherte letzte mtime
LOG_FILE    = Path(__file__).parent / "planer_log.txt"

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    zeile = "[{}] {}".format(ts, msg)
    print(zeile)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(zeile + "\n")

def lese_letzte_mtime():
    if MTIME_FILE.exists():
        try:
            return float(MTIME_FILE.read_text().strip())
        except Exception:
            pass
    return 0.0

def speichere_mtime(mtime):
    MTIME_FILE.write_text(str(mtime))

def main():
    if not LAGER_DATEI.exists():
        log("FEHLER: Lagerbestand.xlsx nicht gefunden: {}".format(LAGER_DATEI))
        sys.exit(1)

    aktuelle_mtime = LAGER_DATEI.stat().st_mtime
    letzte_mtime   = lese_letzte_mtime()

    if aktuelle_mtime <= letzte_mtime:
        log("Keine Änderung am Lagerbestand — kein neuer Plan erstellt.")
        return

    log("Lagerbestand wurde aktualisiert — starte Planer...")
    speichere_mtime(aktuelle_mtime)

    result = subprocess.run(
        [sys.executable, str(PLANNER)],
        cwd=str(BASE),
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    if result.returncode == 0:
        log("Planer erfolgreich abgeschlossen.")
        # Letzte Zeilen der Ausgabe loggen
        for line in result.stdout.strip().splitlines()[-8:]:
            log("  " + line)
    else:
        log("FEHLER beim Planer (Exit {})".format(result.returncode))
        for line in (result.stderr or result.stdout).strip().splitlines()[-10:]:
            log("  " + line)

if __name__ == "__main__":
    main()
