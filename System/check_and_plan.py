#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_and_plan.py
-----------------
Prüft ob Lagerbestand.xlsx seit dem letzten Planer-Durchlauf geändert wurde.
Falls ja: erstellt automatisch einen neuen Wochenplan + Tagespläne.
Falls nein: beendet sich ohne Aktion.

Zusätzlich:
- Archiviert nach jedem erfolgreichen Lauf den aktuellen Lagerstand
  (lagerbestand.json) als System/lager_archiv/lagerbestand_YYYYMMDD.json.
  So existiert immer ein Vorwochen-Stand fürs Auto-Lernen.
- Bei Montag-Lauf wird automatisch lerne_woche.py aufgerufen, sobald ein
  Archiv vor ~7 Tagen existiert. Idempotent: gleicher KW-Eintrag wird in
  Verkaufshistorie.json überschrieben statt dupliziert.

Wird morgens mehrmals ausgeführt (z.B. 6:00, 7:00, 8:00, 9:00 Uhr).
"""
import sys, subprocess, shutil, json
from pathlib import Path
from datetime import datetime, timedelta

BASE         = Path(__file__).parent.parent
LAGER_DATEI  = BASE / "Lagerbestand.xlsx"
LAGER_JSON   = BASE / "lagerbestand.json"
PLANNER      = Path(__file__).parent / "planner.py"
LERNE_WOCHE  = Path(__file__).parent / "lerne_woche.py"
MTIME_FILE   = Path(__file__).parent / ".last_lager_mtime"
LOG_FILE     = Path(__file__).parent / "planer_log.txt"
ARCHIV_DIR   = Path(__file__).parent / "lager_archiv"

ARCHIV_DIR.mkdir(exist_ok=True)


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


def archiviere_lagerstand():
    """Kopiert lagerbestand.json (Fallback: .xlsx) ins Archiv.
    Dateiname: lagerbestand_YYYYMMDD.json — pro Tag ein Eintrag, spaeterer
    ueberschreibt frueheren am gleichen Tag.
    """
    heute = datetime.now().strftime("%Y%m%d")
    if LAGER_JSON.exists():
        ziel = ARCHIV_DIR / "lagerbestand_{}.json".format(heute)
        shutil.copy2(str(LAGER_JSON), str(ziel))
        log("  Archiv: {}".format(ziel.name))
        return ziel
    return None


def finde_vorwochen_archiv(tage_zurueck=7):
    """Findet das Archiv, das ~tage_zurueck Tage alt ist. Erlaubt +/- 2 Tage Toleranz."""
    if not ARCHIV_DIR.exists():
        return None
    ziel_datum = datetime.now() - timedelta(days=tage_zurueck)
    bester_kandidat = None
    bester_abstand = None
    for f in ARCHIV_DIR.glob("lagerbestand_*.json"):
        try:
            datum_str = f.stem.split("_")[1]
            datum = datetime.strptime(datum_str, "%Y%m%d")
        except (IndexError, ValueError):
            continue
        abstand = abs((datum - ziel_datum).total_seconds())
        # Innerhalb von 2 Tagen +/- akzeptieren
        if abstand <= 2 * 86400 and (bester_abstand is None or abstand < bester_abstand):
            bester_kandidat = f
            bester_abstand  = abstand
    return bester_kandidat


def finde_aktuellsten_wochenplan():
    """Findet die juengste Wochenuebersicht_*.xlsx vor heute (= Vorwochen-Plan)."""
    kandidaten = sorted((BASE / "Wochenplan").glob("Wochenuebersicht_*.xlsx"))
    # Filtere auf Pläne von gestern oder älter (heutiger Plan kommt erst gleich)
    heute_str = datetime.now().strftime("%Y%m%d")
    kandidaten = [k for k in kandidaten if heute_str not in k.name]
    return kandidaten[-1] if kandidaten else None


def auto_lerne_vorwoche():
    """Wenn heute Montag und Vorwochen-Archiv existiert: lerne_woche.py aufrufen."""
    if datetime.now().weekday() != 0:
        return  # nur am Montag
    archiv = finde_vorwochen_archiv(tage_zurueck=7)
    if not archiv:
        log("  Auto-Lernen: kein Vorwochen-Archiv gefunden — Schritt uebersprungen")
        return
    plan = finde_aktuellsten_wochenplan()
    if not plan:
        log("  Auto-Lernen: kein Vorwochen-Plan gefunden — Schritt uebersprungen")
        return
    iso = datetime.now().isocalendar()
    kw_label = "{}-{:02d}".format(iso[0], iso[1] - 1 if iso[1] > 1 else 52)
    log("  Auto-Lernen Vorwoche {}: altlager={}, plan={}".format(
        kw_label, archiv.name, plan.name))
    result = subprocess.run(
        [sys.executable, str(LERNE_WOCHE),
         "--altlager", str(archiv),
         "--plan",     str(plan),
         "--kw",       kw_label],
        cwd=str(BASE), capture_output=True, text=True, encoding="utf-8",
    )
    if result.returncode == 0:
        log("  Auto-Lernen erfolgreich.")
        for line in result.stdout.strip().splitlines()[-3:]:
            log("    " + line)
    else:
        log("  Auto-Lernen FEHLER (Exit {})".format(result.returncode))
        for line in (result.stderr or result.stdout).strip().splitlines()[-5:]:
            log("    " + line)


def main():
    # Beobachte BEIDE Lagerquellen: xlsx (manuelle Eingabe) und json (Webformular).
    # Letztere wird vom eingabe.html geschrieben — wenn nur sie sich aendert,
    # soll der Watcher trotzdem ausloesen.
    quellen = []
    if LAGER_DATEI.exists(): quellen.append((LAGER_DATEI, LAGER_DATEI.stat().st_mtime))
    if LAGER_JSON.exists():  quellen.append((LAGER_JSON,  LAGER_JSON.stat().st_mtime))
    if not quellen:
        log("FEHLER: weder Lagerbestand.xlsx noch lagerbestand.json gefunden in {}".format(BASE))
        sys.exit(1)

    aktuelle_mtime = max(m for _, m in quellen)
    letzte_mtime   = lese_letzte_mtime()
    quelle_neueste = max(quellen, key=lambda q: q[1])[0].name

    if aktuelle_mtime <= letzte_mtime:
        log("Keine Änderung am Lagerbestand — kein neuer Plan erstellt.")
        return

    log("Lagerbestand aktualisiert ({}) — starte Planer...".format(quelle_neueste))
    speichere_mtime(aktuelle_mtime)

    # Auto-Lernen VOR dem Planer-Lauf, damit der Forecast den neuen
    # Wocheneintrag schon kennt, wenn der Planer den naechsten Bedarf rechnet.
    auto_lerne_vorwoche()

    result = subprocess.run(
        [sys.executable, str(PLANNER)],
        cwd=str(BASE),
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    if result.returncode == 0:
        log("Planer erfolgreich abgeschlossen.")
        for line in result.stdout.strip().splitlines()[-8:]:
            log("  " + line)
        # Aktuellen Lagerstand archivieren (= zukuenftiger Vorwochen-Stand).
        archiviere_lagerstand()
    else:
        log("FEHLER beim Planer (Exit {})".format(result.returncode))
        for line in (result.stderr or result.stdout).strip().splitlines()[-10:]:
            log("  " + line)


if __name__ == "__main__":
    main()
