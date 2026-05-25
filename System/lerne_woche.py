#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
lerne_woche.py
--------------
Leitet aus Lager-Differenz + Produktion einer Vorwoche die tatsaechlich
verkauften Mengen ab und schreibt sie in Verkaufshistorie.json.

Formel pro Produkt:
    verkauf_woche = lager_vorwoche_montag + produktion_woche - lager_diese_woche_montag

Verwendung:
    python3 lerne_woche.py
    -> interaktiv: fragt nach altem Lagerstand und Wochenplan-Pfad

    python3 lerne_woche.py --altlager altlager.json --plan Wochenplan/Wochenuebersicht_20260518.xlsx
    -> nicht-interaktiv

Output:
    Aktualisiert Verkaufshistorie.json um einen neuen Wocheneintrag.
"""
import sys, json, argparse
from pathlib import Path
from datetime import datetime

BASE = Path(__file__).parent.parent
LAGER_JSON = BASE / "lagerbestand.json"
HISTORIE   = BASE / "Verkaufshistorie.json"
IST_DIR    = BASE / "Wochenplan"

PRODUKTE = [
    "Beeren Tartelette","Bienenstich","Donauwelle",
    "Kaese-Rhabarber Schnitte","Kaesekuchen","Lauch-Speck Quiche",
    "Mango-Passionsfrucht Tartelette","Pistazien Toertchen",
]

# Aufteilung der Wochensumme auf Cafe 1 vs Cafe 2.
# Fester Anteil: Cafe 2 bekommt einmal woechentlich (Dienstag) eine grosse
# Lieferung — der Cafe-1-Anteil ergibt sich aus der restlichen Wochensumme.
# Aus den bisherigen Verkaufszahlen1.xlsx geschaetzt; kann pro Produkt
# verfeinert werden, wenn Yannic genauere Daten hat.
CAFE2_ANTEIL = {
    "Beeren Tartelette":               0.38,
    "Bienenstich":                     0.33,
    "Donauwelle":                      0.33,
    "Kaese-Rhabarber Schnitte":        0.33,
    "Kaesekuchen":                     0.29,
    "Lauch-Speck Quiche":              0.36,
    "Mango-Passionsfrucht Tartelette": 0.38,
    "Pistazien Toertchen":             0.0,   # nur Cafe 1
}


def lade_json(pfad):
    if not Path(pfad).exists():
        print("FEHLER: {} nicht gefunden".format(pfad)); sys.exit(1)
    raw = Path(pfad).read_text(encoding='utf-8', errors='replace')
    # OneDrive-Workaround: trailing Null-Bytes / Whitespace abschneiden
    raw = raw.rstrip('\x00 \t\r\n')
    return json.loads(raw)


def lese_produktion_aus_excel(pfad):
    """Liest die Spalte 'Produzieren' aus Wochenuebersicht_*.xlsx (Fallback,
    wenn keine ist_produktion_KW*.json vorhanden)."""
    import openpyxl
    wb = openpyxl.load_workbook(str(pfad), data_only=True)
    ws = wb.active
    produktion = {}
    # Spalten: A=Produkt, ..., G=Produzieren (laut erstelle_excel-Layout)
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row[0] or row[0] == "Produkt": continue
        produkt = str(row[0]).strip()
        # Mapping auf Standardnamen
        if "Käse-Rhabarber" in produkt: produkt = "Kaese-Rhabarber Schnitte"
        elif "Käsekuchen" in produkt: produkt = "Kaesekuchen"
        elif "Pistazien" in produkt and "Törtchen" in produkt: produkt = "Pistazien Toertchen"
        try:
            menge = float(row[6] or 0)
        except (TypeError, ValueError):
            menge = 0
        if produkt in PRODUKTE:
            produktion[produkt] = menge
    return produktion


def lese_produktion_aus_ist(kw_label):
    """Liest die tatsaechlich produzierte Menge aus ist_produktion_KW*.json.
    Erwartet kw_label im Format 'YYYY-WW'.

    JSON-Schema:
      { "kw": "2026-21", "tage": {
          "Montag":  { "Kaesekuchen": { "ist_menge": 2, "erledigt": true,
                                        "verschoben_nach": null } },
          ...
      }}

    Verschiebungen werden korrekt mitgeloest:
      - Eintrag mit "verschoben_nach": X  ->  zaehlt NICHT an seinem Tag.
      - Eintrag mit "verschoben_von": Y   ->  zaehlt am Verschiebungs-Zieltag.
    Da die Wochensumme das einzige ist was lerne_woche.py braucht, ist die
    Tageverteilung egal — wir summieren ueber alle nicht-verschobenen
    Eintraege (oder mit "erledigt": true) und ueber zugezogene Eintraege.

    Returns:
        ({"Kaesekuchen": 4.0, ...}, anzahl_eintraege) oder (None, 0) wenn
        Datei nicht existiert.
    """
    pfad = IST_DIR / "ist_produktion_KW{}.json".format(kw_label)
    if not pfad.exists():
        return None, 0
    try:
        data = lade_json(pfad)
    except SystemExit:
        return None, 0
    produktion = {pk: 0.0 for pk in PRODUKTE}
    eintraege = 0
    for tag, pks in (data.get("tage") or {}).items():
        for pk, rec in (pks or {}).items():
            if pk not in PRODUKTE:
                continue  # Hilfs-Tasks (_mt_ansetzen etc.) ignorieren
            # Wenn auf anderen Tag verschoben: hier nicht zaehlen
            if rec.get("verschoben_nach"):
                continue
            # ist_menge ueberschreibt Plan; nur zaehlen wenn erledigt oder ist_menge>0
            ist = rec.get("ist_menge")
            if ist is None:
                # Kein expliziter Edit: nur wenn als erledigt markiert, zaehlen wir
                # die Planmenge (kennen wir hier aber nicht). Konservativ: 0.
                # Der Planer fuellt erledigte-Default-Mengen via JS bereits.
                if rec.get("erledigt"):
                    pass  # bleibt 0 wenn ist_menge fehlt
                continue
            try:
                produktion[pk] += float(ist)
                eintraege += 1
            except (TypeError, ValueError):
                pass
    return produktion, eintraege


def berechne_woche(alt_lager, neu_lager, produktion):
    """verkauf = alt_lager + produktion - neu_lager (geclippt auf >=0)."""
    woche = {"Cafe 1": {}, "Cafe 2": {}}
    for pk in PRODUKTE:
        alt = float(alt_lager.get(pk, 0))
        neu = float(neu_lager.get(pk, 0))
        prod = float(produktion.get(pk, 0))
        gesamt = max(0.0, alt + prod - neu)
        c2_anteil = CAFE2_ANTEIL.get(pk, 0.33)
        c2 = round(gesamt * c2_anteil, 1)
        c1 = round(gesamt - c2, 1)
        woche["Cafe 1"][pk] = c1
        woche["Cafe 2"][pk] = c2
    return woche


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--altlager", help="Pfad zu altem lagerbestand.json (Stand Vorwoche Montag)")
    ap.add_argument("--plan", help="Pfad zu Wochenuebersicht_YYYYMMDD.xlsx der Vorwoche")
    ap.add_argument("--kw", help="Kalenderwoche-Label (Default: heutige KW-1)")
    args = ap.parse_args()

    if not args.altlager:
        print("Bitte Pfad zu altem lagerbestand.json angeben (Stand Vorwoche Montag).")
        args.altlager = input("> ").strip()
    if not args.plan:
        # Default: jüngste Wochenuebersicht_*.xlsx
        kandidaten = sorted((BASE / "Wochenplan").glob("Wochenuebersicht_*.xlsx"))
        if kandidaten:
            args.plan = str(kandidaten[-1])
            print("Verwende Wochenplan: {}".format(args.plan))
        else:
            print("Kein Wochenplan gefunden. Pfad bitte angeben:")
            args.plan = input("> ").strip()

    alt_lager  = lade_json(args.altlager)
    neu_lager  = lade_json(LAGER_JSON)

    if args.kw:
        kw = args.kw
    else:
        iso = datetime.now().isocalendar()
        kw = "{}-{:02d}".format(iso[0], iso[1] - 1 if iso[1] > 1 else 52)

    # Bevorzugte Quelle: ist_produktion_KW*.json (echte produzierte Mengen
    # aus den Tagesplan-Häkchen). Fallback: Plan-Spalte aus xlsx.
    quelle = "plan-xlsx"
    produktion, anz = lese_produktion_aus_ist(kw)
    if produktion is not None and anz > 0:
        quelle = "ist-produktion-json"
        print("Quelle: ist_produktion_KW{}.json ({} Eintraege)".format(kw, anz))
    else:
        if produktion is not None:
            print("Hinweis: ist_produktion_KW{}.json existiert, aber leer -> Fallback xlsx".format(kw))
        produktion = lese_produktion_aus_excel(args.plan)
        print("Quelle: {} (Fallback - Plan, nicht Ist!)".format(Path(args.plan).name))

    woche = berechne_woche(alt_lager, neu_lager, produktion)

    eintrag = {
        "kw":  kw,
        "abgeleitet_am": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "quelle": "lager-differenz",
        "produktions_quelle": quelle,
        "verkauf": woche,
    }

    historie = lade_json(HISTORIE) if HISTORIE.exists() else {"wochen": [], "_alpha": 0.3}
    # Doppel-Eintrag fuer gleiche KW vermeiden — ersetzen
    historie["wochen"] = [w for w in historie.get("wochen", []) if w.get("kw") != kw]
    historie["wochen"].append(eintrag)
    HISTORIE.write_text(json.dumps(historie, indent=2, ensure_ascii=False), encoding='utf-8')

    print("\nWoche {} abgeleitet:".format(kw))
    for pk in PRODUKTE:
        c1 = woche["Cafe 1"][pk]; c2 = woche["Cafe 2"][pk]
        print("  {:35s} C1={:6.1f}  C2={:6.1f}  (gesamt {:.1f})".format(pk, c1, c2, c1+c2))
    print("\nIn {} gespeichert ({} Wochen Historie total).".format(
        HISTORIE.name, len(historie["wochen"])))


if __name__ == "__main__":
    main()
