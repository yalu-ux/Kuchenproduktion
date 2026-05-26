#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Kuchenproduktion Wochenplaner"""
import struct, zlib, xml.etree.ElementTree as ET
import openpyxl, math, re, json
from pathlib import Path
from datetime import datetime
from collections import defaultdict

BASE                 = Path(__file__).parent.parent
OUTPUT_PFAD          = BASE / "Wochenplan"
LAGERBESTAND_DATEI   = BASE / "Lagerbestand.xlsx"
VERKAUFSZAHLEN_DATEI = BASE / "Verkaufszahlen1.xlsx"
VERKAUFSHISTORIE     = BASE / "Verkaufshistorie.json"
KONFIGURATION_DATEI  = BASE / "Konfiguration.xlsx"
REZEPTE_PFAD         = BASE / "Rezepte"
OUTPUT_PFAD.mkdir(exist_ok=True)

TAGE        = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag","Sonntag"]
ARBEITSTAGE = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]

def cafe2_bereits_abgeholt():
    """True wenn heute Dienstag oder spaeter — Cafe-2-Abgabe gilt als erledigt.
    Der Lagerbestand spiegelt dann bereits den Bestand NACH der Cafe-2-Entnahme.
    """
    return datetime.now().weekday() >= 1   # 0=Mo, 1=Di, 2=Mi ...

def ab_heute_tage(ab_morgen=False):
    """Gibt die noch verbleibenden Arbeitstage zurueck (ab heute inkl.).
    Samstag/Sonntag → komplette Woche (Planung fuer naechste Woche).
    ab_morgen=True → startet erst ab dem naechsten Arbeitstag (z.B. wenn
    Zahlen nach der heutigen Produktion aktualisiert werden).
    """
    wochentag_idx = datetime.now().weekday()  # 0=Mo, 1=Di, ..., 6=So
    if wochentag_idx >= 5:   # Wochenende → volle Woche planen
        return list(ARBEITSTAGE)
    if ab_morgen:
        # Naechsten Arbeitstag als Start: Freitag → volle naechste Woche
        naechster = wochentag_idx + 1
        if naechster >= 5:   # Wochenende ueberspringen → volle Woche
            return list(ARBEITSTAGE)
        return ARBEITSTAGE[naechster:]
    return ARBEITSTAGE[wochentag_idx:]

FROSTER_SCHWELLEN = {
    "Beeren Tartelette":               30,
    "Mango-Passionsfrucht Tartelette": 30,
    "Pistazien Toertchen":             30,
}

# C2-Puffer: Stueckzahl je Produkt, die Montag im Froster sein muss (Cafe-2-Lieferung Di)
# Wird aus Konfiguration.xlsx gelesen; Fallback-Wert hier
C2_PUFFER = {
    "Donauwelle":              4,
    "Kaese-Rhabarber Schnitte":4,
    "Kaesekuchen":             4,
    "Bienenstich":             4,
    "Lauch-Speck Quiche":      4,
}

# max_ofen = Stueck pro Ofengang | formen = Backformen vorhanden | min_charge = Mindestmenge
PRODUKT_CONFIG = {
    "Donauwelle":                     {"max_ofen":6,  "min_charge":6,  "formen":6,  "cafe2":True, "murbeteig_typ":"30x20 Schnitte","rezept":"Donauwelle.xlsx"},
    "Kaese-Rhabarber Schnitte":       {"max_ofen":6,  "min_charge":6,  "formen":6,  "cafe2":True, "murbeteig_typ":"30x20 Schnitte","rezept":"Käse-Rhabarber Schnitte.xlsx"},
    "Kaesekuchen":                    {"max_ofen":2,  "min_charge":2,  "formen":4,  "cafe2":True, "murbeteig_typ":"Murbeteig rund", "rezept":"Käsekuchen.xlsx"},
    "Bienenstich":                    {"max_ofen":6,  "min_charge":6,  "formen":6,  "cafe2":True, "murbeteig_typ":None,            "rezept":"Bienenstich.xlsx"},
    "Lauch-Speck Quiche":             {"max_ofen":6,  "min_charge":5,  "formen":6,  "cafe2":True, "murbeteig_typ":None,            "rezept":"Lauch-Speck Quiche.xlsx"},
    "Beeren Tartelette":              {"max_ofen":60, "min_charge":60, "formen":60, "cafe2":True, "murbeteig_typ":"Tartelette",    "rezept":"Beeren Tartelette.xlsx"},
    "Mango-Passionsfrucht Tartelette":{"max_ofen":60, "min_charge":60, "formen":60, "cafe2":True, "murbeteig_typ":"Tartelette",    "rezept":"Mango-Passionsfrucht Tartelette.xlsx"},
    "Pistazien Toertchen":            {"max_ofen":60, "min_charge":60, "formen":60, "cafe2":False,"murbeteig_typ":None,            "rezept":"Pistazien-Schoko Törtchen.xlsx"},
}

# Produktspezifische Mindestchargen pro Tag (laut Planungsregeln.md).
# Diese Chargen werden VOR der bestandssicheren Verteilung reserviert, damit
# der dokumentierte Frische-/Rhythmus-Anspruch eingehalten wird. Der Restbedarf
# wird dann ueber verteile_bestandssicher auf die uebrigen Tage gelegt.
# (Käsekuchen hat seine eigene, komplexere Logik in erstelle_wochenplan.)
MIN_CHARGEN = {
    "Kaese-Rhabarber Schnitte": {"Donnerstag": 1},
    "Lauch-Speck Quiche":       {"Montag": 1},
}

# Tiebreaker-Reihenfolge je Produkt (laut Planungsregeln.md).
# Wenn mehrere Tage gleich viel Aktiv-Zeit haben, gewinnt der Tag mit
# dem niedrigsten Index in dieser Liste. Default = allgemeine Reihenfolge
# (Fr -> Mo -> Mi -> Do -> Di), siehe Planungsregeln Z. 21-25.
TIEBREAKER_DEFAULT = ["Freitag","Montag","Mittwoch","Donnerstag","Dienstag"]
TIEBREAKER = {
    "Kaesekuchen":                ["Freitag","Montag","Mittwoch","Donnerstag","Dienstag"],
    "Kaese-Rhabarber Schnitte":   ["Donnerstag","Montag","Dienstag","Mittwoch","Freitag"],
    "Lauch-Speck Quiche":         ["Montag","Donnerstag","Dienstag","Mittwoch","Freitag"],
}

def tiebreaker_rang(produkt_key, tag):
    """Kleinerer Rang = bevorzugt. Default-Liste fuer unbekannte Produkte."""
    order = TIEBREAKER.get(produkt_key, TIEBREAKER_DEFAULT)
    try:    return order.index(tag)
    except ValueError: return len(order)

# Tageslimit fuer aktive Arbeitsminuten (Soft-Cap).
# Tage, deren aktiv_min ueber diesem Wert liegen, werden in
# verteile_bestandssicher nur als letztes Mittel gewaehlt. Pflicht-Chargen
# (mit Deadline oder vorab gesetzte Mindestchargen) ignorieren das Limit.
MAX_AKTIV_MIN_PRO_TAG = 420  # = 7h. Anpassen wenn Arbeitstage laenger/kuerzer.

VERKAUF_ZU_CONFIG = {
    "Beeren Tartelette":"Beeren Tartelette","Bienenstich":"Bienenstich",
    "Donauwelle":"Donauwelle","Kaese-Rhabarber Schnitte":"Kaese-Rhabarber Schnitte",
    "Kaesekuchen":"Kaesekuchen","Lauch-Speck Quiche":"Lauch-Speck Quiche",
    "Mango-Passionsfrucht Tartelette":"Mango-Passionsfrucht Tartelette",
    "Pistazien Toertchen":"Pistazien Toertchen",
    "Käse-Rhabarber Schnitte":"Kaese-Rhabarber Schnitte",
    "Käsekuchen":"Kaesekuchen","Pistazien Törtchen":"Pistazien Toertchen",
}

MURBETEIG_PORTIONEN = {"Tartelette":40,"30x20 Schnitte":200,"Kuchenform":350,"Mürbeteig rund":180}
MURBETEIG_REZEPT    = {"Zucker":250,"Butter":500,"Salz":1,"Zitronenaroma":1,"Vanillearoma":2,"Vollei":100,"Weizenmehl":750}
MURBETEIG_GESAMT    = sum(MURBETEIG_REZEPT.values())
MURBETEIG_AUSROLL   = {"Tartelette":0.5,"30x20 Schnitte":1,"Kuchenform":2,"Mürbeteig rund":1}
MURBETEIG_ANBACK    = 9  # Minuten Anbackzeit

# Sub-Rezept-Zuordnung: Zutat-Name (lowercase) → Schluessel in sub_rezepte dict
SUB_REZEPT_MAPPING = {
    "salzmasse":                "Salzmasse",
    "dinkelhefeteig":           "Dinkelhefeteig",
    "sandmasse hell":           "Sandmasse hell",
    "sandmasse dunkel":         "Sandmasse dunkel",
    "deutsche buttercreme":     "Deutsche Buttercreme",
    "überzugsganache":          "Überzugsganache",
    "uberzugsganache":          "Überzugsganache",
    "amerikanische käsemasse":  "Amerikanische Käsemasse",
    "amerikanische kasemasse":  "Amerikanische Käsemasse",
    "streusel":                 "Streusel",
    "mandelsandmasse":          "Mandel Sandmasse",
    "mandel sandmasse":         "Mandel Sandmasse",
    "vanillecreme":             "Vanillecreme",
}

# Sub-Rezept Dateien
# Fallback-Chargengrenzen fuer Sub-Rezepte deren Datei nicht lesbar ist
# Werte werden nur verwendet wenn die Datei keinen lesbaren "Maximal"-Eintrag hat
SUB_REZEPT_MAX_FALLBACK = {
    "Amerikanische Käsemasse": {"max_charge_gramm": 5000},
}

SUB_REZEPT_DATEIEN = {
    # "Dinkelhefeteig"       → in Lauch-Speck Quiche integriert
    # "Amerikanische Käsemasse" → in Käse-Rhabarber Schnitte integriert
    # "Salzmasse"            → in Lauch-Speck Quiche integriert
    # "Deutsche Buttercreme" → in Donauwelle integriert
    # "Mandel Sandmasse"     → in Pistazien-Schoko Törtchen integriert
    # "Überzugsganache"      → in Donauwelle integriert
    "Sandmasse hell":       "Sandmasse hell.xlsx",
    "Sandmasse dunkel":     "Sandmasse dunkel.xlsx",
    "Streusel":             "Streusel.xlsx",
    "Vanillecreme":         "Vanillecreme.xlsx",
}

# Aufgaben-Prioritaeten fuer die Tagesreihenfolge
PRIO = {
    "CAFE2":      0,   # Excel only, nicht im HTML
    "MURT_BACK":  1,   # Muerbeteig backen (Dienstag)
    "MURT_ANSET": 2,   # Muerbeteig ansetzen (Montag)
    "PIST_PREP":  3,   # Pistazien Vorbereitung
    "KK":         10,  # Kaesekuchen
    "BIEN":       20,  # Bienenstich
    "QUICHE":     30,  # Quiche
    "DONA":       40,  # Donauwelle
    "KR":         50,  # Kaese-Rhabarber
    "STREUSEL":   5,   # Streusel (Montag fuer ganze Woche)
    "TART":       60,  # Tartelettes
    "PIST_MAIN":  70,  # Pistazien Hauptproduktion
}

# ============================================================
# ZIP/XML HELFER
# ============================================================
def lese_alle_zellen(pfad, sheet='xl/worksheets/sheet1.xml'):
    """Liest alle Zellen aus einer xlsx (inkl. SharedStrings). Funktioniert auch ohne Central Directory."""
    with open(pfad,'rb') as f: data=f.read()
    entries={}
    pos=0
    while pos<len(data)-30:
        if data[pos:pos+4]!=b'PK\x03\x04': pos+=1; continue
        comp =struct.unpack('<H',data[pos+8:pos+10])[0]
        csz  =struct.unpack('<I',data[pos+18:pos+22])[0]
        fnlen=struct.unpack('<H',data[pos+26:pos+28])[0]
        exlen=struct.unpack('<H',data[pos+28:pos+30])[0]
        fname=data[pos+30:pos+30+fnlen].decode('utf-8',errors='replace')
        dstart=pos+30+fnlen+exlen
        raw=data[dstart:dstart+csz]
        try: entries[fname]=zlib.decompress(raw,-15) if comp==8 else raw
        except: entries[fname]=b''
        pos=dstart+max(csz,1)

    shared=[]
    if 'xl/sharedStrings.xml' in entries and entries['xl/sharedStrings.xml'].strip():
        try:
            ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            root=ET.fromstring(entries['xl/sharedStrings.xml'])
            for si in root.findall(f'{{{ns}}}si'):
                t=si.find(f'{{{ns}}}t')
                if t is not None: shared.append(t.text or '')
                else: shared.append(''.join(p.text or '' for p in si.findall(f'.//{{{ns}}}t')))
        except: pass

    ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    root=ET.fromstring(entries.get(sheet,b'<root/>'))
    cols={c:i+1 for i,c in enumerate('ABCDEFGHIJKLMNOPQRSTUVWXYZ')}
    rows={}
    for row_el in root.iter(f'{{{ns}}}row'):
        r=int(row_el.attrib.get('r',0))
        rd={}
        for c_el in row_el.iter(f'{{{ns}}}c'):
            ref=c_el.attrib.get('r','')
            col=cols.get(''.join(ch for ch in ref if ch.isalpha()),0)
            t=c_el.attrib.get('t','')
            v_el=c_el.find(f'{{{ns}}}v')
            if v_el is not None and v_el.text is not None:
                try:
                    val=shared[int(v_el.text)] if t=='s' and int(v_el.text)<len(shared) else v_el.text
                except: val=v_el.text
                rd[col]=val
        if rd: rows[r]=rd
    return rows

# ============================================================
# REZEPT-PARSER
# ============================================================
SKIP_MUSTER = [
    'herstellung','haltbarkeit','backen:','portionsgröß','portionsgroeß',
    'herzustellend','grundrezept mit gleichem','1 grundrezept entspricht',
    'vorbereitung:','zeitaufwand',
]

def _parse_zeitaufwand(rows):
    """Liest Zeitaufwand-Felder (aktiv/passiv in Minuten) aus einem Zeilen-Dict."""
    aktiv = passiv = 0
    for r in sorted(rows):
        b = str(rows[r].get(2, '') or '').strip().lower()
        c = str(rows[r].get(3, '') or '').strip()
        if 'zeitaufwand aktiv' in b:
            m = re.search(r'(\d+)', c)
            if m: aktiv = int(m.group(1))
        elif 'zeitaufwand passiv' in b:
            m = re.search(r'(\d+)', c)
            if m: passiv = int(m.group(1))
    return aktiv, passiv

def _lese_zeitaufwand_openpyxl(pfad):
    """Liest Zeitaufwand direkt via openpyxl (liest SharedStrings korrekt).
    Fallback fuer den Fall, dass lese_alle_zellen keine Textzellen findet.
    """
    try:
        wb = openpyxl.load_workbook(str(pfad), data_only=True)
        ws = wb.active
        aktiv = passiv = 0
        for row in ws.iter_rows(values_only=True):
            b = str(row[1] or '').strip().lower()
            c = str(row[2] or '').strip()
            if 'zeitaufwand aktiv' in b:
                m = re.search(r'(\d+)', c)
                if m: aktiv = int(m.group(1))
            elif 'zeitaufwand passiv' in b:
                m = re.search(r'(\d+)', c)
                if m: passiv = int(m.group(1))
        return aktiv, passiv
    except Exception:
        return 0, 0

def _lade_rezept_zeilen(pfad):
    """Laedt Rezeptzeilen via openpyxl (primär) oder lese_alle_zellen (Fallback).
    Gibt Liste von (col_B, col_C) Strings zurueck — immer mit korrekten Texten.
    """
    # openpyxl liest SharedStrings korrekt — immer bevorzugen
    try:
        wb = openpyxl.load_workbook(str(pfad), data_only=True)
        ws = wb.active
        zeilen = []
        for row in ws.iter_rows(values_only=True):
            b = str(row[1] or '').strip() if len(row) > 1 else ''
            c = str(row[2] or '').strip() if len(row) > 2 else ''
            zeilen.append((b, c))
        return zeilen
    except Exception:
        pass
    # Fallback: custom ZIP-Parser fuer OneDrive-korrumpierte Dateien
    try:
        rows_dict = lese_alle_zellen(str(pfad))
        zeilen = []
        for r in sorted(rows_dict.keys()):
            row = rows_dict[r]
            b = str(row.get(2, '') or '').strip()
            c = str(row.get(3, '') or '').strip()
            zeilen.append((b, c))
        return zeilen
    except Exception:
        return []

def parse_rezept_datei(pfad):
    """Parst eine Rezeptdatei in strukturierte Daten."""
    zeilen = _lade_rezept_zeilen(pfad)
    if not zeilen:
        return None

    # Erste nicht-leere Zeile in Spalte B = Produktname
    name = Path(pfad).stem
    for b, c in zeilen:
        if b:
            name = b
            break

    max_charge_num  = 1
    max_charge_text = '1x'
    sections = []
    current_section = None
    pending_name    = None
    in_zutaten      = False

    def speichere_section():
        nonlocal current_section
        if current_section and current_section.get('zutaten'):
            sections.append(current_section)
        current_section = None

    for idx, (b, c) in enumerate(zeilen):
        if not b: in_zutaten=False; continue
        bl = b.lower()

        # Produktname ueberspringen
        if b == name: continue

        # Maximal-Menge
        if 'maximal' in bl and ('herzustellend' in bl or 'menge' in bl or 'gleichz' in bl):
            in_zutaten = False
            speichere_section()
            # Nächste nicht-leere Zeile nach diesem Index
            nxt = zeilen[idx+1][0] if idx+1 < len(zeilen) else ''
            max_charge_text = nxt
            m = re.match(r'(\d+)', nxt)
            if m: max_charge_num = int(m.group(1))
            continue

        # Andere Skip-Zeilen
        if any(p in bl for p in SKIP_MUSTER): in_zutaten=False; continue

        # Zutat/Menge-Kopfzeile
        if bl.rstrip(':') in ('zutat','zutaten') and c.lower().rstrip(':') in ('menge',''):
            speichere_section()
            current_section = {'name': pending_name or name, 'zutaten': []}
            pending_name = None; in_zutaten = True; continue

        # Zutatenzeile (col B = Name, col C = Zahl oder "Nx")
        if c:
            try:
                menge = float(c.replace(',','.'))
                if current_section is None:
                    current_section = {'name': pending_name or name, 'zutaten': []}
                    pending_name = None; in_zutaten = True
                current_section['zutaten'].append({'zutat':b,'menge':menge,'einheit':'g'})
                in_zutaten = True; continue
            except ValueError:
                if re.match(r'^\d*x$', c, re.I):
                    if current_section is None:
                        current_section = {'name': pending_name or name, 'zutaten': []}
                        pending_name = None; in_zutaten = True
                    current_section['zutaten'].append({'zutat':b,'menge':c,'einheit':'ref'})
                    in_zutaten = True; continue

        # Abschnittsueberschrift (nur col B)
        if in_zutaten:
            in_zutaten = False
            speichere_section()
        if not any(p in bl for p in SKIP_MUSTER):
            pending_name = b

    speichere_section()
    # Zeitaufwand: openpyxl liest SharedStrings korrekt (bereits über _lade_rezept_zeilen geladen)
    aktiv_min = passiv_min = 0
    for b, c in zeilen:
        bl = b.lower()
        if 'zeitaufwand aktiv' in bl:
            m2 = re.search(r'(\d+)', c)
            if m2: aktiv_min = int(m2.group(1))
        elif 'zeitaufwand passiv' in bl:
            m2 = re.search(r'(\d+)', c)
            if m2: passiv_min = int(m2.group(1))
    # Haltbarkeit parsen: Format "4 Tage (Kuehlung) / 3 Monate (Froster)"
    # Wir suchen den ersten Tages-Wert vor dem Stichwort "Kuehl"/"Kühl".
    haltbarkeit_kuehl_tage = None
    for b, c in zeilen:
        if 'haltbarkeit' not in b.lower(): continue
        txt = c or ''
        # 1) explizit: "X Tage (Kuehlung)" / "X Tage (Kühlung)"
        m = re.search(r'(\d+)\s*tage?\s*\(\s*k[uü]hl', txt, re.I)
        if m:
            haltbarkeit_kuehl_tage = int(m.group(1)); break
        # 2) Fallback: erster Tageswert ueberhaupt
        m = re.search(r'(\d+)\s*tage', txt, re.I)
        if m:
            haltbarkeit_kuehl_tage = int(m.group(1)); break
    return {'name':name,'max_charge_num':max_charge_num,'max_charge_text':max_charge_text,
            'sections':sections,'aktiv_min':aktiv_min,'passiv_min':passiv_min,
            'haltbarkeit_kuehl_tage':haltbarkeit_kuehl_tage}

def lade_alle_rezepte():
    """Laedt alle Rezeptdateien aus dem Rezepte-Ordner."""
    rezepte = {}
    if not REZEPTE_PFAD.exists(): return rezepte
    for key, cfg in PRODUKT_CONFIG.items():
        datei = cfg.get('rezept')
        if not datei: continue
        pfad = REZEPTE_PFAD / datei
        if not pfad.exists(): continue
        r = parse_rezept_datei(pfad)
        if r: rezepte[key] = r
    return rezepte

def parse_sub_rezept(pfad):
    """Parst eine einfache Sub-Rezept-Datei (Zutatenliste + Yield).
    Liest auch max_charge_num (Anzahl Grundrezepte) und max_charge_gramm (g-Limit).
    """
    zeilen = _lade_rezept_zeilen(pfad)
    name_val = Path(pfad).stem
    for b, c in zeilen:
        if b: name_val = b; break

    # Nicht-Zutaten-Zeilen die uebersprungen werden (kein break mehr)
    SKIP = ['herstellung', 'haltbarkeit', 'hinweis', 'zubereitung', 'vorbereitung']
    zutaten = []; yield_gramm = 0.0
    aktiv_min = passiv_min = 0; yield_info = None
    max_charge_num = 99; max_charge_gramm = None
    _lese_max_naechste = False  # True = naechste nicht-leere Zeile ist der Max-Wert

    for b, c in zeilen:
        if not b: continue
        bl = b.lower()
        # Zeitaufwand (immer zuerst pruefen, egal wo im Dokument)
        if 'zeitaufwand aktiv' in bl:
            m = re.search(r'(\d+)', c)
            if m: aktiv_min = int(m.group(1)); continue
        if 'zeitaufwand passiv' in bl:
            m = re.search(r'(\d+)', c)
            if m: passiv_min = int(m.group(1)); continue
        # Yield-Info (Dinkelhefeteig: "5 Böden")
        if 'böden' in bl or 'boden' in bl:
            m = re.search(r'(\d+)\s*[Bb][öo]den', b)
            if m: yield_info = int(m.group(1))
        # Maximale Charge: naechste nicht-leere Zeile nach "Maximal..."-Label lesen
        if _lese_max_naechste:
            _lese_max_naechste = False
            m_x = re.match(r'(\d+)\s*x', bl)   # "1x Grundrezept"
            if m_x: max_charge_num = int(m_x.group(1))
            m_kg = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:kg|kilo(?:gramm)?)\b', bl)
            if not m_kg:
                m_kg = re.search(r'ca\.?\s*(\d+(?:[.,]\d+)?)\s*(?:kg|kilo(?:gramm)?)\b', bl)
            if m_kg: max_charge_gramm = round(float(m_kg.group(1).replace(',', '.')) * 1000)
            elif re.search(r'(\d+)\s*g\b', bl):
                m_g = re.search(r'(\d+)\s*g\b', bl)
                if m_g: max_charge_gramm = int(m_g.group(1))
            continue
        if 'maximal' in bl:
            _lese_max_naechste = True; continue
        # Nicht-Zutaten-Zeilen ueberspringen (kein break - Zeitaufwand steht oft am Ende)
        if any(w in bl for w in SKIP): continue
        # Zutaten-Header
        if bl in ('zutat', 'rezept:', 'rezept', 'rezept :'): continue
        # Zutat parsen
        if c:
            try:
                menge = float(c.replace(',', '.'))
                zutaten.append({'zutat': b, 'menge': menge, 'einheit': 'g'})
                yield_gramm += menge
            except: pass

    return {'name': str(name_val), 'zutaten': zutaten, 'yield_gramm': yield_gramm,
            'yield_info': yield_info, 'aktiv_min': aktiv_min, 'passiv_min': passiv_min,
            'max_charge_num': max_charge_num, 'max_charge_gramm': max_charge_gramm}

def _init_murbeteig():
    """Laedt Muerb­eteig-Konfiguration aus Rezepte/Muer­beteig.xlsx
    und aktualisiert die globalen Konstanten MURBETEIG_*.
    Faellt auf die hardcodierten Standardwerte zurueck, wenn die Datei fehlt.
    """
    global MURBETEIG_PORTIONEN, MURBETEIG_REZEPT, MURBETEIG_GESAMT
    global MURBETEIG_AUSROLL, MURBETEIG_ANBACK
    pfad = REZEPTE_PFAD / "Mürbeteig.xlsx"
    if not pfad.exists():
        return
    try:
        zeilen = _lade_rezept_zeilen(pfad)
        modus   = 'zutaten'
        rezept  = {}
        portionen = {}
        ausroll   = {}
        anback    = MURBETEIG_ANBACK
        SKIP = {'maximal', 'haltbarkeit', 'hinweis', 'zutat', 'backform',
                'rezept', 'zeitaufwand aktiv', 'zeitaufwand passiv'}
        for b, c in zeilen:
            if not b:
                continue
            bl = b.lower().strip()
            # Abschnitts-Umschalter
            if 'portionsgröß' in bl or 'portionsgroeß' in bl:
                modus = 'portionen'; continue
            if 'ausrollzeit' in bl:
                modus = 'ausroll'; continue
            if 'anbackzeit' in bl:
                try: anback = int(float(str(c).replace(',', '.')))
                except: pass
                continue
            # Überschriften / irrelevante Zeilen überspringen
            if any(bl.startswith(s) for s in SKIP):
                continue
            if re.match(r'\d+\s*x$', bl):   # "1x"
                continue
            # Wert einlesen
            if c:
                try:
                    val = float(str(c).replace(',', '.'))
                    if modus == 'zutaten':
                        rezept[b] = val
                    elif modus == 'portionen':
                        portionen[b] = val
                    elif modus == 'ausroll':
                        ausroll[b] = val
                except (ValueError, TypeError):
                    pass
        if rezept:
            MURBETEIG_REZEPT  = rezept
            MURBETEIG_GESAMT  = sum(rezept.values())
        if portionen:
            MURBETEIG_PORTIONEN = portionen
        if ausroll:
            MURBETEIG_AUSROLL = ausroll
        MURBETEIG_ANBACK = anback
        print("  Mürbeteig aus xlsx geladen: {} Zutaten, {} Portionsformen".format(
            len(MURBETEIG_REZEPT), len(MURBETEIG_PORTIONEN)))
    except Exception as e:
        print("  Warnung: Mürbeteig.xlsx Ladefehler: {}".format(e))


def lade_sub_rezepte():
    """Laedt alle Sub-Rezept-Dateien."""
    sub = {}
    if not REZEPTE_PFAD.exists(): return sub
    for key, datei in SUB_REZEPT_DATEIEN.items():
        pfad = REZEPTE_PFAD / datei
        if not pfad.exists(): continue
        r = parse_sub_rezept(pfad)
        if r:
            # Fallback-Werte anwenden wenn Datei keinen Maximal-Eintrag hat
            fb = SUB_REZEPT_MAX_FALLBACK.get(key, {})
            if fb.get("max_charge_gramm") and not r.get("max_charge_gramm"):
                r["max_charge_gramm"] = fb["max_charge_gramm"]
            if fb.get("max_charge_num") and r.get("max_charge_num", 99) == 99:
                r["max_charge_num"] = fb["max_charge_num"]
            sub[key] = r
    return sub

def _find_sub_recipes(zutaten, sub_rezepte, scale_factor, is_per_piece, batch_size, already):
    """Findet Sub-Rezepte rekursiv mit korrekter Mengenskalierung.
    - is_per_piece=True: Zutatenmengen sind pro Stück → mit batch_size multiplizieren
    - is_per_piece=False: Zutatenmengen sind absolut → mit scale_factor multiplizieren
    - scale_factor: für Sub-Sub-Rezepte = n_mal des Eltern-Sub-Rezepts
    """
    results = []
    for z in zutaten:
        z_low = z['zutat'].lower()
        for muster, key in SUB_REZEPT_MAPPING.items():
            if muster in z_low and key not in already and key in sub_rezepte:
                already.add(key)
                srz = sub_rezepte[key]
                if z['einheit'] == 'ref':
                    n_mal = math.ceil(batch_size / srz['yield_info']) if srz.get('yield_info') else 1
                else:
                    menge_total = z['menge'] * (batch_size if is_per_piece else scale_factor)
                    yield_g = srz['yield_gramm']
                    n_mal = math.ceil(menge_total / yield_g) if yield_g > 0 else 1
                results.append((key, srz, n_mal))
                # Rekursiv in Sub-Sub-Rezepte, mit n_mal als neuer Skalierung
                results.extend(_find_sub_recipes(
                    srz['zutaten'], sub_rezepte, n_mal, False, batch_size, already))
                break
    return results

def berechne_zeitaufwand(produkt_key, rezepte, sub_rezepte, menge):
    """Berechnet Gesamtzeitaufwand (aktiv + passiv) fuer eine Produktion.
    Kaskadiert durch alle Sub-Rezepte (z.B. Donauwelle → Buttercreme → Vanillecreme).
    Gibt (aktiv_min, passiv_min) zurueck.
    """
    rzp = rezepte.get(produkt_key)
    if not rzp: return 0, 0
    max_c = rzp.get('max_charge_num', 1) or 1
    n_chargen = math.ceil(menge / max_c) if max_c > 0 else 1

    # Eigene Zeit: pro Charge × Anzahl Chargen
    aktiv  = rzp.get('aktiv_min', 0) * n_chargen
    passiv = rzp.get('passiv_min', 0)  # Passivzeit laeuft parallel → nur einmal zaehlen

    # Sub-Rezepte kaskadieren
    already = set()
    for i, sec in enumerate(rzp.get('sections', [])):
        is_per_piece = (i == 0)
        found = _find_sub_recipes(sec.get('zutaten', []), sub_rezepte,
                                  scale_factor=1, is_per_piece=is_per_piece,
                                  batch_size=menge, already=already)
        for key, srz, n_mal in found:
            aktiv  += srz.get('aktiv_min', 0) * n_mal
            passiv  = max(passiv, srz.get('passiv_min', 0))  # parallel
    return aktiv, passiv

def format_sub_rezept_html(key, srz, n_mal):
    """Erstellt HTML-Block fuer ein einzelnes Sub-Rezept.
    Beruecksichtigt max_charge_num und max_charge_gramm aus dem Rezept.
    """
    yield_g   = srz.get('yield_gramm', 0) or 1
    max_num   = srz.get('max_charge_num', 99)   # Max Grundrezepte gleichzeitig
    max_gram  = srz.get('max_charge_gramm')      # Max Gramm gleichzeitig (oder None)
    total_g   = yield_g * n_mal

    # Physikalische Chargen berechnen
    if max_gram and max_gram > 0:
        n_phys  = math.ceil(total_g / max_gram)
        g_je    = total_g / n_phys
    elif max_num < 99:
        n_phys  = math.ceil(n_mal / max_num)
        g_je    = yield_g * min(n_mal, max_num)
    else:
        n_phys  = 1
        g_je    = total_g

    # Titel
    if n_phys > 1:
        title = ('Sub-Rezept: {} &mdash; {:.0f}g gesamt &rarr; '
                 '<b>{}x Charge à je {:.0f}g</b>').format(
                     srz['name'], total_g, n_phys, g_je)
    elif n_mal > 1:
        title = 'Sub-Rezept: {} &mdash; {}x Grundrezept ({:.0f}g gesamt)'.format(
            srz['name'], n_mal, total_g)
    else:
        title = 'Sub-Rezept: {} &mdash; 1x Grundrezept ({:.0f}g)'.format(srz['name'], yield_g)

    # Zutaten: immer pro Charge anzeigen (nicht gesamt)
    skala = g_je / yield_g if yield_g > 0 else 1
    zhtml = ''.join(
        '<li><span class="zutat">{}</span><span class="menge">{:.0f}g</span></li>'.format(
            z['zutat'], z['menge'] * skala)
        for z in srz['zutaten']
    )
    suffix = ' <span style="color:#e67e22;font-size:11px">(je Charge)</span>' if n_phys > 1 else ''
    return (
        '<div class="sub-recipe">'
        '<div class="sub-recipe-title">{}{}</div>'
        '<ul class="zutaten">{}</ul>'
        '</div>'
    ).format(title, suffix, zhtml)

def format_alle_sub_rezepte_html(rezept, batch_size, sub_rezepte, skip_keys=None):
    """Zeigt alle Sub-Rezepte die für ein Produkt benötigt werden.
    Erste Sektion = Hauptzutaten (per Stück).
    Weitere Sektionen = eingebettete Rezeptabschnitte (absolute Mengen).
    skip_keys: Sub-Rezepte die bereits an einem anderen Tag hergestellt wurden.
    """
    if not sub_rezepte: return ''
    already = set(skip_keys or set())  # Vorbereitete Rezepte von vornherein überspringen
    all_found = []
    for i, sec in enumerate(rezept.get('sections', [])):
        is_per_piece = (i == 0)  # Nur erste Sektion ist pro Stück
        results = _find_sub_recipes(
            sec.get('zutaten', []), sub_rezepte,
            scale_factor=1, is_per_piece=is_per_piece,
            batch_size=batch_size, already=already)
        all_found.extend(results)
    html = ''
    for key, srz, n_mal in all_found:
        html += format_sub_rezept_html(key, srz, n_mal)
    return html

# ============================================================
# KONFIGURATION
# ============================================================
def lade_konfiguration():
    if not KONFIGURATION_DATEI.exists(): return
    try:
        # Versuche zuerst openpyxl (fuer gueltige ZIP-Dateien)
        try:
            wb = openpyxl.load_workbook(str(KONFIGURATION_DATEI), data_only=True)
            ws1 = wb['Ofenkapazitäten'] if 'Ofenkapazitäten' in wb.sheetnames else wb.worksheets[0]
            ws2 = wb['Produktionsregeln'] if 'Produktionsregeln' in wb.sheetnames else wb.worksheets[1]
            cfg_keys = list(PRODUKT_CONFIG.keys())
            for i, key in enumerate(cfg_keys):
                val = ws1.cell(row=i+4, column=2).value
                if val is not None:
                    try:
                        v = int(float(str(val)))
                        if 0 < v < 1000:
                            PRODUKT_CONFIG[key]["max_ofen"] = v
                            if key not in ("Beeren Tartelette","Mango-Passionsfrucht Tartelette","Pistazien Toertchen","Lauch-Speck Quiche"):
                                PRODUKT_CONFIG[key]["min_charge"] = v
                    except: pass
            schwellen_keys=["Beeren Tartelette","Mango-Passionsfrucht Tartelette","Pistazien Toertchen"]
            for i, key in enumerate(schwellen_keys):
                val = ws2.cell(row=i+4, column=4).value
                if val is not None:
                    try: FROSTER_SCHWELLEN[key] = int(float(str(val)))
                    except: pass
            # C2-Puffer aus Spalte 6, Zeilen 12-16 (frische Produkte)
            frisch_keys=["Donauwelle","Kaese-Rhabarber Schnitte","Kaesekuchen","Bienenstich","Lauch-Speck Quiche"]
            for i, key in enumerate(frisch_keys):
                val = ws2.cell(row=i+12, column=6).value
                if val is not None:
                    try: C2_PUFFER[key] = int(float(str(val)))
                    except: pass
        except Exception:
            # Fallback: custom parser fuer OneDrive-korrumpierte Dateien
            sheet1 = lese_alle_zellen(str(KONFIGURATION_DATEI), 'xl/worksheets/sheet1.xml')
            cfg_keys = list(PRODUKT_CONFIG.keys())
            for i, key in enumerate(cfg_keys):
                val = sheet1.get(i+4, {}).get(2)
                if val:
                    try:
                        v = int(float(str(val)))
                        if 0 < v < 1000:
                            PRODUKT_CONFIG[key]["max_ofen"] = v
                            if key not in ("Beeren Tartelette","Mango-Passionsfrucht Tartelette","Pistazien Toertchen","Lauch-Speck Quiche"):
                                PRODUKT_CONFIG[key]["min_charge"] = v
                    except: pass
            sheet2 = lese_alle_zellen(str(KONFIGURATION_DATEI), 'xl/worksheets/sheet2.xml')
            schwellen_keys=["Beeren Tartelette","Mango-Passionsfrucht Tartelette","Pistazien Toertchen"]
            for i, key in enumerate(schwellen_keys):
                val = sheet2.get(i+4, {}).get(4)
                if val:
                    try: FROSTER_SCHWELLEN[key] = int(float(str(val)))
                    except: pass
            frisch_keys=["Donauwelle","Kaese-Rhabarber Schnitte","Kaesekuchen","Bienenstich","Lauch-Speck Quiche"]
            for i, key in enumerate(frisch_keys):
                val = sheet2.get(i+12, {}).get(6)
                if val:
                    try: C2_PUFFER[key] = int(float(str(val)))
                    except: pass
        print("Konfiguration: Ofen={}, Schwellen={}, C2-Puffer={}".format(
            {k:v['max_ofen'] for k,v in PRODUKT_CONFIG.items()}, FROSTER_SCHWELLEN, C2_PUFFER))
    except Exception as e:
        print("Warnung Konfiguration: {}".format(e))

# ============================================================
# DATEN LESEN
# ============================================================
def lese_verkaufszahlen():
    try: wb=openpyxl.load_workbook(str(VERKAUFSZAHLEN_DATEI),data_only=True)
    except:
        for sfx in ['1','2','_backup']:
            alt=BASE/"Verkaufszahlen{}.xlsx".format(sfx)
            if alt.exists():
                try: wb=openpyxl.load_workbook(str(alt),data_only=True); break
                except: pass
        else: return {}
    verkauf={}
    for cname in ["Cafe 1","Cafe 2","Café 1","Café 2"]:
        if cname not in wb.sheetnames: continue
        ws=wb[cname]; ck="Cafe 1" if "1" in cname else "Cafe 2"
        verkauf[ck]={}
        for row in ws.iter_rows(min_row=4,values_only=True):
            if not row[0] or row[0]=="Produkt": continue
            pk=VERKAUF_ZU_CONFIG.get(str(row[0]).strip())
            if not pk: continue
            verkauf[ck][pk]={t:float(row[i+1]) if row[i+1] else 0.0 for i,t in enumerate(TAGE)}
    # Forecast aus Historie ueberlagern (wenn Daten vorhanden)
    verkauf = wende_forecast_an(verkauf)
    return verkauf

# ============================================================
# FORECAST (EWMA-basiert)
# ============================================================
def lese_verkaufshistorie():
    """Laedt Verkaufshistorie.json, gibt Dict {wochen:[...], _alpha:0.3} zurueck.
    Robust gegen OneDrive-Korruption (trailing Null-Bytes)."""
    import json as _json
    if not VERKAUFSHISTORIE.exists():
        return {"wochen": [], "_alpha": 0.3}
    try:
        raw = VERKAUFSHISTORIE.read_text(encoding='utf-8', errors='replace')
        raw = raw.rstrip('\x00 \t\r\n')
        return _json.loads(raw)
    except Exception as e:
        print("  Warnung Verkaufshistorie: {} — ignoriere".format(e))
        return {"wochen": [], "_alpha": 0.3}

def berechne_forecast(historie, fallback_verkauf):
    """EWMA-Forecast aus Historie.

    Ergebnis-Format ist identisch zu lese_verkaufszahlen():
      {cafe: {produkt: {tag: stk}}}

    Algorithmus:
      1. Aus Historie pro (cafe, produkt) die Wochensummen extrahieren.
      2. EWMA mit alpha=0.3 ueber Wochensummen rechnen.
      3. Tagesprofil (Anteil pro Wochentag) aus fallback_verkauf uebernehmen.
      4. EWMA-Wochensumme * Tagesanteil = neuer Tagesvertikbet.

    Wenn historie leer ist, wird fallback_verkauf unveraendert zurueckgegeben.
    """
    wochen = historie.get("wochen", [])
    if not wochen:
        return fallback_verkauf  # Bootstrap-Phase

    alpha = float(historie.get("_alpha", 0.3))
    forecast = {}

    # Sammle alle (cafe, produkt) Kombinationen aus Fallback + Historie
    cafes = set(fallback_verkauf.keys())
    produkte = set()
    for cf, pmap in fallback_verkauf.items():
        produkte.update(pmap.keys())
    for w in wochen:
        for cf in w.get("verkauf", {}):
            cafes.add(cf)
            produkte.update(w["verkauf"][cf].keys())

    for cf in cafes:
        forecast[cf] = {}
        for pk in produkte:
            # Wochensumme aus Fallback als Startwert
            fb_profil = fallback_verkauf.get(cf, {}).get(pk, {t:0.0 for t in TAGE})
            fb_summe = sum(fb_profil.values())
            ewma = fb_summe  # Startwert
            # EWMA chronologisch durchlaufen
            sortierte = sorted(wochen, key=lambda w: w.get("kw",""))
            for w in sortierte:
                ist = w.get("verkauf", {}).get(cf, {}).get(pk)
                if ist is None: continue
                ewma = alpha * float(ist) + (1 - alpha) * ewma
            # Tagesprofil aus Fallback uebernehmen (Anteil pro Tag)
            if fb_summe > 0:
                forecast[cf][pk] = {t: ewma * fb_profil.get(t,0)/fb_summe for t in TAGE}
            else:
                # Kein Fallback-Profil: gleichmaessig auf ARBEITSTAGE
                forecast[cf][pk] = {t: (ewma/len(ARBEITSTAGE) if t in ARBEITSTAGE else 0.0) for t in TAGE}
    return forecast

def wende_forecast_an(fallback_verkauf):
    """Liefert verkauf-Dict mit angewandtem Forecast (oder Fallback wenn Historie leer)."""
    hist = lese_verkaufshistorie()
    n_wochen = len(hist.get("wochen", []))
    if n_wochen == 0:
        print("  Forecast: keine Historie — nutze statische Verkaufszahlen1.xlsx")
        return fallback_verkauf
    print("  Forecast: aktiv (EWMA alpha={} ueber {} Wochen Historie)".format(
        hist.get("_alpha", 0.3), n_wochen))
    return berechne_forecast(hist, fallback_verkauf)

def lese_lagerbestand():
    """Liest Lagerbestand — JSON hat Vorrang (Webformular), xlsx als Fallback."""
    import json as _json
    json_datei = BASE / "lagerbestand.json"
    if json_datei.exists():
        try:
            data = _json.loads(json_datei.read_text(encoding='utf-8'))
            lager = {}
            # JSON-Schluessel direkt auf Produkt-Keys mappen
            JSON_KEYS = {
                "Beeren Tartelette":              "Beeren Tartelette",
                "Bienenstich":                    "Bienenstich",
                "Donauwelle":                     "Donauwelle",
                "Kaese-Rhabarber Schnitte":       "Kaese-Rhabarber Schnitte",
                "Kaesekuchen":                    "Kaesekuchen",
                "Lauch-Speck Quiche":             "Lauch-Speck Quiche",
                "Mango-Passionsfrucht Tartelette":"Mango-Passionsfrucht Tartelette",
                "Pistazien Toertchen":            "Pistazien Toertchen",
            }
            for jk, pk in JSON_KEYS.items():
                lager[pk] = float(data.get(jk, 0))
            nt = data.get("naechste_tartelette", "")
            if "Beeren" in str(nt):
                lager['_naechste_tartelette'] = 'Beeren Tartelette'
            elif "Mango" in str(nt):
                lager['_naechste_tartelette'] = 'Mango-Passionsfrucht Tartelette'
            else:
                lager['_naechste_tartelette'] = None
            lager['_plan_ab_morgen'] = bool(data.get("plan_ab_morgen", False))
            for key in PRODUKT_CONFIG:
                lager.setdefault(key, 0.0)
            print("  Lagerbestand aus JSON gelesen.")
            return lager
        except Exception as e:
            print("  Warnung JSON-Lager: {} — lade xlsx".format(e))
    # Fallback: xlsx
    if not LAGERBESTAND_DATEI.exists():
        return {k:0.0 for k in PRODUKT_CONFIG}
    rows=lese_alle_zellen(str(LAGERBESTAND_DATEI))
    ZEILEN={4:"Beeren Tartelette",5:"Bienenstich",6:"Donauwelle",
            7:"Kaese-Rhabarber Schnitte",8:"Kaesekuchen",9:"Lauch-Speck Quiche",
            10:"Mango-Passionsfrucht Tartelette",11:"Pistazien Toertchen"}
    lager={}
    for z,prod in ZEILEN.items():
        val=rows.get(z,{}).get(2)
        try: lager[prod]=float(str(val)) if val else 0.0
        except: lager[prod]=0.0
    b14=rows.get(14,{}).get(2); a4=rows.get(4,{}).get(1); a10=rows.get(10,{}).get(1)
    lager['_naechste_tartelette']=None
    if b14 and a4 and str(b14)==str(a4): lager['_naechste_tartelette']='Beeren Tartelette'
    elif b14 and a10 and str(b14)==str(a10): lager['_naechste_tartelette']='Mango-Passionsfrucht Tartelette'
    for key in PRODUKT_CONFIG:
        if key not in lager: lager[key]=0.0
    return lager

# ============================================================
# PLANUNGSLOGIK
# ============================================================
def tartelette_diese_woche(lager):
    b=lager.get("Beeren Tartelette",0); m=lager.get("Mango-Passionsfrucht Tartelette",0)
    bs=FROSTER_SCHWELLEN["Beeren Tartelette"]; ms=FROSTER_SCHWELLEN["Mango-Passionsfrucht Tartelette"]
    res={"Beeren Tartelette":0,"Mango-Passionsfrucht Tartelette":0,"grund":""}
    if b>=bs and m>=ms:
        res["grund"]="Beide ueber Schwelle - keine Produktion"; return res
    if b<bs and m<ms:
        if b<=m: res["Beeren Tartelette"]=60;  res["grund"]="Beide unter Schwelle - Beeren ({}<{})".format(b,bs)
        else:    res["Mango-Passionsfrucht Tartelette"]=60; res["grund"]="Beide unter Schwelle - Mango ({}<{})".format(m,ms)
        return res
    if b<bs: res["Beeren Tartelette"]=60; res["grund"]="Beeren unter Schwelle ({}<{})".format(b,bs); return res
    res["Mango-Passionsfrucht Tartelette"]=60; res["grund"]="Mango unter Schwelle ({}<{})".format(m,ms); return res

def berechne_wochenbedarf(verkauf, lager, cafe2_erledigt=False):
    bedarf={}; t_plan=tartelette_diese_woche(lager)
    for key,cfg in PRODUKT_CONFIG.items():
        c1=sum(verkauf.get("Cafe 1",{}).get(key,{}).get(t,0) for t in TAGE)
        c2=sum(verkauf.get("Cafe 2",{}).get(key,{}).get(t,0) for t in TAGE) if cfg["cafe2"] else 0
        # Cafe 2 wurde bereits abgeholt: nur Cafe-1-Bedarf fuer Produktion relevant
        gesamt = c1 if cafe2_erledigt else c1+c2
        stock=lager.get(key,0)
        if key in ("Beeren Tartelette","Mango-Passionsfrucht Tartelette"):
            prod=t_plan.get(key,0); grund=t_plan.get("grund","")
        elif key=="Pistazien Toertchen":
            sw=FROSTER_SCHWELLEN[key]; prod=60 if stock<sw else 0
            grund="Produzieren ({}<{})".format(int(stock),sw) if prod else "Lager OK ({}>={})".format(int(stock),sw)
        else:
            # Dynamischer Soll-Stand: C1-Puffer (So+Mo aus Verkaufszahlen) + C2-Puffer aus Konfiguration
            tv_so = verkauf.get("Cafe 1",{}).get(key,{}).get("Sonntag",0)
            tv_mo = verkauf.get("Cafe 1",{}).get(key,{}).get("Montag",0)
            c1_puffer = math.ceil(tv_so + tv_mo)
            c2_puffer_val = C2_PUFFER.get(key, 4)
            soll_stand = c1_puffer + c2_puffer_val
            # Ziel: Cafe-1-Wochenbedarf abdecken UND Soll-Stand im Froster erreichen
            target = gesamt + soll_stand
            netto = max(0, target - stock)
            mc = cfg["min_charge"]
            prod = math.ceil(netto/mc)*mc if netto>0 else 0
            grund = "Soll-Stand: {} (C1-Puffer So+Mo={}+{}={}, C2-Puffer={})".format(
                soll_stand, int(tv_so), int(tv_mo), c1_puffer, c2_puffer_val)
        cafe2_lief=0 if cafe2_erledigt else (min(math.ceil(c2),int(stock)) if cfg.get("cafe2") else 0)
        soll_stand_val = soll_stand if key not in ("Beeren Tartelette","Mango-Passionsfrucht Tartelette","Pistazien Toertchen") else FROSTER_SCHWELLEN.get(key)
        # Tagesverbrauchsprofil: pro Arbeitstag den Cafe-1-Schnitt (Verkaufszahlen.xlsx).
        # Fallback fuer Tage mit 0: Wochenmittel/7 — damit der Rolling-Inventory-Check
        # nicht plötzlich Tage als "leer" markiert, wenn der Verkaufsplan luckenhaft ist.
        tv_profil = {}
        c1_tage = verkauf.get("Cafe 1",{}).get(key,{}) or {}
        woche_mittel = gesamt/7.0 if gesamt > 0 else 0.0
        for _t in ARBEITSTAGE:
            tv_profil[_t] = float(c1_tage.get(_t, 0)) or woche_mittel
        bedarf[key]={"cafe1_woche":round(c1,1),"cafe2_woche":round(c2,1),"cafe2_erledigt":cafe2_erledigt,
                     "gesamt_woche":round(gesamt,1),"im_lager":stock,
                     "netto_bedarf":round(max(0,gesamt-stock),1),
                     "zu_produzieren":prod,"cafe2_lieferung":cafe2_lief,"grund":grund,
                     "soll_stand":soll_stand_val,"tv_profil":tv_profil}
    return bedarf

def berechne_murbeteig(bedarf, lager=None):
    """Berechnet Muerb­eteig-Bedarf fuer die Woche.
    Zieht vorhandene Boeden und rohen Teig aus lager ab.
    """
    lager = lager or {}
    portionen={}; gesamt_g=0
    for key,b in bedarf.items():
        menge=b["zu_produzieren"]
        if not menge: continue
        murt=PRODUKT_CONFIG[key].get("murbeteig_typ")
        if not murt: continue
        g=menge*MURBETEIG_PORTIONEN.get(murt,0)
        if murt not in portionen: portionen[murt]={"anzahl":0,"gramm":0}
        portionen[murt]["anzahl"]+=menge; portionen[murt]["gramm"]+=g; gesamt_g+=g

    # ── Vorhandene gebackene Boeden abziehen ─────────────────────
    lager_map = {
        "Tartelette":    lager.get("Muerbeteig_Tartelette", 0),
        "30x20 Schnitte":lager.get("Muerbeteig_Schnitte",   0),
        "Kuchenform":    lager.get("Muerbeteig_Kuchen",      0),
        "Mürbeteig rund":lager.get("Muerbeteig_Rund",        0),
    }
    for form, vorrat in lager_map.items():
        if vorrat <= 0: continue
        g_vorrat = vorrat * MURBETEIG_PORTIONEN.get(form, 0)
        gesamt_g = max(0, gesamt_g - g_vorrat)
        if form in portionen:
            abzug = min(vorrat, portionen[form]["anzahl"])
            portionen[form]["anzahl"] -= abzug
            portionen[form]["gramm"]  -= abzug * MURBETEIG_PORTIONEN.get(form, 0)
            if portionen[form]["anzahl"] <= 0:
                del portionen[form]

    # ── Rohen Teig im Kuehlschrank abziehen ─────────────────────
    chargen_vorrat = lager.get("Muerbeteig_Chargen", 0)
    if chargen_vorrat > 0 and MURBETEIG_GESAMT:
        g_vorrat = chargen_vorrat * MURBETEIG_GESAMT
        gesamt_g = max(0, gesamt_g - g_vorrat)

    gr=math.ceil(gesamt_g/MURBETEIG_GESAMT) if MURBETEIG_GESAMT and gesamt_g > 0 else 0
    return {"portionen":portionen,"gesamt_gramm":gesamt_g,"grundrezepte":gr,
            "skalierung":gesamt_g/MURBETEIG_GESAMT if MURBETEIG_GESAMT else 0,
            "kk_montag_reserve":0}

# ============================================================
# ROLLING INVENTORY (Bestandsprüfung)
# ============================================================
def berechne_rolling_inventory(bedarf, lager, wplan):
    """Berechnet den rollierenden Lagerbestand fuer jeden Arbeitstag.
    Gibt (rolling_stock, tages_info, engpaesse) zurueck:
      rolling_stock[tag][key] = Bestand am Tagesende
      tages_info[tag]         = Liste von Eintraegen mit Status (niedrig/leer)
      engpaesse               = Liste von (tag, key, bestand) wo bestand < 0
    """
    from collections import defaultdict as _dd
    # Tagesproduktion aus Wochenplan ableiten
    produktion = {t: _dd(int) for t in ARBEITSTAGE}
    for tag in ARBEITSTAGE:
        for a in wplan.get(tag, []):
            pk = a.get("produkt_key")
            if pk and a.get("menge", 0) > 0 and not a.get("excel_only"):
                produktion[tag][pk] += int(a["menge"])

    prev = {key: float(lager.get(key, 0)) for key in PRODUKT_CONFIG}
    rolling = {}
    tages_info = {t: [] for t in ARBEITSTAGE}
    engpaesse = []

    for tag in ARBEITSTAGE:
        heute = {}
        for key in PRODUKT_CONFIG:
            bd = bedarf.get(key, {})
            # Tagesverbrauch aus Profil (Verkaufszahlen pro Wochentag), Fallback Wochenmittel
            tv_profil = bd.get("tv_profil") or {}
            tv = tv_profil.get(tag, bd.get("gesamt_woche", 0.0) / 7.0)
            prod = produktion[tag].get(key, 0)
            bestand = prev[key] + prod - tv
            heute[key] = round(bestand, 1)
            # Warnung wenn Bestand < 2 Tagesportionen ODER negativ
            if tv > 0 and bestand < tv * 2:
                status = "leer" if bestand < 0 else "niedrig"
                tages_info[tag].append({
                    "produkt": key,
                    "bestand": round(bestand, 1),
                    "tagesverbrauch": round(tv, 1),
                    "status": status,
                    "produziert": prod,
                })
                if bestand < 0:
                    engpaesse.append((tag, key, round(bestand, 1)))
        rolling[tag] = heute
        prev = heute

    return rolling, tages_info, engpaesse

# ============================================================
# HILFSFUNKTION: gleichmaessige Tagesverteilung
# ============================================================
def verteile_produktionstage(n_batches, verfuegbare_tage):
    """Verteilt n_batches Chargen gleichmaessig ueber verfuegbare_tage.
    Der letzte Tag wird immer belegt (Wochenend-Frische fuer Cafe 1).
    Beispiel: 2 Chargen in [Di,Mi,Do,Fr] -> [Di, Fr]
              3 Chargen in [Di,Mi,Do,Fr] -> [Di, Do, Fr]
    """
    if n_batches <= 0 or not verfuegbare_tage:
        return []
    if n_batches >= len(verfuegbare_tage):
        return list(verfuegbare_tage)
    if n_batches == 1:
        return [verfuegbare_tage[-1]]
    step = (len(verfuegbare_tage) - 1) / (n_batches - 1)
    indices = sorted(set(round(i * step) for i in range(n_batches)))
    # Sicherheit: falls Rundung zu wenige Indizes liefert, auffuellen
    while len(indices) < n_batches:
        for j in range(len(indices) - 1):
            if indices[j+1] - indices[j] > 1:
                indices.insert(j+1, indices[j] + 1)
                break
    return [verfuegbare_tage[i] for i in indices[:n_batches]]

# ============================================================
# BESTANDSSICHERES + ZEITOPTIMALES SCHEDULING
# ============================================================
def verteile_bestandssicher(produkt_key, bedarf, aktiv_plan, verfuegbare_tage, per_charge):
    """Verteilt Produktionschargen bestandssicher UND zeitoptimal.

    Fuer jede Charge:
      1. Simuliert den rollierenden Bestand — findet den spaetesten Tag,
         an dem produziert werden MUSS (Deadline: erster Tag mit Bestand < 0).
      2. Unter den gueltigen Kandidaten (vor/bis Deadline) wird der Tag
         mit der geringsten bereits geplanten aktiven Arbeitszeit gewaehlt.

    produkt_key    : Schluessel in bedarf
    bedarf         : Wochenbedarf-Dict (enthaelt im_lager, gesamt_woche, zu_produzieren)
    aktiv_plan     : {tag: aktiv_min_bisher} — wird NICHT veraendert
    verfuegbare_tage: geordnete Liste erlaubter Produktionstage
    per_charge     : Stueckzahl pro Charge (= max_ofen)
    Gibt: Liste von Tagen (evtl. Wiederholungen) zurueck.
    """
    from collections import defaultdict as _dd
    bd = bedarf.get(produkt_key, {})
    n_total    = bd.get("zu_produzieren", 0)
    tv_profil  = bd.get("tv_profil") or {}
    tv_avg     = bd.get("gesamt_woche", 0.0) / 7.0  # Fallback / Pre-Tage
    def _tv(tagname):
        # Profil-Wert pro Tag; Fallback Wochenmittel falls Tag fehlt
        return tv_profil.get(tagname, tv_avg)
    start_lager = float(bd.get("im_lager", 0))

    if n_total <= 0 or not verfuegbare_tage:
        return []
    if tv_avg <= 0:
        # Kein laufender Verbrauch: einfach gleichmaessig verteilen
        return verteile_produktionstage(math.ceil(n_total / per_charge), verfuegbare_tage)

    n_chargen = math.ceil(n_total / per_charge)
    tage = list(verfuegbare_tage)
    n = len(tage)

    # Verbrauch an Tagen VOR dem ersten verfuegbaren Tag beruecksichtigen.
    # z.B. Donauwelle startet erst Dienstag → Montag-Verbrauch wurde schon abgezogen.
    pre_tage = [t for t in ARBEITSTAGE if t != tage[0] and ARBEITSTAGE.index(t) < ARBEITSTAGE.index(tage[0])]
    start_lager = start_lager - sum(_tv(t) for t in pre_tage)

    placed = {}     # {tag_index: menge_produziert}
    result  = []

    for _ in range(n_chargen):
        # --- Rollierenden Bestand mit bisher platzierten Chargen simulieren ---
        s = start_lager
        stock_end = []
        for i, t in enumerate(tage):
            s += placed.get(i, 0)
            s -= _tv(t)
            stock_end.append(s)

        # --- Deadline: erster Tag an dem Bestand ohne weitere Produktion < 0 waere ---
        deadline_idx = None
        for i in range(n):
            if i not in placed and stock_end[i] < 0:
                deadline_idx = i
                break

        # --- Kandidaten: alle noch nicht belegten Tage bis inkl. Deadline ---
        if deadline_idx is not None:
            kandidaten = [i for i in range(n) if i not in placed and i <= deadline_idx]
        else:
            kandidaten = [i for i in range(n) if i not in placed]

        if not kandidaten:
            # Notfall: irgendeinen freien Tag nehmen
            kandidaten = [i for i in range(n) if i not in placed]
        if not kandidaten:
            break  # alle Tage belegt — mehr Chargen als Tage

        # --- Zeitoptimal + Tiebreaker laut Planungsregeln ---
        # Berechne pro Kandidat die hypothetische neue Tageslast nach Platzierung
        # dieser Charge. Tage, die das Limit MAX_AKTIV_MIN_PRO_TAG ueberschreiten
        # wuerden, werden als letztes gewaehlt — es sei denn, alle Kandidaten
        # ueberschreiten (dann gewinnt der mit dem kleinsten Ueberlauf).
        # Charge-Aktivzeit aus per_charge schaetzen: wir kennen sie nicht direkt,
        # aber aktiv_plan wird vom Aufrufer aktualisiert; hier reicht es, den
        # bisherigen aktiv_plan-Wert gegen MAX zu vergleichen.
        def _sort_key(i):
            tag = tage[i]
            cur = aktiv_plan.get(tag, 0)
            overload = 1 if cur >= MAX_AKTIV_MIN_PRO_TAG else 0
            return (overload, cur, tiebreaker_rang(produkt_key, tag))
        best_i = min(kandidaten, key=_sort_key)
        placed[best_i] = per_charge
        # notwendig = True wenn deadline_idx gesetzt war (Lager waere sonst negativ geworden)
        result.append((tage[best_i], deadline_idx is not None))

    return result

# ============================================================
# WOCHENPLAN-ERSTELLER
# ============================================================
def erstelle_wochenplan(bedarf, murt, rezepte=None, sub_rezepte=None, verbleibende_tage=None):
    plan={t:[] for t in ARBEITSTAGE}
    rezepte = rezepte or {}
    sub_rezepte = sub_rezepte or {}
    aktiv_plan = defaultdict(int)  # {tag: geplante_aktiv_min} fuer Lastverteilung
    # Verbleibende Tage: nur diese werden beplant
    vtage = list(verbleibende_tage) if verbleibende_tage else ab_heute_tage()

    # ── Fixe Tagesaufgaben vorab in aktiv_plan einkalkulieren ───────────────────
    # Der Scheduler (verteile_bestandssicher) waehlt Tage nach niedrigster
    # Arbeitszeit. Er muss daher die fixen Pflichtaufgaben kennen, BEVOR er
    # flexible Produkte verteilt — sonst wirken alle Tage gleich frei.
    _murt_gr  = (murt or {}).get("grundrezepte", 0)
    _murt_srz = (sub_rezepte or {}).get("Mürbeteig") or (sub_rezepte or {}).get("Murbeteig")
    _murt_min = (_murt_srz.get("aktiv_min", 20) if _murt_srz else 20)

    # Montag: Mürbeteig ansetzen + Streusel + Lager-Overhead
    if "Montag" in vtage:
        if _murt_gr > 0:
            aktiv_plan["Montag"] += _murt_min * _murt_gr
        _kr_prod = (bedarf or {}).get("Kaese-Rhabarber Schnitte", {}).get("zu_produzieren", 0)
        if _kr_prod > 0:
            _streusel_gr  = math.ceil(_kr_prod * 175 / 1925)
            _streusel_srz = (sub_rezepte or {}).get("Streusel", {})
            aktiv_plan["Montag"] += (_streusel_srz.get("aktiv_min", 10) or 10) * _streusel_gr
        aktiv_plan["Montag"] += 15   # Lager-Check / Vorbereitung

    # Dienstag: Mürbeteig backen (alle Boeden fuer die Woche)
    if "Dienstag" in vtage and _murt_gr > 0:
        aktiv_plan["Dienstag"] += _murt_min * _murt_gr

    # Mittwoch: Pistazien Vorbereitung (falls diese Woche) + Beeren Tartelette
    if "Mittwoch" in vtage:
        if (bedarf or {}).get("Pistazien Toertchen", {}).get("zu_produzieren", 0) > 0:
            aktiv_plan["Mittwoch"] += 45   # Brownieboden + Ganache ansetzen
        _bt_prod = (bedarf or {}).get("Beeren Tartelette", {}).get("zu_produzieren", 0)
        if _bt_prod > 0:
            _bt_a, _ = berechne_zeitaufwand("Beeren Tartelette", rezepte or {}, sub_rezepte or {}, _bt_prod)
            aktiv_plan["Mittwoch"] += _bt_a

    # Donnerstag: Pistazien Hauptproduktion + Mango-Passionsfrucht Tartelette
    if "Donnerstag" in vtage:
        _pt_prod = (bedarf or {}).get("Pistazien Toertchen", {}).get("zu_produzieren", 0)
        if _pt_prod > 0:
            _pt_a, _ = berechne_zeitaufwand("Pistazien Toertchen", rezepte or {}, sub_rezepte or {}, 60)
            aktiv_plan["Donnerstag"] += _pt_a
        _mt_prod = (bedarf or {}).get("Mango-Passionsfrucht Tartelette", {}).get("zu_produzieren", 0)
        if _mt_prod > 0:
            _mt_a, _ = berechne_zeitaufwand("Mango-Passionsfrucht Tartelette", rezepte or {}, sub_rezepte or {}, _mt_prod)
            aktiv_plan["Donnerstag"] += _mt_a
    def _filter(tage_liste):
        """Filtert eine Tagesliste auf die verbleibenden Tage."""
        return [t for t in tage_liste if t in vtage]

    def add(tag,produkt,menge,notiz="",prio=10,excel_only=False,produkt_key=None,aktiv_min=0,passiv_min=0,notwendig=None,grund=None):
        # grund: "pflicht" (Mindestcharge), "bestand" (Lager-Deadline),
        #        "frische" (Wochenend-Frische), "puffer" (Soll-Stand aufbauen)
        if grund is None:
            if notwendig is True:  grund = "bestand"
            elif notwendig is False: grund = "puffer"
        plan[tag].append({"produkt":produkt,"menge":menge,"notiz":notiz,
                          "prio":prio,"excel_only":excel_only,
                          "produkt_key":produkt_key or produkt,
                          "aktiv_min":aktiv_min,"passiv_min":passiv_min,
                          "notwendig":notwendig,"grund":grund})
        if not excel_only and aktiv_min: aktiv_plan[tag] += aktiv_min
    def info(tag,produkt,notiz="",prio=10,excel_only=False,aktiv_min=0,passiv_min=0,notwendig=None):
        plan[tag].append({"produkt":produkt,"menge":0,"notiz":notiz,
                          "prio":prio,"excel_only":excel_only,"produkt_key":None,
                          "aktiv_min":aktiv_min,"passiv_min":passiv_min,
                          "notwendig":notwendig})
        if not excel_only and aktiv_min: aktiv_plan[tag] += aktiv_min

    # --- CAFE 2 LIEFERUNG (nur Excel) ---
    lief_teile=["{}x {}".format(b["cafe2_lieferung"],k)
                for k,b in bedarf.items()
                if b.get("cafe2_lieferung",0)>0 and PRODUKT_CONFIG[k]["cafe2"]]
    lief_text=", ".join(lief_teile) if lief_teile else "(Lager leer)"
    # Cafe-2-Lieferung nur anzeigen wenn noch nicht erledigt UND Dienstag noch kommen
    _cafe2_done = any(b.get("cafe2_erledigt") for b in bedarf.values())
    if "Dienstag" in vtage and not _cafe2_done:
        info("Dienstag","CAFE 2 LIEFERUNG - aus Froster holen",lief_text,prio=PRIO["CAFE2"],excel_only=True)

    # --- MUERBETEIG ANSETZEN (Montag) ---
    if murt["grundrezepte"]>0 and "Montag" in vtage:
        murt_gr = murt["grundrezepte"]
        murt_srz = sub_rezepte.get("Mürbeteig") or sub_rezepte.get("Murbeteig")
        murt_aktiv = (murt_srz.get("aktiv_min",20) if murt_srz else 20) * murt_gr
        info("Montag","MUERBETEIG ANSETZEN",
             "{}x Grundrezept ({:.0f}g) - ueber Nacht kuehl stellen".format(
                 murt_gr,murt["gesamt_gramm"]),prio=PRIO["MURT_ANSET"],
             aktiv_min=murt_aktiv)

    # --- PISTAZIEN Vorbereitung (Mittwoch) ---
    pt=bedarf.get("Pistazien Toertchen",{})
    if pt.get("zu_produzieren",0)>0 and "Mittwoch" in vtage:
        info("Mittwoch","PISTAZIEN TOERTCHEN Vorbereitung",
             "Brownieboden backen + auskuehlen, Pistazienganache ansetzen",prio=PRIO["PIST_PREP"],
             aktiv_min=45, passiv_min=60)

    # --- KAESEKUCHEN: 4 Formen = 2 Charges/Tag ---
    kk_menge  =bedarf.get("Kaesekuchen",{}).get("zu_produzieren",0)
    kk_max    =PRODUKT_CONFIG["Kaesekuchen"]["max_ofen"]
    kk_formen =PRODUKT_CONFIG["Kaesekuchen"].get("formen",kk_max)
    kk_per_day=(kk_formen//kk_max)*kk_max if kk_max>0 else kk_formen
    kk_ges    =math.ceil(kk_menge/kk_max) if kk_menge else 0
    rem_kk,cn =kk_menge,1
    kk_montag =min(kk_per_day,rem_kk) if "Montag" in vtage else 0
    kk_a, kk_p = berechne_zeitaufwand("Kaesekuchen", rezepte, sub_rezepte, kk_max)
    for _ in range(math.ceil(kk_montag/kk_max) if kk_montag else 0):
        ch=min(kk_max,rem_kk)
        add("Montag","Kaesekuchen",ch,
            "Charge {}/{} - Murbeteigboden von letzter Woche".format(cn,kk_ges),
            prio=PRIO["KK"],produkt_key="Kaesekuchen",aktiv_min=kk_a,passiv_min=kk_p)
        rem_kk-=ch; cn+=1
        if rem_kk<=0: break
    if rem_kk > 0:
        _kk_rest = dict(bedarf["Kaesekuchen"])
        _kk_rest["zu_produzieren"] = rem_kk
        _kk_rest["im_lager"] = bedarf["Kaesekuchen"]["im_lager"] + kk_montag
        tage_kk = verteile_bestandssicher("Kaesekuchen", {"Kaesekuchen": _kk_rest}, aktiv_plan,
                                          _filter(["Dienstag","Mittwoch","Donnerstag","Freitag"]), kk_max)
    else:
        tage_kk = []
    for tag, _notw_kk in tage_kk:
        if rem_kk<=0: break
        heute=0
        while rem_kk>0 and heute<kk_per_day:
            ch=min(kk_max,rem_kk)
            add(tag,"Kaesekuchen",ch,"Charge {}/{}".format(cn,kk_ges),
                prio=PRIO["KK"],produkt_key="Kaesekuchen",aktiv_min=kk_a,passiv_min=kk_p,
                notwendig=_notw_kk)
            rem_kk-=ch; cn+=1; heute+=ch
    murt["kk_montag_reserve"]=min(kk_per_day,kk_menge)

    # --- KAESEKUCHEN Freitags-Mindestmenge erzwingen ---
    # Laut Planungsregeln: Freitag immer mind. 2 Chargen (4 Stk) fuer Wochenfrische.
    # Wird auch dann eingeplant wenn zu_produzieren=0 (Lager scheinbar ausreichend),
    # weil Frische-Aspekt und C2-Puffer-Aufbau es erfordern.
    if "Freitag" in vtage:
        fr_bereits = sum(a.get("menge",0) for a in plan["Freitag"]
                         if a.get("produkt_key")=="Kaesekuchen")
        kk_fr_min = kk_per_day   # 4 Stk = 2 Chargen
        fr_fehlt = max(0, kk_fr_min - fr_bereits)
        if fr_fehlt > 0:
            kk_ges = max(kk_ges, math.ceil((kk_menge + fr_fehlt) / kk_max))
            fr_rest = fr_fehlt
            while fr_rest > 0:
                ch = min(kk_max, fr_rest)
                add("Freitag","Kaesekuchen",ch,
                    "Charge {}/{} — Freitags-Mindestmenge (Frische Wochenende)".format(cn, kk_ges),
                    prio=PRIO["KK"],produkt_key="Kaesekuchen",
                    aktiv_min=kk_a,passiv_min=kk_p,notwendig=True,grund="frische")
                fr_rest -= ch; cn += 1

    # --- MUERBETEIG BACKEN (Dienstag) ---
    if murt["grundrezepte"]>0 and "Dienstag" in vtage:
        detail=", ".join("{}x {}".format(v["anzahl"],t) for t,v in murt["portionen"].items())
        res=murt.get("kk_montag_reserve",0)
        restext=" + {}x Murbeteig rund fuer naechsten Montag".format(res) if res else ""
        murt_back_aktiv = (sub_rezepte.get("Mürbeteig",{}).get("aktiv_min",20) or 20) * murt["grundrezepte"]
        info("Dienstag","MUERBETEIG BACKEN","Alle Boeden: {}{}".format(detail,restext),
             prio=PRIO["MURT_BACK"], aktiv_min=murt_back_aktiv, passiv_min=10)

    # --- STREUSEL herstellen (Montag, fuer ganze Woche) ---
    kr_streusel = bedarf.get("Kaese-Rhabarber Schnitte",{}).get("zu_produzieren",0)
    if kr_streusel > 0:
        streusel_pro_stk  = 175   # g pro KR Schnitte
        streusel_yield_gr = 1925  # g pro Grundrezept Streusel
        streusel_gesamt   = kr_streusel * streusel_pro_stk
        streusel_gr       = math.ceil(streusel_gesamt / streusel_yield_gr)
        streusel_srz = sub_rezepte.get("Streusel",{})
        streusel_max   = streusel_srz.get("max_charge_num", 99)
        streusel_aktiv_1 = streusel_srz.get("aktiv_min", 10) or 10  # pro Grundrezept
        if "Montag" in vtage:
            if streusel_gr <= streusel_max:
                info("Montag","STREUSEL HERSTELLEN",
                     "{}x Grundrezept ({:.0f}g) - fuer die ganze Woche".format(
                         streusel_gr, streusel_gesamt),
                     prio=PRIO["STREUSEL"], aktiv_min=streusel_aktiv_1 * streusel_gr)
            else:
                # In Einzelchargen aufteilen (max_charge_num Grundrezepte gleichzeitig)
                for _si in range(streusel_gr):
                    info("Montag","STREUSEL Charge {}/{}".format(_si+1, streusel_gr),
                         "1x Grundrezept ({:.0f}g) - Teil fuer die ganze Woche".format(
                             streusel_yield_gr),
                         prio=PRIO["STREUSEL"], aktiv_min=streusel_aktiv_1)

    # --- BIENENSTICH ---
    b_menge=bedarf.get("Bienenstich",{}).get("zu_produzieren",0)
    b_max=PRODUKT_CONFIG["Bienenstich"]["max_ofen"]
    b_ges=math.ceil(b_menge/b_max) if b_menge else 0
    rem,c=b_menge,1
    b_a, b_p = berechne_zeitaufwand("Bienenstich", rezepte, sub_rezepte, b_max)
    tage_b = verteile_bestandssicher("Bienenstich", bedarf, aktiv_plan,
                                     _filter(["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]), b_max)
    for tag, _notw in tage_b:
        if rem<=0: break
        ch=min(rem,b_max)
        add(tag,"Bienenstich",ch,"Charge {}/{}".format(c,b_ges),prio=PRIO["BIEN"],
            produkt_key="Bienenstich",aktiv_min=b_a,passiv_min=b_p,notwendig=_notw)
        rem-=ch; c+=1

    # --- QUICHE ---
    q_menge=bedarf.get("Lauch-Speck Quiche",{}).get("zu_produzieren",0)
    q_max=PRODUKT_CONFIG["Lauch-Speck Quiche"]["max_ofen"]
    q_ges=math.ceil(q_menge/q_max) if q_menge else 0
    rem,c=q_menge,1
    q_a, q_p = berechne_zeitaufwand("Lauch-Speck Quiche", rezepte, sub_rezepte, q_max)
    # Pflicht-Mindestchargen (Planungsregeln: Mo=1)
    for _tag, _n in MIN_CHARGEN.get("Lauch-Speck Quiche", {}).items():
        if _tag not in vtage: continue
        for _ in range(_n):
            if rem <= 0: break
            ch = min(rem, q_max)
            add(_tag,"Lauch-Speck Quiche",ch,
                "Charge {}/{} — Pflicht ({}=mind. {}/Tag)".format(c, q_ges, _tag, _n),
                prio=PRIO["QUICHE"],produkt_key="Lauch-Speck Quiche",
                aktiv_min=q_a,passiv_min=q_p,notwendig=True,grund="pflicht")
            rem -= ch; c += 1
    if rem > 0:
        _q_rest = dict(bedarf["Lauch-Speck Quiche"])
        _q_rest["zu_produzieren"] = rem
        _q_rest["im_lager"] = bedarf["Lauch-Speck Quiche"]["im_lager"] + (q_menge - rem)
        tage_q = verteile_bestandssicher("Lauch-Speck Quiche", {"Lauch-Speck Quiche": _q_rest}, aktiv_plan,
                                         _filter(["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]), q_max)
    else:
        tage_q = []
    for tag, _notw in tage_q:
        if rem<=0: break
        ch=min(rem,q_max)
        add(tag,"Lauch-Speck Quiche",ch,"Charge {}/{}".format(c,q_ges),prio=PRIO["QUICHE"],
            produkt_key="Lauch-Speck Quiche",aktiv_min=q_a,passiv_min=q_p,notwendig=_notw)
        rem-=ch; c+=1

    # --- DONAUWELLE ---
    d_menge=bedarf.get("Donauwelle",{}).get("zu_produzieren",0)
    d_max=PRODUKT_CONFIG["Donauwelle"]["max_ofen"]
    d_ges=math.ceil(d_menge/d_max) if d_menge else 0
    rem,c=d_menge,1
    d_a, d_p = berechne_zeitaufwand("Donauwelle", rezepte, sub_rezepte, d_max)
    tage_d = verteile_bestandssicher("Donauwelle", bedarf, aktiv_plan,
                                     _filter(["Dienstag","Mittwoch","Donnerstag","Freitag"]), d_max)
    for tag, _notw in tage_d:
        if rem<=0: break
        ch=min(rem,d_max)
        add(tag,"Donauwelle",ch,"Charge {}/{}".format(c,d_ges),prio=PRIO["DONA"],
            produkt_key="Donauwelle",aktiv_min=d_a,passiv_min=d_p,notwendig=_notw)
        rem-=ch; c+=1

    # --- KAESE-RHABARBER ---
    kr_menge=bedarf.get("Kaese-Rhabarber Schnitte",{}).get("zu_produzieren",0)
    kr_max=PRODUKT_CONFIG["Kaese-Rhabarber Schnitte"]["max_ofen"]
    kr_ges=math.ceil(kr_menge/kr_max) if kr_menge else 0
    rem,c=kr_menge,1
    kr_a, kr_p = berechne_zeitaufwand("Kaese-Rhabarber Schnitte", rezepte, sub_rezepte, kr_max)
    # Pflicht-Mindestchargen (Planungsregeln: Do=1)
    for _tag, _n in MIN_CHARGEN.get("Kaese-Rhabarber Schnitte", {}).items():
        if _tag not in vtage: continue
        for _ in range(_n):
            if rem <= 0: break
            ch = min(rem, kr_max)
            add(_tag,"Kaese-Rhabarber Schnitte",ch,
                "Charge {}/{} — Pflicht ({}=mind. {}/Tag)".format(c, kr_ges, _tag, _n),
                prio=PRIO["KR"],produkt_key="Kaese-Rhabarber Schnitte",
                aktiv_min=kr_a,passiv_min=kr_p,notwendig=True,grund="pflicht")
            rem -= ch; c += 1
    if rem > 0:
        _kr_rest = dict(bedarf["Kaese-Rhabarber Schnitte"])
        _kr_rest["zu_produzieren"] = rem
        _kr_rest["im_lager"] = bedarf["Kaese-Rhabarber Schnitte"]["im_lager"] + (kr_menge - rem)
        tage_kr = verteile_bestandssicher("Kaese-Rhabarber Schnitte", {"Kaese-Rhabarber Schnitte": _kr_rest},
                                         aktiv_plan,
                                         _filter(["Dienstag","Mittwoch","Donnerstag","Freitag"]), kr_max)
    else:
        tage_kr = []
    for tag, _notw in tage_kr:
        if rem<=0: break
        ch=min(rem,kr_max)
        add(tag,"Kaese-Rhabarber Schnitte",ch,"Charge {}/{}".format(c,kr_ges),prio=PRIO["KR"],
            produkt_key="Kaese-Rhabarber Schnitte",aktiv_min=kr_a,passiv_min=kr_p,notwendig=_notw)
        rem-=ch; c+=1

    # --- TARTELETTES ---
    for key,tag in [("Beeren Tartelette","Mittwoch"),("Mango-Passionsfrucht Tartelette","Donnerstag")]:
        m=bedarf.get(key,{}).get("zu_produzieren",0)
        if m>0 and tag in vtage:
            t_a, t_p = berechne_zeitaufwand(key, rezepte, sub_rezepte, m)
            add(tag,key,m,"{} Stueck".format(m),prio=PRIO["TART"],produkt_key=key,
                aktiv_min=t_a,passiv_min=t_p)

    # --- PISTAZIEN Hauptproduktion (Donnerstag) ---
    if pt.get("zu_produzieren",0)>0 and "Donnerstag" in vtage:
        pt_a, pt_p = berechne_zeitaufwand("Pistazien Toertchen", rezepte, sub_rezepte, 60)
        add("Donnerstag","PISTAZIEN TOERTCHEN fertigstellen",pt["zu_produzieren"],
            "Kern + Schoko-Mousse, schockfrosten, Ganache",prio=PRIO["PIST_MAIN"],
            aktiv_min=pt_a, passiv_min=pt_p)

    # Sortiere jeden Tag nach Prioritaet
    for tag in ARBEITSTAGE:
        plan[tag].sort(key=lambda x: x["prio"])

    # Einfrier-Empfehlung pro Charge berechnen (siehe Planungsregeln.md "Einfrierempfehlung")
    berechne_einfrierempfehlung(plan, bedarf, rezepte, vtage)

    return vtage, plan

def berechne_einfrierempfehlung(plan, bedarf, rezepte, vtage):
    """Berechnet pro Produktionscharge "frisch lassen" und "einfrieren".

    Regel (Planungsregeln.md):
      frisch_lassen = tagesverbrauch * tage_bis_naechste_charge_verfuegbar
      einfrieren    = produktion - frisch_lassen
      obergrenze    = haltbarkeit_kuehl * tagesverbrauch  (cap)

    Verfuegbar erst ab Folgetag — eine Charge produziert am Tag X wird also
    erst ab X+1 Cafe-1-Bestand. Das letzte Batch der Woche muss ueber das
    Wochenende plus bis zum naechsten Backtag in der Folgewoche reichen.
    """
    # Pro Produkt alle Chargen mit ihrem Backtag sammeln
    from collections import defaultdict as _dd
    chargen_pro_produkt = _dd(list)
    for tag in ARBEITSTAGE:
        for a in plan.get(tag, []):
            pk = a.get("produkt_key")
            if not pk or a.get("excel_only"): continue
            if pk not in PRODUKT_CONFIG: continue
            if a.get("menge",0) <= 0: continue
            chargen_pro_produkt[pk].append((tag, a))

    for pk, chargen in chargen_pro_produkt.items():
        bd = bedarf.get(pk, {})
        tv_profil = bd.get("tv_profil") or {}
        tv_avg    = bd.get("gesamt_woche", 0.0) / 7.0
        rez       = (rezepte or {}).get(pk) or {}
        h_kuehl   = rez.get("haltbarkeit_kuehl_tage")

        # Tage absteigend nach Wochentag-Index sortieren, damit "naechste Charge"
        # via Listen-Lookup einfach ist. Wir gehen aber chronologisch durch.
        # Sortiere nach Reihenfolge in ARBEITSTAGE.
        chargen.sort(key=lambda x: ARBEITSTAGE.index(x[0]))

        # Backtage des Produkts (geordnete eindeutige Tage)
        backtage = []
        for tag,_ in chargen:
            if tag not in backtage: backtage.append(tag)

        for idx, (tag, a) in enumerate(chargen):
            menge = int(a.get("menge",0))
            if menge <= 0: continue
            # Verfuegbar ab Folgetag (idx im ARBEITSTAGE-Raster)
            i_tag = ARBEITSTAGE.index(tag)
            # Naechste Charge desselben Produkts (chronologisch)
            naechste_tage = [t for t,_ in chargen if ARBEITSTAGE.index(t) > i_tag]
            if naechste_tage:
                i_next = ARBEITSTAGE.index(naechste_tage[0])
                # Fenster = Tage von (i_tag+1) bis (i_next) inklusive Folgetag der naechsten Produktion?
                # Regel: bis das naechste Batch verfuegbar ist (=i_next + 1). Also Tage i_tag+1 .. i_next.
                fenster_tage = list(range(i_tag+1, i_next+1))
            else:
                # Letzte Charge der Woche: Fr/Sa/So + Mo der Folgewoche (bis naechster Backtag)
                # Wir nehmen pauschal: bis einschliesslich Dienstag naechste Woche (~ 4 zusaetzliche Tage)
                # = Folgetag + Sa + So + Mo + Di -> ueblicherweise 5 Tage.
                fenster_tage = list(range(i_tag+1, 5)) + [5,6,0,1]  # Rest-Woche + Sa+So+Mo+Di
                # Da wir Mengen×Tage ueber Wochentag-Indizes summieren, bauen wir Liste explizit.
                # Wir verwenden Wochentag-Namen via Modulo.
            # Wochentage-Namen aus den Indizes ableiten (Sa/So → kein Verkauf, tv=0)
            kalender = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag","Sonntag"]
            fenster_namen = [kalender[i%7] for i in fenster_tage]
            tv_fenster = sum(tv_profil.get(t, tv_avg) if t in ARBEITSTAGE else 0
                             for t in fenster_namen)
            frisch = int(round(tv_fenster))
            # Obergrenze: Haltbarkeit (Kuehlung) * durchschnittlicher Tagesverbrauch
            if h_kuehl:
                # cap = haltbarkeit * Mittel der Verbrauchs-Tage im Fenster (oder Wochenmittel)
                tv_mit = (tv_fenster / max(1,len([t for t in fenster_namen if t in ARBEITSTAGE]))) \
                          if any(t in ARBEITSTAGE for t in fenster_namen) else tv_avg
                cap = int(round(h_kuehl * (tv_mit or tv_avg)))
                if frisch > cap: frisch = cap
            frisch = max(0, min(menge, frisch))
            einfrier = max(0, menge - frisch)
            a["frisch_stk"]   = frisch
            a["einfrier_stk"] = einfrier
            a["haltbarkeit_kuehl_tage"] = h_kuehl
            # Notiz erweitern
            if einfrier > 0 or frisch < menge:
                zusatz = " → {} frisch, {} einfrieren".format(frisch, einfrier)
                if a.get("notiz"): a["notiz"] = a["notiz"] + zusatz
                else: a["notiz"] = "{} Stk".format(menge) + zusatz

# ============================================================
# REZEPT-HTML HELFER
# ============================================================
def format_rezept_zutaten_html(sections, skala):
    """Gibt HTML fuer skalierte Zutaten-Abschnitte zurueck."""
    out=""
    for sec in sections:
        if len(sections)>1:
            out+='<div class="rzp-sec">{}</div>'.format(sec['name'])
        for z in sec['zutaten']:
            if z['einheit']=='ref':
                ref_txt=str(z['menge'])
                out+='<li><span class="zutat">{}</span><span class="menge">{}  (vorh. backen/vorbereiten)</span></li>'.format(z['zutat'],ref_txt)
            else:
                menge_skaliert=round(z['menge']*skala,1)
                menge_str="{}g".format(int(menge_skaliert) if menge_skaliert==int(menge_skaliert) else menge_skaliert)
                out+='<li><span class="zutat">{}</span><span class="menge">{}</span></li>'.format(z['zutat'],menge_str)
    return out

def format_rezept_block_html(rezept, daily_qty):
    """Gibt einen vollstaendigen Rezept-Block als HTML zurueck."""
    if not rezept or not rezept.get('sections'): return ""
    max_c=rezept['max_charge_num']
    # Batch: min(max_charge, daily_qty) - wie viele auf einmal
    batch=min(max_c, daily_qty) if max_c<=daily_qty else daily_qty
    n_mal=math.ceil(daily_qty/max_c) if max_c>0 else 1
    # Skalierungsfaktor: Rezept ist pro 1 Stueck, zeige fuer 1 Batch
    skala=batch
    titel="Rezept fuer {} Stk".format(batch) if batch>1 else "Grundrezept (1 Stk)"
    if n_mal>1: titel+=" &mdash; heute {}x ansetzen".format(n_mal)
    else: titel+=" &mdash; 1x ansetzen"
    zutaten_html=format_rezept_zutaten_html(rezept['sections'],skala)
    return (
        '<div class="recipe">'
        '<div class="recipe-title">{}</div>'
        '<ul class="zutaten">{}</ul>'
        '</div>'
    ).format(titel,zutaten_html)

# ============================================================
# PREISE & EINKAUFSLISTE
# ============================================================
def lade_preise():
    """Laedt Preise.xlsx (Kalkulation/Preise.xlsx).
    Gibt dict {zutat_lower: {name, preis_kg, selbst_bestellen}} zurueck.
    """
    pfad = BASE / "Kalkulation" / "Preise.xlsx"
    preise = {}
    if not pfad.exists():
        print("  Warnung: Preise.xlsx nicht gefunden ({})".format(pfad))
        return preise
    try:
        wb = openpyxl.load_workbook(pfad, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 3:
                continue
            zutat = str(row[1] or '').strip()
            if not zutat or zutat.lower() == 'zutat':
                continue
            try:
                preis_kg = float(str(row[2] or '0').replace(',', '.'))
            except Exception:
                preis_kg = 0.0
            # Spalte D = Selbst bestellen (default: Ja)
            selbst_raw = str(row[3] or 'Ja').strip().lower() if len(row) > 3 else 'ja'
            selbst = selbst_raw in ('ja', 'yes', 'j', 'true', '1')
            preise[zutat.lower()] = {
                'name':             zutat,
                'preis_kg':         preis_kg,
                'selbst_bestellen': selbst,
            }
    except Exception as e:
        print("  Preise laden fehlgeschlagen: {}".format(e))
    return preise


def _preis_lookup(zutat, preise):
    """Sucht Preis-Eintrag fuer eine Zutat (case-insensitive, partial match)."""
    key = zutat.lower().strip()
    if key in preise:
        return preise[key]
    for pk, pv in preise.items():
        if key in pk or pk in key:
            return pv
    return None


def berechne_zutaten(aufgaben, rezepte, sub_rezepte):
    """Berechnet alle benoetigten Zutaten (in Gramm) fuer eine Liste von Aufgaben.

    Unterstuetzt:
    - Hauptprodukt-Aufgaben (produkt_key gesetzt, menge > 0): Rezept * Menge
    - MUERBETEIG ANSETZEN: MURBETEIG_REZEPT * Anzahl Grundrezepte (aus notiz)
    - STREUSEL-Aufgaben: Sub-Rezept Streusel * n_mal (aus notiz oder 1x pro Charge)
    - Andere Sub-Rezept-Aufgaben: Namens-Matching gegen sub_rezepte

    Gibt {zutat_name: gramm} zurueck.
    """
    zutaten = defaultdict(float)

    for a in aufgaben:
        if a.get('excel_only'):
            continue
        produkt = a.get('produkt', '')
        menge   = a.get('menge', 0)
        notiz   = a.get('notiz', '')
        pk      = a.get('produkt_key')
        prod_up = produkt.upper()

        # ── Hauptprodukt-Rezept ──────────────────────────────────
        if pk and pk in rezepte and menge > 0:
            # Sub-Rezept-Namen (z.B. "Streusel", "Vanillecreme") überspringen:
            # ihre Rohzutaten erscheinen an dem Tag, an dem sie hergestellt werden.
            sub_namen = {k.lower() for k in sub_rezepte.keys()}
            # Rezeptzutaten sind pro Stück → skalieren mit Stückzahl
            for sec in rezepte[pk].get('sections', []):
                for z in sec.get('zutaten', []):
                    if z['einheit'] == 'g':
                        if z['zutat'].lower() in sub_namen:
                            continue  # In-House-Produkt: nicht in Einkaufsliste
                        zutaten[z['zutat']] += z['menge'] * menge

        # ── Muerbeteig Ansetzen ──────────────────────────────────
        elif 'MUERBETEIG ANSETZEN' in prod_up:
            m = re.search(r'(\d+)x', notiz)
            n_mal = int(m.group(1)) if m else 1
            for z, mg in MURBETEIG_REZEPT.items():
                zutaten[z] += mg * n_mal

        # ── Streusel ─────────────────────────────────────────────
        elif 'STREUSEL' in prod_up:
            srz = sub_rezepte.get('Streusel') or sub_rezepte.get('Streusel ')
            if srz:
                if 'CHARGE' in prod_up:
                    n_mal = 1
                else:
                    m = re.search(r'(\d+)x', notiz)
                    n_mal = int(m.group(1)) if m else 1
                for z in srz.get('zutaten', []):
                    if z['einheit'] == 'g':
                        zutaten[z['zutat']] += z['menge'] * n_mal

        # ── Andere Sub-Rezepte per Namens-Match ──────────────────
        elif not pk:
            for srz_key, srz in sub_rezepte.items():
                if srz_key.upper() in prod_up:
                    m = re.search(r'(\d+)x', notiz)
                    n_mal = int(m.group(1)) if m else 1
                    for z in srz.get('zutaten', []):
                        if z['einheit'] == 'g':
                            zutaten[z['zutat']] += z['menge'] * n_mal
                    break

    return dict(zutaten)


def erstelle_html_einkaufsliste(tag, zutaten, preise):
    """Gibt HTML fuer die Tages-Einkaufsliste zurueck (einzubetten in Tagesplan)."""
    if not zutaten:
        return ''

    # Zutaten per kanonischem Preis-Namen zusammenfassen
    # (z.B. 'Milch' + 'Milch (Staerke)' -> 'Milch')
    merged = {}  # canon_name -> {gramm, preis_kg}
    for zutat_name, gramm in zutaten.items():
        if gramm <= 0:
            continue
        info = _preis_lookup(zutat_name, preise)
        canon = info['name'] if info else zutat_name
        preis_kg = info['preis_kg'] if info else 0.0
        if canon not in merged:
            merged[canon] = {'gramm': 0.0, 'preis_kg': preis_kg}
        merged[canon]['gramm'] += gramm

    rows = []
    gesamt_kosten = 0.0
    for canon in sorted(merged.keys()):
        gramm    = merged[canon]['gramm']
        preis_kg = merged[canon]['preis_kg']
        kg       = gramm / 1000.0
        kosten   = kg * preis_kg
        gesamt_kosten += kosten
        kg_str     = '{:.2f} kg'.format(round(kg, 2)) if kg >= 0.1 else '{:.0f} g'.format(gramm)
        kosten_str = '{:.2f} &euro;'.format(kosten) if kosten > 0 else '&ndash;'
        rows.append(
            '<tr><td class="z-name">{}</td>'
            '<td class="z-menge">{}</td>'
            '<td class="z-kosten">{}</td></tr>'.format(
                canon, kg_str, kosten_str
            )
        )

    if not rows:
        return ''

    gesamt_str = '{:.2f} &euro;'.format(gesamt_kosten) if gesamt_kosten > 0 else ''
    gesamt_row = (
        '<tr class="ek-total">'
        '<td colspan="2">Gesamtkosten</td>'
        '<td class="z-kosten">{}</td></tr>'
    ).format(gesamt_str) if gesamt_str else ''

    return (
        '<div class="ek-panel">'
        '<div class="ek-title">&#128722; Einkaufsliste {}</div>'
        '<table class="ek-table">'
        '<thead><tr>'
        '<th>Zutat</th><th>Menge</th><th>Kosten</th>'
        '</tr></thead>'
        '<tbody>{}</tbody>'
        '<tfoot>{}</tfoot>'
        '</table></div>'
    ).format(tag, ''.join(rows), gesamt_row)


def erstelle_html_bestellliste(bedarf, wplan, vtage, verkauf, lager,
                                rezepte, sub_rezepte, preise, puffer_pct=20):
    """Erstellt Wochenplan/Bestellliste.html.

    Berechnung:
    - Diese Woche: Summe aller Zutaten aus wplan (vtage)
    - Naechste Woche Mo-Do: (diese Woche / Anzahl_Tage) * 4 * (1 + puffer/100)
    - Gesamt = Summe beider Perioden
    - Nur Zutaten mit selbst_bestellen=True werden angezeigt.
    """
    # ── Zutaten dieser Woche ─────────────────────────────────────
    alle_aufgaben = []
    for tag in vtage:
        alle_aufgaben.extend(wplan.get(tag, []))
    zutaten_woche = berechne_zutaten(alle_aufgaben, rezepte, sub_rezepte)

    # ── Naechste Woche schätzen (Mo-Do = 4 Tage) ─────────────────
    n_tage = max(len(vtage), 1)
    faktor = (4.0 / n_tage) * (1.0 + puffer_pct / 100.0)
    zutaten_naechste = {z: g * faktor for z, g in zutaten_woche.items()}

    # ── Gesamt ───────────────────────────────────────────────────
    gesamt_d = defaultdict(float)
    for z, g in zutaten_woche.items():
        gesamt_d[z] += g
    for z, g in zutaten_naechste.items():
        gesamt_d[z] += g

    # ── Nur "Selbst bestellen"-Positionen, nach kanonischem Preis-Namen zusammenfassen ─
    merged = {}   # canonical_name -> {g_woche, g_naechste, info}
    for zutat_name, gramm_ges in gesamt_d.items():
        if gramm_ges <= 0:
            continue
        info = _preis_lookup(zutat_name, preise)
        if not info or not info['selbst_bestellen']:
            continue
        canon = info['name']
        if canon not in merged:
            merged[canon] = {'g_woche': 0.0, 'g_naechste': 0.0, 'info': info}
        merged[canon]['g_woche']    += zutaten_woche.get(zutat_name, 0)
        merged[canon]['g_naechste'] += zutaten_naechste.get(zutat_name, 0)

    bestell_items = []
    for canon in sorted(merged.keys()):
        m = merged[canon]
        g_ges = m['g_woche'] + m['g_naechste']
        bestell_items.append({
            'name':       canon,
            'g_woche':    m['g_woche'],
            'g_naechste': m['g_naechste'],
            'g_gesamt':   g_ges,
            'preis_kg':   m['info']['preis_kg'],
            'kosten':     (g_ges / 1000.0) * m['info']['preis_kg'],
        })

    # ── HTML ─────────────────────────────────────────────────────
    rows_html = ''
    gesamt_kosten = 0.0
    for item in bestell_items:
        def _kg(g):
            return '{:.2f} kg'.format(round(g / 1000.0, 2))
        rows_html += (
            '<tr>'
            '<td class="bl-name">{name}</td>'
            '<td class="bl-num">{dw}</td>'
            '<td class="bl-num">{nw}</td>'
            '<td class="bl-num bl-bold">{g}</td>'
            '<td class="bl-num">{p:.2f} &euro;/kg</td>'
            '<td class="bl-num">{k:.2f} &euro;</td>'
            '</tr>'
        ).format(
            name=item['name'],
            dw=_kg(item['g_woche']),
            nw=_kg(item['g_naechste']),
            g=_kg(item['g_gesamt']),
            p=item['preis_kg'],
            k=item['kosten'],
        )
        gesamt_kosten += item['kosten']

    if not rows_html:
        rows_html = '<tr><td colspan="6" style="text-align:center;color:#888">Keine Bestellpositionen gefunden</td></tr>'

    kw      = datetime.now().isocalendar()[1]
    kw_next = kw + 1
    datum   = datetime.now().strftime('%d.%m.%Y')

    css = (
        'body{font-family:sans-serif;padding:16px;max-width:960px;margin:0 auto;background:#f8f9fa}'
        'h1{color:#2E75B6;font-size:1.4em;margin-bottom:4px}'
        '.info{background:#e8f0fe;border-radius:8px;padding:10px 14px;margin-bottom:16px;'
        '      font-size:.88em;color:#1a3a6b;line-height:1.5}'
        'table{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;'
        '      overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,.1);font-size:.9em}'
        'th{background:#2E75B6;color:#fff;padding:9px 8px;text-align:left}'
        'th.bl-num{text-align:right}'
        'td{padding:7px 8px;border-bottom:1px solid #eee}'
        'td.bl-num{text-align:right}'
        'td.bl-bold{font-weight:700}'
        'tr:last-child td{border-bottom:none}'
        'tr:hover td{background:#f0f7ff}'
        '.total-row td{font-weight:700;background:#D6E4F0;border-top:2px solid #2E75B6}'
        '.footer{text-align:center;color:#aaa;font-size:.78em;margin-top:20px}'
    )

    html = (
        '<!DOCTYPE html>\n'
        '<html lang="de"><head>'
        '<meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1.0">'
        '<title>Bestellliste KW {kw}</title>'
        '<style>{css}</style></head><body>'
        '<h1>&#128230; Bestellliste KW {kw}</h1>'
        '<div class="info">'
        '&#128197; Bestellung: Montag {datum} &nbsp;&bull;&nbsp; '
        'Lieferung: Di&ndash;Do &nbsp;&bull;&nbsp; '
        'Reicht bis: Do naechste Woche (KW {kw_next})<br>'
        '&#128204; Nur Artikel mit <strong>Selbst bestellen</strong> '
        '(Logistiker-Artikel sind ausgeblendet)<br>'
        '&#128200; Naechste Woche: Hochrechnung aus dieser Woche ({puffer}% Puffer, Mo&ndash;Do = 4 Tage)'
        '</div>'
        '<table>'
        '<thead><tr>'
        '<th>Zutat</th>'
        '<th class="bl-num">Diese Woche</th>'
        '<th class="bl-num">Naechste Woche (Mo&ndash;Do)</th>'
        '<th class="bl-num">Bestellen</th>'
        '<th class="bl-num">Preis / kg</th>'
        '<th class="bl-num">Kosten</th>'
        '</tr></thead>'
        '<tbody>{rows}</tbody>'
        '<tfoot><tr class="total-row">'
        '<td colspan="5">Gesamtkosten (Selbst bestellen)</td>'
        '<td class="bl-num">{gesamt:.2f}&nbsp;&euro;</td>'
        '</tr></tfoot>'
        '</table>'
        '<div class="footer">Kuchenproduktion &middot; KW {kw}</div>'
        '</body></html>'
    ).format(
        kw=kw, kw_next=kw_next, datum=datum, puffer=puffer_pct,
        css=css, rows=rows_html, gesamt=gesamt_kosten
    )

    pfad = OUTPUT_PFAD / 'Bestellliste.html'
    pfad.write_text(html, encoding='utf-8')
    print('Bestellliste: {}'.format(pfad.name))



# ============================================================
# OUTPUT EXCEL
# ============================================================
def erstelle_excel(bedarf, murt, wplan):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb=openpyxl.Workbook()
    blau=PatternFill("solid",fgColor="2E75B6"); hblau=PatternFill("solid",fgColor="D6E4F0")
    gruen=PatternFill("solid",fgColor="375623"); hgruen=PatternFill("solid",fgColor="D9EAD3")
    gelb=PatternFill("solid",fgColor="FFF2CC"); oran=PatternFill("solid",fgColor="FCE4D6")
    wf=Font(bold=True,color="FFFFFF"); bf=Font(bold=True)
    thin=Side(style="thin"); brd=Border(left=thin,right=thin,top=thin,bottom=thin)
    def sc(cell,fill=None,font=None,align="center",wrap=False):
        if fill: cell.fill=fill
        if font: cell.font=font
        cell.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
        cell.border=brd

    ws=wb.active; ws.title="Produktionsbedarf"
    ws.merge_cells("A1:H1")
    ws["A1"]="Produktionsbedarf Woche ab {}".format(datetime.now().strftime("%d.%m.%Y"))
    sc(ws["A1"],blau,Font(bold=True,color="FFFFFF",size=13)); ws.row_dimensions[1].height=28
    for i,h in enumerate(["Produkt","Cafe 1","Cafe 2","Gesamt","Im Lager","Nettobedarf","Produzieren","Cafe2 Di"]):
        sc(ws.cell(row=2,column=i+1,value=h),blau,wf)
    ws.row_dimensions[2].height=30
    for r,(key,bd) in enumerate(bedarf.items()):
        row=r+3
        sc(ws.cell(row=row,column=1,value=key),hblau,bf,"left")
        for ci,v in enumerate([bd["cafe1_woche"],bd["cafe2_woche"],bd["gesamt_woche"],
                                bd["im_lager"],bd["netto_bedarf"],bd["zu_produzieren"],
                                bd["cafe2_lieferung"] if PRODUKT_CONFIG[key]["cafe2"] else "-"]):
            fill=oran if ci==5 else (gelb if ci==4 else None)
            sc(ws.cell(row=row,column=ci+2,value=v),fill)
        if bd.get("grund"):
            ws.cell(row=row,column=9,value=bd["grund"]).alignment=Alignment(horizontal="left",wrap_text=True)
        ws.row_dimensions[row].height=22
    mr=len(bedarf)+4
    ws.merge_cells("A{}:H{}".format(mr,mr))
    sc(ws.cell(row=mr,column=1,value="MUERBETEIG"),gruen,Font(bold=True,color="FFFFFF")); mr+=1
    ws.cell(row=mr,column=1,value="Grundrezepte:"); ws.cell(row=mr,column=2,value=murt["grundrezepte"])
    ws.cell(row=mr,column=3,value="{:.0f}g".format(murt["gesamt_gramm"]))
    sc(ws.cell(row=mr,column=1),hgruen,bf,"left"); sc(ws.cell(row=mr,column=2),gelb,bf); mr+=1
    for t,ti in murt["portionen"].items():
        ws.cell(row=mr,column=1,value="  {}:".format(t))
        ws.cell(row=mr,column=2,value="{}x/{}g".format(ti["anzahl"],int(ti["gramm"])))
        sc(ws.cell(row=mr,column=1),hgruen,align="left"); sc(ws.cell(row=mr,column=2),hgruen); mr+=1
    ws.column_dimensions["A"].width=32
    for col in "BCDEFGH": ws.column_dimensions[col].width=13

    ws2=wb.create_sheet("Wochenplan")
    ws2.merge_cells("A1:D1")
    sc(ws2.cell(row=1,column=1,value="WOCHENPLAN"),blau,Font(bold=True,color="FFFFFF",size=13))
    ws2.row_dimensions[1].height=28
    for i,h in enumerate(["Tag","Aufgabe","Menge","Notiz"]):
        sc(ws2.cell(row=2,column=i+1,value=h),blau,wf)
    tfill={"Montag":PatternFill("solid",fgColor="E2EFDA"),"Dienstag":PatternFill("solid",fgColor="DDEBF7"),
           "Mittwoch":PatternFill("solid",fgColor="FFF2CC"),"Donnerstag":PatternFill("solid",fgColor="FCE4D6"),
           "Freitag":PatternFill("solid",fgColor="F2F2F2")}
    row=3
    for tag in ARBEITSTAGE:
        aufgaben=wplan.get(tag,[])
        if not aufgaben: continue
        ws2.merge_cells("A{}:D{}".format(row,row))
        sc(ws2.cell(row=row,column=1,value=tag.upper()),tfill.get(tag),bf,"left")
        ws2.row_dimensions[row].height=22; row+=1
        for a in aufgaben:
            ws2.cell(row=row,column=1,value="")
            ws2.cell(row=row,column=2,value=a["produkt"])
            ws2.cell(row=row,column=3,value=a["menge"] if a["menge"] else "")
            ws2.cell(row=row,column=4,value=a["notiz"])
            for col in range(1,5):
                ws2.cell(row=row,column=col).alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
                ws2.cell(row=row,column=col).border=brd
            ws2.row_dimensions[row].height=22; row+=1
    ws2.column_dimensions["A"].width=12; ws2.column_dimensions["B"].width=38
    ws2.column_dimensions["C"].width=8;  ws2.column_dimensions["D"].width=55
    fname=OUTPUT_PFAD/"Wochenuebersicht_{}.xlsx".format(datetime.now().strftime("%Y%m%d"))
    wb.save(fname); print("Wochenuebersicht: {}".format(fname.name)); return fname


def aktuelle_kw_label():
    """Aktuelle Kalenderwoche als 'YYYY-WW' Label (passend zu ist_produktion_KW*.json)."""
    iso = datetime.now().isocalendar()
    return "{}-{:02d}".format(iso[0], iso[1])


def lese_ist_produktion(kw_label=None):
    """Liest Wochenplan/ist_produktion_KW{}.json (vom Tagesplan-UI gepflegt).

    Schema: {kw, _format:1, tage:{Montag:{Kaesekuchen:{ist_menge, erledigt, verschoben_nach}}}}
    Returns: {tag: {produkt_key: rec}} oder {} wenn Datei fehlt/defekt.
    """
    if kw_label is None:
        kw_label = aktuelle_kw_label()
    pfad = OUTPUT_PFAD / "ist_produktion_KW{}.json".format(kw_label)
    if not pfad.exists():
        return {}
    try:
        with open(pfad, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("tage", {}) or {}
    except (json.JSONDecodeError, IOError, OSError):
        return {}


def lese_wplan_kw(kw_label=None):
    """Liest die zuletzt vom Planer geschriebene wplan-Snapshot der Woche.
    Wird genutzt, um vergangene Tage in der Wochenmatrix zu rekonstruieren,
    wenn der current wplan fuer diese Tage leer ist.

    Returns: {tag: [{produkt_key, menge, aktiv_min, passiv_min}, ...]} oder {}.
    """
    if kw_label is None:
        kw_label = aktuelle_kw_label()
    pfad = OUTPUT_PFAD / "wplan_KW{}.json".format(kw_label)
    if not pfad.exists():
        return {}
    try:
        with open(pfad, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("tage", {}) or {}
    except (json.JSONDecodeError, IOError, OSError):
        return {}


def speichere_wplan_kw(wplan, vtage=None):
    """Schreibt Wochenplan/wplan_KW{}.json. Nur Tage in vtage werden ueberschrieben;
    Eintraege fuer vergangene Tage aus frueheren Laeufen dieser Woche bleiben
    erhalten. So sammelt der Snapshot ueber die Woche hinweg den jeweils
    aktuellsten Plan jedes Tages.
    """
    kw_label = aktuelle_kw_label()
    pfad = OUTPUT_PFAD / "wplan_KW{}.json".format(kw_label)
    bestehend = lese_wplan_kw(kw_label)
    tage = dict(bestehend)
    aktive = set(vtage or [])
    for tag, aufgaben in (wplan or {}).items():
        if aktive and tag not in aktive:
            continue
        relevante = []
        for a in (aufgaben or []):
            if a.get("excel_only"):
                continue
            pk = a.get("produkt_key")
            menge = int(a.get("menge", 0) or 0)
            if not pk or menge <= 0:
                continue
            relevante.append({
                "produkt_key": pk,
                "menge": menge,
                "aktiv_min": int(a.get("aktiv_min", 0) or 0),
                "passiv_min": int(a.get("passiv_min", 0) or 0),
            })
        tage[tag] = relevante  # leer = nichts an dem Tag geplant
    payload = {
        "kw": kw_label,
        "_format": 1,
        "_updated": datetime.now().isoformat(),
        "tage": tage,
    }
    try:
        with open(pfad, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print("Wplan-Snapshot: {}".format(pfad.name))
    except (IOError, OSError) as e:
        print("WARN: wplan-Snapshot konnte nicht geschrieben werden: {}".format(e))


def erstelle_html_wochenuebersicht(bedarf, murt, wplan, rolling, tages_info, vtage=None):
    """Generiert Wochenplan/Wochenuebersicht.html fuer mobilen Zugriff."""
    FARBEN = {
        "Montag":    ("#E2EFDA","#375623"),
        "Dienstag":  ("#DDEBF7","#2E75B6"),
        "Mittwoch":  ("#FFF2CC","#7B6000"),
        "Donnerstag":("#FCE4D6","#833C00"),
        "Freitag":   ("#F2F2F2","#404040"),
    }

    def zeit_str(minuten):
        if not minuten: return ""
        h,m = divmod(int(minuten),60)
        return "{}h{}min".format(h,m) if h else "{}min".format(m)

    html = []
    html.append('<!DOCTYPE html><html lang="de"><head>')
    html.append('<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">')
    html.append('<title>Wochenuebersicht</title>')
    html.append('<style>')
    html.append('*{box-sizing:border-box;margin:0;padding:0}')
    html.append('body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f0f2f5;padding-bottom:32px}')
    html.append('.header{background:#2E75B6;color:#fff;padding:14px 18px;position:sticky;top:0;z-index:10}')
    html.append('.header h1{font-size:17px;font-weight:700}.header .sub{font-size:11px;opacity:.8;margin-top:2px}')
    html.append('.day-card{margin:12px 12px 0;border-radius:14px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.09)}')
    html.append('.day-header{padding:11px 16px;font-size:13px;font-weight:700;letter-spacing:.5px;text-transform:uppercase}')
    html.append('.task-row{background:#fff;display:flex;align-items:flex-start;padding:11px 14px;border-top:1px solid #f0f0f0;gap:10px}')
    html.append('.task-name{flex:1;font-size:14px;font-weight:500;line-height:1.3}')
    html.append('.task-notiz{font-size:12px;color:#888;margin-top:2px}')
    html.append('.task-zeit{font-size:11px;color:#2E75B6;margin-top:2px}')
    html.append('.lager-panel{background:#fff;margin:12px 12px 0;border-radius:14px;padding:12px 14px;box-shadow:0 1px 4px rgba(0,0,0,.09)}')
    html.append('.lager-title{font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.7px;margin-bottom:8px}')
    html.append('.lager-row{display:flex;justify-content:space-between;font-size:13px;padding:3px 0;border-bottom:1px solid #f5f5f5}')
    html.append('.lager-row:last-child{border-bottom:none}')
    html.append('.lager-leer{color:#c0392b;font-weight:700}.lager-niedrig{color:#e67e22;font-weight:600}')
    html.append('.grund-badge{display:inline-block;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;margin-left:6px;vertical-align:middle;letter-spacing:.3px}')
    html.append('.g-pflicht{background:#E1F0FF;color:#1B5FAB}.g-bestand{background:#FFE4E1;color:#C0392B}')
    html.append('.g-frische{background:#E2F0CB;color:#3F7D20}.g-puffer{background:#FFF4D4;color:#8A6D1F}')
    html.append('.matrix{margin:12px 12px 0;border-radius:14px;background:#fff;padding:12px 14px;box-shadow:0 1px 4px rgba(0,0,0,.09);overflow-x:auto}')
    html.append('.matrix-title{font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.7px;margin-bottom:8px}')
    html.append('.matrix table{width:100%;border-collapse:collapse;font-size:12px;min-width:340px}')
    html.append('.matrix th,.matrix td{padding:5px 6px;text-align:center;border-bottom:1px solid #f3f3f3}')
    html.append('.matrix th{font-weight:600;color:#555;background:#fafafa}.matrix td:first-child,.matrix th:first-child{text-align:left;font-weight:500;color:#333}')
    html.append('.matrix .z{color:#bbb}.matrix .sum{font-weight:700;color:#1B5FAB;background:#f7fafd}')
    # Erledigt = gruene Zahl auf hellgruenem Hintergrund. "mod" = vom User geaendert.
    html.append('.matrix td.done{background:#E8F5EA}.matrix td.done b{color:#2D8B3E}')
    html.append('.matrix td.mod{position:relative}.matrix td.mod::after{content:"\\270e";position:absolute;top:1px;right:3px;font-size:9px;color:#999}')
    html.append('.matrix td.past{background-image:linear-gradient(135deg,transparent 0,transparent 6px,#f5f5f5 6px,#f5f5f5 7px,transparent 7px,transparent 12px,#f5f5f5 12px,#f5f5f5 13px)}')
    html.append('.matrix-legend{margin:4px 12px 0;font-size:10px;color:#888;text-align:right}')
    html.append('.matrix-legend .lg-done{color:#2D8B3E;font-weight:700;background:#E8F5EA;padding:0 4px;border-radius:3px}')
    html.append('</style></head><body>')

    html.append('<div class="header"><h1>&#128197; Wochenuebersicht</h1>')
    html.append('<div class="sub">Stand: {}</div></div>'.format(datetime.now().strftime("%d.%m.%Y %H:%M")))

    # Forecast-Status: zeigt ob Bedarf aus Historie oder statisch
    _hist = lese_verkaufshistorie()
    _n_wochen = len(_hist.get("wochen", []))
    if _n_wochen > 0:
        _alpha = _hist.get("_alpha", 0.3)
        html.append('<div style="margin:10px 12px 0;padding:6px 12px;background:#E1F0FF;color:#1B5FAB;'
                    'border-radius:8px;font-size:11px;font-weight:600">'
                    '&#128200; Forecast aktiv: EWMA &alpha;={} ueber {} Wochen Historie</div>'.format(_alpha, _n_wochen))
    else:
        html.append('<div style="margin:10px 12px 0;padding:6px 12px;background:#FFF4D4;color:#8A6D1F;'
                    'border-radius:8px;font-size:11px;font-weight:600">'
                    '&#128221; Statische Verkaufszahlen — noch keine Historie. '
                    'Mit <code>System/lerne_woche.py</code> Wochen einlernen.</div>')

    # ── Chargen-Matrix Produkt x Tag ───────────────────────────────────────
    # Logik:
    #   1) Plan-Basis pro (Tag, Produkt) sammeln:
    #      - Aktive Tage (in vtage): aus current wplan
    #      - Vergangene Tage: aus gespeichertem wplan_KW snapshot
    #   2) ist_produktion-Overlay drueberlegen: ist_menge ersetzt Plan-Menge,
    #      verschoben_nach verschiebt Eintrag auf Zieltag, erledigt-Flag
    #      markiert Zelle gruen.
    #   3) Chargen pro Zelle aus PRODUKT_CONFIG.max_ofen neu berechnen,
    #      damit eine geaenderte Menge auch die Chargenzahl korrekt zeigt.
    vtage_set       = set(vtage or [])
    saved_wplan_kw  = lese_wplan_kw()
    ist_produktion  = lese_ist_produktion()

    # plan[(tag, pk)] = Plan-Stueckzahl
    plan = defaultdict(int)
    for tag in ARBEITSTAGE:
        if tag in vtage_set:
            quelle = wplan.get(tag, [])
        else:
            quelle = saved_wplan_kw.get(tag, [])
        for a in (quelle or []):
            pk = a.get("produkt_key")
            if not pk or a.get("excel_only"): continue
            if pk not in PRODUKT_CONFIG: continue
            stk = int(a.get("menge", 0) or 0)
            if stk > 0:
                plan[(tag, pk)] += stk

    # finale[(tag, pk)] = {stk, erledigt, geaendert}
    finale = {}
    for key, stk in plan.items():
        finale[key] = {"stk": stk, "erledigt": False, "geaendert": False}

    # Overlay aus ist_produktion: Menge, Erledigt-Flag, Verschiebung
    for tag, eintraege in (ist_produktion or {}).items():
        if tag not in ARBEITSTAGE: continue
        for pk, rec in (eintraege or {}).items():
            if pk not in PRODUKT_CONFIG: continue
            if not isinstance(rec, dict): continue
            verschoben_nach = rec.get("verschoben_nach")
            ist_menge       = rec.get("ist_menge")
            erledigt        = bool(rec.get("erledigt"))
            key = (tag, pk)

            if verschoben_nach and verschoben_nach in ARBEITSTAGE:
                # Quelle entfernen, Ziel hinzufuegen
                quell_menge = (int(ist_menge) if isinstance(ist_menge, (int, float))
                               else plan.get(key, 0))
                if key in finale:
                    del finale[key]
                if quell_menge > 0:
                    ziel_key = (verschoben_nach, pk)
                    zelle = finale.setdefault(ziel_key,
                        {"stk": 0, "erledigt": False, "geaendert": True})
                    zelle["stk"] += quell_menge
                    zelle["geaendert"] = True
                    if erledigt:
                        zelle["erledigt"] = True
                continue

            # Normaler Overlay (kein Verschieben)
            zelle = finale.get(key)
            if zelle is None:
                if isinstance(ist_menge, (int, float)) and ist_menge > 0:
                    finale[key] = {"stk": int(ist_menge),
                                   "erledigt": erledigt, "geaendert": True}
                elif erledigt and plan.get(key, 0) > 0:
                    finale[key] = {"stk": plan[key],
                                   "erledigt": True, "geaendert": False}
            else:
                if isinstance(ist_menge, (int, float)):
                    neu = int(ist_menge)
                    if neu != zelle["stk"]:
                        zelle["geaendert"] = True
                    zelle["stk"] = neu
                if erledigt:
                    zelle["erledigt"] = True

    # In Matrix-Form ueberfuehren (mit Chargen-Neuberechnung aus max_ofen)
    matrix = {}  # {pk: {tag: {stk, ch, erledigt, geaendert}}}
    for (tag, pk), c in finale.items():
        if c["stk"] <= 0: continue
        cfg = PRODUKT_CONFIG.get(pk, {})
        max_ofen = cfg.get("max_ofen") or c["stk"] or 1
        ch = max(1, math.ceil(c["stk"] / max(1, max_ofen)))
        row = matrix.setdefault(pk, {})
        zelle = row.setdefault(tag,
            {"stk": 0, "ch": 0, "erledigt": False, "geaendert": False})
        zelle["stk"] += c["stk"]
        zelle["ch"]  += ch
        if c["erledigt"]:  zelle["erledigt"]  = True
        if c["geaendert"]: zelle["geaendert"] = True

    if matrix:
        produkt_order = [pk for pk in PRODUKT_CONFIG if pk in matrix]
        kurz = {"Montag":"Mo","Dienstag":"Di","Mittwoch":"Mi","Donnerstag":"Do","Freitag":"Fr"}
        html.append('<div class="matrix">')
        html.append('<div class="matrix-title">&#128202; Chargen-Verteilung</div>')
        html.append('<table><thead><tr><th>Produkt</th>')
        for t in ARBEITSTAGE:
            html.append('<th>{}</th>'.format(kurz[t]))
        html.append('<th class="sum">&Sigma;</th></tr></thead><tbody>')
        for pk in produkt_order:
            row = matrix[pk]
            html.append('<tr><td>{}</td>'.format(pk))
            stk_sum = 0; ch_sum = 0
            for t in ARBEITSTAGE:
                zelle = row.get(t)
                ist_vergangen = vtage_set and t not in vtage_set
                if not zelle or zelle["stk"] == 0:
                    cls = ' class="z past"' if ist_vergangen else ' class="z"'
                    html.append('<td{}>&middot;</td>'.format(cls))
                else:
                    stk_sum += zelle["stk"]; ch_sum += zelle["ch"]
                    klassen = []
                    if zelle.get("erledigt"):  klassen.append("done")
                    if zelle.get("geaendert"): klassen.append("mod")
                    if ist_vergangen and not zelle.get("erledigt"):
                        klassen.append("past")
                    cls = ' class="' + ' '.join(klassen) + '"' if klassen else ''
                    html.append('<td{}><b>{}</b><br><span style="color:#888;font-size:10px">{}&times;</span></td>'.format(
                        cls, zelle["stk"], zelle["ch"]))
            html.append('<td class="sum">{}<br><span style="font-weight:400;font-size:10px">{}&times;</span></td>'.format(stk_sum, ch_sum))
            html.append('</tr>')
        html.append('</tbody></table></div>')
        html.append('<div class="matrix-legend">'
                    '<span class="lg-done">2</span> erledigt '
                    '&nbsp;&middot;&nbsp; &#9998; vom Plan abweichend '
                    '&nbsp;&middot;&nbsp; &middot; nichts geplant'
                    '</div>')

    hat_inhalt = False
    for tag in ARBEITSTAGE:
        aufgaben = [a for a in wplan.get(tag, []) if not a.get("excel_only")]
        if not aufgaben:
            continue
        hat_inhalt = True
        bg, fg = FARBEN.get(tag, ("#eee","#333"))

        gesamt_aktiv = sum(a.get("aktiv_min",0) or 0 for a in aufgaben)
        gesamt_passiv = sum(a.get("passiv_min",0) or 0 for a in aufgaben)

        html.append('<div class="day-card">')
        ueberlast = gesamt_aktiv > MAX_AKTIV_MIN_PRO_TAG
        # Bei Ueberlast: roten Rand und ein Warn-Symbol
        extra = ';border-left:4px solid #C0392B' if ueberlast else ''
        html.append('<div class="day-header" style="background:{};color:{}{}">{}'.format(bg, fg, extra, tag.upper()))
        if gesamt_aktiv:
            warn = ' &#9888;' if ueberlast else ''
            html.append(' &nbsp;<span style="font-weight:400;font-size:11px">&#9201; aktiv {} / passiv {}{}</span>'.format(
                zeit_str(gesamt_aktiv), zeit_str(gesamt_passiv), warn))
        html.append('</div>')
        if ueberlast:
            html.append('<div style="background:#FFE4E1;color:#C0392B;padding:6px 14px;font-size:11px;font-weight:600">'
                        '&#9888; Tageslimit {} Min ueberschritten (+{} Min) — manuell entzerren oder Verteilung pruefen</div>'.format(
                MAX_AKTIV_MIN_PRO_TAG, int(gesamt_aktiv - MAX_AKTIV_MIN_PRO_TAG)))

        GRUND_BADGES = {
            "pflicht": ('g-pflicht', 'Pflicht'),
            "bestand": ('g-bestand', 'Bestand'),
            "frische": ('g-frische', 'Frische'),
            "puffer":  ('g-puffer',  'Puffer'),
        }
        for a in aufgaben:
            prod = a.get("produkt","")
            menge = a.get("menge")
            notiz = a.get("notiz","")
            aktiv = a.get("aktiv_min",0)
            passiv = a.get("passiv_min",0)
            frisch = a.get("frisch_stk")
            einfrier = a.get("einfrier_stk")
            grund = a.get("grund")
            badge_html = ""
            if grund and grund in GRUND_BADGES:
                cls, label = GRUND_BADGES[grund]
                badge_html = '<span class="grund-badge {}">{}</span>'.format(cls, label)
            html.append('<div class="task-row"><div style="flex:1">')
            if menge and int(menge) > 0:
                html.append('<div class="task-name">{} &mdash; <span style="color:#555;font-size:13px">{} Stk</span>{}</div>'.format(prod, int(menge), badge_html))
            else:
                html.append('<div class="task-name">{}{}</div>'.format(prod, badge_html))
            if notiz:
                html.append('<div class="task-notiz">{}</div>'.format(notiz))
            if einfrier is not None and (einfrier > 0 or (frisch is not None and frisch < int(menge or 0))):
                html.append('<div class="task-notiz" style="color:#2E75B6">&#10052; {} frisch / {} einfrieren</div>'.format(int(frisch or 0), int(einfrier or 0)))
            if aktiv:
                html.append('<div class="task-zeit">&#9201; aktiv {} / passiv {}</div>'.format(
                    zeit_str(aktiv), zeit_str(passiv)))
            html.append('</div></div>')
        html.append('</div>')

        warnungen = tages_info.get(tag, [])
        if warnungen:
            html.append('<div class="lager-panel">')
            html.append('<div class="lager-title">&#128230; Lagerstand nach {}</div>'.format(tag))
            for w in warnungen:
                cls = "lager-leer" if w["status"]=="leer" else "lager-niedrig"
                icon = "&#128308;" if w["status"]=="leer" else "&#129473;"
                html.append('<div class="lager-row"><span class="{}">{} {}</span><span>{:.1f} Stk</span></div>'.format(
                    cls, icon, w["produkt"], w["bestand"]))
            html.append('</div>')

    if not hat_inhalt:
        html.append('<p style="text-align:center;color:#aaa;font-size:14px;padding:40px 20px">Keine Aufgaben diese Woche.</p>')

    html.append('<div class="lager-panel" style="margin-top:16px">')
    html.append('<div class="lager-title">&#128203; Wochenbedarf</div>')
    for key, bd in bedarf.items():
        zu_prod = bd.get("zu_produzieren", 0)
        im_lager = bd.get("im_lager", 0)
        if zu_prod > 0:
            html.append('<div class="lager-row"><span>{}</span><span style="color:#375623;font-weight:600">+{} Stk produzieren</span></div>'.format(key, zu_prod))
        else:
            html.append('<div class="lager-row"><span>{}</span><span style="color:#888">Lager reicht ({} Stk)</span></div>'.format(key, int(im_lager)))
    html.append('</div>')

    html.append('</body></html>')

    fname = OUTPUT_PFAD / "Wochenuebersicht.html"
    fname.write_text("\n".join(html), encoding="utf-8")
    print("Wochenuebersicht HTML: {}".format(fname.name))
    return fname

# ============================================================
# OUTPUT HTML
# ============================================================
HTML_CSS = (
    "*{box-sizing:border-box;margin:0;padding:0}"
    "body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;"
    "background:#f0f2f5;color:#1a1a2e;padding:12px;font-size:15px}"
    "h1{font-size:22px;font-weight:700;color:#2E75B6;text-align:center;padding:12px;"
    "background:white;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:4px}"
    ".datum{text-align:center;color:#666;font-size:13px;margin-bottom:4px}"
    ".zeitbadge{text-align:center;margin-bottom:14px}"
    ".zeitbadge span{display:inline-block;background:#2E75B6;color:white;border-radius:20px;"
    "padding:4px 14px;font-size:13px;font-weight:600;margin:2px}"
    ".zeitbadge .passiv{background:#8AAECC}"
    ".task{background:white;border-radius:12px;padding:14px;margin-bottom:10px;"
    "box-shadow:0 1px 4px rgba(0,0,0,.08);border-left:4px solid #2E75B6;cursor:pointer}"
    ".task.special{border-left-color:#375623}"
    ".task.prep{border-left-color:#9B59B6;background:#F9F0FF}"
    ".task-header{font-weight:700;font-size:16px;margin-bottom:4px;display:flex;justify-content:space-between;align-items:flex-start;gap:6px}"
    ".task-title-text{flex:1}"
    ".task-toggle{font-size:13px;color:#bbb;flex-shrink:0;padding-top:2px;transition:transform .2s;display:inline-block;line-height:1}"
    ".task.collapsed .task-toggle{transform:rotate(-90deg)}"
    ".task-body{}"
    ".task.collapsed .task-body{display:none}"
    ".task.done{opacity:0.55;border-left-color:#ccc!important}"
    ".task.done .task-title-text{text-decoration:line-through;color:#aaa}"
    ".task-note{color:#555;font-size:13px;margin-bottom:6px}"
    ".task-zeit{color:#2E75B6;font-size:12px;font-weight:600;margin-bottom:6px}"
    ".recipe{background:#F0F7F0;border-radius:8px;padding:10px;margin-top:8px;border:1px solid #C6DFC6}"
    ".recipe-title{font-weight:700;color:#375623;font-size:13px;margin-bottom:6px}"
    ".rzp-sec{font-weight:600;color:#666;font-size:12px;text-transform:uppercase;"
    "letter-spacing:0.5px;margin:8px 0 3px;border-top:1px solid #D9EAD3;padding-top:5px}"
    ".zutaten{list-style:none;padding:0}"
    ".zutaten li{display:flex;justify-content:space-between;padding:5px 0;"
    "border-bottom:1px solid #e8f4e8;font-size:14px}"
    ".zutaten li:last-child{border-bottom:none}"
    ".ek-panel{background:#f0faf0;border:1px solid #b2dfdb;border-radius:8px;padding:12px 14px;margin-top:16px}"
    ".ek-title{font-weight:700;color:#2e7d32;margin-bottom:8px;font-size:.95em}"
    ".ek-table{width:100%;border-collapse:collapse;font-size:.87em}"
    ".ek-table th{background:#43a047;color:#fff;padding:6px 8px;text-align:left}"
    ".ek-table td{padding:5px 8px;border-bottom:1px solid #c8e6c9}"
    ".ek-table tr:last-child td{border-bottom:none}"
    ".z-menge,.z-kosten{text-align:right}"
    ".ek-total td{font-weight:700;background:#c8e6c9}"
    ".zutat{color:#333}.menge{font-weight:600;color:#375623}"
    ".sub-recipe{background:#FFF8E1;border-radius:8px;padding:10px;margin-top:6px;border:1px solid #FFE082}"
    ".sub-recipe-title{font-weight:700;color:#7B5800;font-size:12px;margin-bottom:5px;text-transform:uppercase;letter-spacing:0.3px}"
    ".footer{text-align:center;color:#aaa;font-size:11px;margin-top:16px}"
    ".lager-panel{background:white;border-radius:12px;padding:12px;margin-top:10px;border-left:4px solid #e67e22;box-shadow:0 1px 4px rgba(0,0,0,.08)}"
    ".lager-panel-title{font-weight:700;font-size:13px;color:#888;margin-bottom:8px}"
    ".lager-item{display:flex;justify-content:space-between;align-items:center;padding:5px 0;font-size:13px;border-bottom:1px solid #f5f5f5}"
    ".lager-item:last-child{border-bottom:none}"
    ".lager-leer .lager-name{color:#c0392b;font-weight:700}"
    ".lager-leer .lager-bestand{color:#c0392b;font-weight:700}"
    ".lager-niedrig .lager-name{color:#e67e22}"
    ".lager-niedrig .lager-bestand{color:#e67e22;font-weight:600}"
    ".badge{display:inline-block;font-size:11px;font-weight:700;padding:2px 8px;"
    "border-radius:10px;margin-left:8px;vertical-align:middle;white-space:nowrap}"
    ".badge-notwendig{background:#FDECEA;color:#C62828;border:1px solid #EF9A9A}"
    ".badge-puffer{background:#FFF8E1;color:#7B6200;border:1px solid #FDD835}"
)

def erstelle_html(tag, aufgaben, murt, rezepte, sub_rezepte=None, lager_info=None, ist_heut=False, preise=None):
    aufgaben_html = aufgaben  # bereits nach prio sortiert
    inhalt=""

    def fmt_min(m):
        if m >= 60: return "{}h{}".format(m//60, ("{}min".format(m%60) if m%60 else ""))
        return "{}min".format(m)

    def _zeit_badge(am, pm):
        if not am: return ""
        parts = ["&#9998; {}".format(fmt_min(am))]
        if pm: parts.append("&#9201; {} passiv".format(fmt_min(pm)))
        return '<div class="task-zeit">{}</div>'.format(" &nbsp;|&nbsp; ".join(parts))

    # Gesamtzeitaufwand des Tages summieren
    gesamt_aktiv  = sum(a.get("aktiv_min",0)  for a in aufgaben_html if not a.get("excel_only"))
    gesamt_passiv = sum(a.get("passiv_min",0) for a in aufgaben_html if not a.get("excel_only"))

    # ---- Tasks zu Gruppen zusammenfassen (1 Karte pro Produkt pro Tag) ----
    def _group_key_for(a):
        p_up = (a.get("produkt") or "").upper()
        if "MUERBETEIG ANSETZEN" in p_up: return "_mt_ansetzen"
        if "MUERBETEIG BACKEN"   in p_up: return "_mt_backen"
        if "STREUSEL HERSTELLEN" in p_up: return "_streusel_haupt"
        if p_up.startswith("STREUSEL"):   return "_streusel_chargen"
        pk = a.get("produkt_key")
        if pk: return "p:" + pk
        return "x:" + (a.get("produkt") or "")

    groups = []  # erhalten Reihenfolge der Erst-Vorkommen
    g_idx = {}
    for a in aufgaben_html:
        if a.get("excel_only"): continue
        k = _group_key_for(a)
        if k in g_idx:
            g = groups[g_idx[k]]
        else:
            g_idx[k] = len(groups)
            g = {"key": k, "first": a, "aufgaben": [],
                 "total_menge": 0, "total_aktiv": 0, "total_passiv": 0}
            groups.append(g)
        g["aufgaben"].append(a)
        g["total_menge"]  += int(a.get("menge") or 0)
        g["total_aktiv"]  += int(a.get("aktiv_min") or 0)
        g["total_passiv"] += int(a.get("passiv_min") or 0)

    def _attr_json(obj):
        return json.dumps(obj, ensure_ascii=False).replace('"', '&quot;').replace("'", '&#39;')

    def _rezept_to_json(rzp):
        if not rzp or not rzp.get('sections'): return None
        secs = []
        for sec in rzp['sections']:
            zts = []
            for z in sec.get('zutaten', []):
                zts.append({"z": z['zutat'], "m": z['menge'], "e": z.get('einheit','g')})
            secs.append({"name": sec.get('name',''), "zutaten": zts})
        return {"sections": secs, "max_charge": rzp.get('max_charge_num',1) or 1}

    # Pro-Tag-Mengen pro Produkt (fuer Rezept-Skalierung, identisch zu group totals)
    tagesmengen = {g["key"]: g["total_menge"] for g in groups}

    # ---- Gruppen rendern ----
    for g in groups:
        k = g["key"]
        first = g["first"]
        n = first.get("notiz") or ""
        a_min = g["total_aktiv"]
        p_min = g["total_passiv"]
        zeit_html = _zeit_badge(a_min, p_min)

        # --- Muerbeteig ansetzen ---
        if k == "_mt_ansetzen":
            gr=murt["grundrezepte"]
            zutaten_html="".join(
                '<li><span class="zutat" data-z="{z}">{z}</span><span class="menge" data-per="{m}">{m}g</span></li>'.format(z=z,m=mg)
                for z,mg in MURBETEIG_REZEPT.items()
            )
            recipe_data = {"sections":[{"name":"","zutaten":[{"z":z,"m":mg,"e":"g"} for z,mg in MURBETEIG_REZEPT.items()]}],"max_charge":1}
            inhalt+=(
                '<div class="task special" data-tag="{tag}" data-key="_mt_ansetzen" data-typ="hilfe" '
                'data-plan-menge="{gr}" data-recipe="{rcp}">'
                '<div class="task-header" onclick="toggleTask(event,this)">'
                '<span class="task-title-text">Muerbeteig ansetzen &mdash; <span class="ist-menge">{gr}</span>x Grundrezept</span>'
                '<span class="task-toggle">&#9660;</span></div>'
                '<div class="task-body">'
                '<div class="task-note">{n}</div>{zb}'
                '<div class="recipe"><div class="recipe-title">Grundrezept (1x ansetzen, <span class="repeat-n">{gr}</span>x wiederholen)</div>'
                '<ul class="zutaten">{zh}</ul></div>'
                '</div></div>'
            ).format(tag=tag,gr=gr,rcp=_attr_json(recipe_data),n=n,zb=zeit_html,zh=zutaten_html)
            continue

        # --- Muerbeteig backen ---
        if k == "_mt_backen":
            portionen_html="".join(
                '<li><span class="zutat">{}x {}</span><span class="menge">{}g</span></li>'.format(
                    v["anzahl"],t,int(v["gramm"]))
                for t,v in murt["portionen"].items()
            )
            res=murt.get("kk_montag_reserve",0)
            if res:
                portionen_html+='<li><span class="zutat">{}x Mürbeteig rund (fuer naechsten Montag)</span><span class="menge">{}g</span></li>'.format(
                    res,res*MURBETEIG_PORTIONEN.get("Mürbeteig rund", 180))
            inhalt+=(
                '<div class="task special" data-tag="{tag}" data-key="_mt_backen" data-typ="hilfe" data-plan-menge="1">'
                '<div class="task-header" onclick="toggleTask(event,this)">'
                '<span class="task-title-text">Muerbeteig backen &mdash; alle Boeden heute</span>'
                '<span class="task-toggle">&#9660;</span></div>'
                '<div class="task-body">'
                '<div class="task-note">180&deg;C, 7&ndash;9 Minuten</div>{zb}'
                '<div class="recipe"><div class="recipe-title">Boeden aufteilen</div>'
                '<ul class="zutaten">{ph}</ul></div>'
                '</div></div>'
            ).format(tag=tag,zb=zeit_html,ph=portionen_html)
            continue

        # --- Streusel herstellen (zentral mit Rezept) ---
        if k == "_streusel_haupt":
            sr_key = "Streusel"
            srz = (sub_rezepte or {}).get(sr_key)
            n_mal_str = n.split("x ")[0] if "x " in n else "1"
            try: n_mal_s = int(n_mal_str.strip())
            except: n_mal_s = 1
            streusel_recipe_html = ''
            if srz:
                streusel_recipe_html = format_sub_rezept_html(sr_key, srz, n_mal_s)
            rd = _rezept_to_json(srz) if srz else None
            rcp_attr = ' data-recipe="{}"'.format(_attr_json(rd)) if rd else ''
            inhalt += (
                '<div class="task special" data-tag="{tag}" data-key="_streusel" data-typ="hilfe" '
                'data-plan-menge="{m}"{rcp}>'
                '<div class="task-header" onclick="toggleTask(event,this)">'
                '<span class="task-title-text">Streusel herstellen &mdash; <span class="ist-menge">{m}</span>x Grundrezept</span>'
                '<span class="task-toggle">&#9660;</span></div>'
                '<div class="task-body">'
                '<div class="task-note">{n}</div>{zb}{rh}'
                '</div></div>'
            ).format(tag=tag,m=n_mal_s,rcp=rcp_attr,n=n,zb=zeit_html,rh=streusel_recipe_html)
            continue

        # --- Streusel Chargen (mehrere zusammengefasst) ---
        if k == "_streusel_chargen":
            anzahl = len(g["aufgaben"])
            inhalt += (
                '<div class="task" data-tag="{tag}" data-key="_streusel_chargen" data-typ="hilfe" data-plan-menge="{m}">'
                '<div class="task-header" onclick="toggleTask(event,this)">'
                '<span class="task-title-text">STREUSEL &mdash; <span class="ist-menge">{m}</span> Chargen</span>'
                '<span class="task-toggle">&#9660;</span></div>'
                '<div class="task-body">'
                '<div class="task-note">{n}</div>{zb}'
                '</div></div>'
            ).format(tag=tag,m=anzahl,n=n,zb=zeit_html)
            continue

        # --- Normaler Produkt-Eintrag (alle Chargen eines Produkts zusammen) ---
        p = first["produkt"]; pk = first.get("produkt_key")
        m = g["total_menge"]
        is_pistazien = "PISTAZIEN" in (p or "").upper()
        cls_extra = " prep" if is_pistazien else ""

        notiz_extra=""
        if "letzter Woche" in n:
            notiz_extra = ("Murbeteigboeden von dieser Woche verwenden (Dienstag gebacken)"
                           if ist_heut else "Murbeteigboeden von letzter Woche verwenden")

        # Charge-Infos aller Aufgaben dieses Produkts zusammenfuegen
        charge_infos = []
        for a in g["aufgaben"]:
            ci = (a.get("notiz") or "").split(" - ")[0]
            if ci and ci not in charge_infos:
                charge_infos.append(ci)
        charge_info = " + ".join(charge_infos)

        notwendig_val = first.get("notwendig")
        if notwendig_val is True:
            badge_html = '<span class="badge badge-notwendig">&#128308; Notwendig</span>'
        elif notwendig_val is False:
            badge_html = '<span class="badge badge-puffer">&#128993; Puffer aufbauen</span>'
        else:
            badge_html = ""

        # Rezept einmal pro Produkt pro Tag (entfaellt fuer Pistazien wie bisher)
        rezept_html = ""
        rcp_attr = ""
        sub_attr = ""
        if pk and not is_pistazien:
            rzp = (rezepte or {}).get(pk)
            if rzp:
                rezept_html = format_rezept_block_html(rzp, m)
                rd = _rezept_to_json(rzp)
                if rd: rcp_attr = ' data-recipe="{}"'.format(_attr_json(rd))
                if sub_rezepte:
                    skip = {"Streusel"} if tag != "Montag" else set()
                    rezept_html += format_alle_sub_rezepte_html(rzp, m, sub_rezepte, skip_keys=skip)
                    # Sub-Rezepte referenziert vom Hauptrezept (fuer Client-Rescaling)
                    used_subs = {}
                    for sec in rzp.get('sections', []):
                        for z in sec.get('zutaten', []):
                            if z.get('einheit') == 'ref':
                                sub_key = z['zutat']
                                if sub_key in (sub_rezepte or {}) and sub_key not in skip:
                                    sub_data = _rezept_to_json(sub_rezepte[sub_key])
                                    if sub_data: used_subs[sub_key] = sub_data
                    if used_subs:
                        sub_attr = ' data-subrecipe="{}"'.format(_attr_json(used_subs))

        if m and m > 0:
            head = '<span class="task-title-text">{} &mdash; <span class="ist-menge">{}</span> Stk{}</span>'.format(p,int(m),badge_html)
        else:
            head = '<span class="task-title-text">{}{}</span>'.format(p,badge_html)

        body_extras = ""
        if charge_info: body_extras += '<div class="task-note charge-note">{}</div>'.format(charge_info)
        if notiz_extra: body_extras += '<div class="task-note" style="color:#8B4513">{}</div>'.format(notiz_extra)
        if zeit_html:   body_extras += zeit_html

        inhalt += (
            '<div class="task{cls}" data-tag="{tag}" data-key="{pk}" data-typ="produkt" '
            'data-plan-menge="{m}"{rcp}{sub}>'
            '<div class="task-header" onclick="toggleTask(event,this)">{head}<span class="task-toggle">&#9660;</span></div>'
            '<div class="task-body">{body}{rh}</div>'
            '</div>'
        ).format(cls=cls_extra,tag=tag,pk=(pk or first.get("produkt") or ""),
                 m=int(m),rcp=rcp_attr,sub=sub_attr,head=head,body=body_extras,rh=rezept_html)

    # --- Lager-Status-Panel ---
    if lager_info:
        panel_items = ""
        for entry in lager_info:
            cls = "lager-" + entry["status"]
            icon = "&#128308;" if entry["status"] == "leer" else "&#128993;"
            prod_hint = " (+{} prod.)".format(int(entry["produziert"])) if entry["produziert"] else ""
            panel_items += (
                '<div class="lager-item {cls}">'
                '<span class="lager-name">{icon} {produkt}{hint}</span>'
                '<span class="lager-bestand">{bestand} Stk</span>'
                '</div>'
            ).format(
                cls=cls, icon=icon,
                produkt=entry["produkt"], hint=prod_hint,
                bestand=entry["bestand"]
            )
        if panel_items:
            inhalt += (
                '<div class="lager-panel">'
                '<div class="lager-panel-title">&#128202; Lagerbestand Ende {}:</div>'
                '{}'
                '</div>'
            ).format(tag, panel_items)

    if not inhalt:
        inhalt='<div class="task"><div class="task-header">Kein Produktionstag</div></div>'

    # Einkaufsliste berechnen
    einkauf_html = ""
    if preise and rezepte:
        _zt = berechne_zutaten(aufgaben_html, rezepte, sub_rezepte or {})
        einkauf_html = erstelle_html_einkaufsliste(tag, _zt, preise)

    kw=datetime.now().isocalendar()[1]; datum=datetime.now().strftime("%d.%m.%Y")
    # Zeitbadge aufbauen
    zeitbadge=""
    if gesamt_aktiv > 0:
        zeitbadge = (
            '<div class="zeitbadge">'
            '<span>&#9998; Aktiv: {}</span>'
            '{}'
            '</div>'
        ).format(
            fmt_min(gesamt_aktiv),
            '<span class="passiv">&#9201; Passiv: {}</span>'.format(fmt_min(gesamt_passiv)) if gesamt_passiv else ""
        )
    heute_banner = (
        '<div style="text-align:center;background:#fff3cd;border-radius:8px;padding:6px 12px;'
        'font-size:12px;color:#856404;margin-bottom:8px;font-weight:600">'
        '&#9888; Neuplanung ab heute &mdash; vergangene Tage dieser Woche nicht enthalten'
        '</div>'
    ) if ist_heut else ""
    toggle_js = """<script>
(function(){
var DAYS=['Montag','Dienstag','Mittwoch','Donnerstag','Freitag','Samstag','Sonntag'];
var IST={kw:'',tage:{},_sha:null};
var KW='';

function cfg(k){return localStorage.getItem('kp_'+k)||'';}
function ghBase(){return 'https://api.github.com/repos/'+cfg('owner')+'/'+cfg('repo');}
function ghReady(){return cfg('owner')&&cfg('repo')&&cfg('token');}

function getKW(){
  var dt=document.querySelector('.datum');
  if(dt){
    var m=dt.textContent.match(/KW\\s*(\\d+)/i);
    var ym=dt.textContent.match(/(\\d{2})\\.(\\d{2})\\.(\\d{4})/);
    if(m){
      var jahr=ym?ym[3]:String(new Date().getFullYear());
      var w=m[1]; if(w.length<2) w='0'+w;
      return jahr+'-'+w;
    }
  }
  return String(new Date().getFullYear())+'-00';
}

function showToast(msg,kind){
  var d=document.createElement('div');
  d.className='ip-toast '+(kind||'');
  d.textContent=msg;
  document.body.appendChild(d);
  setTimeout(function(){d.classList.add('show');},10);
  setTimeout(function(){d.classList.remove('show');setTimeout(function(){d.remove();},300);},2200);
}

function setIst(tag,key,fields){
  if(!IST.tage[tag]) IST.tage[tag]={};
  if(!IST.tage[tag][key]) IST.tage[tag][key]={};
  Object.assign(IST.tage[tag][key],fields);
}
function getIst(tag,key){
  return (IST.tage[tag]||{})[key]||{};
}

window.toggleTask=function(e,hdrEl){
  if(e && e.target){
    var tgt=e.target;
    if(tgt.closest('.done-btn,.gear-btn,.gear-panel,input,select,button')) return;
  }
  var el=hdrEl.closest('.task');
  if(el) el.classList.toggle('collapsed');
};

window.markDone=function(e,btn){
  e.stopPropagation();
  var el=btn.closest('.task');
  el.classList.toggle('done');
  var isDone=el.classList.contains('done');
  if(isDone){
    el.classList.add('collapsed');
    var rec=getIst(el.dataset.tag,el.dataset.key);
    if(typeof rec.ist_menge!=='number'){
      setIst(el.dataset.tag,el.dataset.key,{ist_menge:parseFloat(el.dataset.planMenge)});
    }
    setIst(el.dataset.tag,el.dataset.key,{erledigt:true});
  }else{
    el.classList.remove('collapsed');
    setIst(el.dataset.tag,el.dataset.key,{erledigt:false});
  }
  saveIst();
};

window.openGear=function(e,btn){
  e.stopPropagation();
  var el=btn.closest('.task');
  var panel=el.querySelector('.gear-panel');
  if(!panel) return;
  if(panel.style.display==='block'){panel.style.display='none';return;}
  var rec=getIst(el.dataset.tag,el.dataset.key);
  var stkInp=panel.querySelector('.stk-input');
  var cur=(typeof rec.ist_menge==='number')?rec.ist_menge:parseFloat(el.dataset.planMenge);
  stkInp.value=cur;
  var daySel=panel.querySelector('.day-select');
  daySel.value=rec.verschoben_nach||el.dataset.tag;
  panel.style.display='block';
  el.classList.remove('collapsed');
};

window.saveGearChanges=function(btn){
  var panel=btn.closest('.gear-panel');
  var el=panel.closest('.task');
  var stk=parseFloat(panel.querySelector('.stk-input').value);
  if(isNaN(stk)||stk<0) stk=0;
  applyStk(el,stk);
  applyVerschieben(el,panel.querySelector('.day-select').value);
  panel.style.display='none';
  saveIst();
};

window.cancelGear=function(btn){
  btn.closest('.gear-panel').style.display='none';
};

function applyStk(el,newStk){
  setIst(el.dataset.tag,el.dataset.key,{ist_menge:newStk});
  setStkDisplay(el,newStk);
}

function applyVerschieben(el,newTag){
  if(newTag===el.dataset.tag){
    setIst(el.dataset.tag,el.dataset.key,{verschoben_nach:null});
    el.classList.remove('verschoben');
    var b=el.querySelector('.verschoben-banner');if(b)b.remove();
  }else{
    setIst(el.dataset.tag,el.dataset.key,{verschoben_nach:newTag});
    el.classList.add('verschoben');
    addVerschobenBanner(el,'auf '+newTag);
  }
}

function addVerschobenBanner(el,txt){
  var existing=el.querySelector('.verschoben-banner');
  if(existing) existing.remove();
  var b=document.createElement('div');
  b.className='verschoben-banner';
  b.textContent='↪ Verschoben '+txt;
  var body=el.querySelector('.task-body');
  if(body) body.insertBefore(b,body.firstChild);
}

function setStkDisplay(el,newStk){
  el.querySelectorAll('.ist-menge').forEach(function(s){s.textContent=newStk;});
  var rcpStr=el.dataset.recipe;
  if(rcpStr){
    try{renderRecipe(el,JSON.parse(rcpStr),newStk);}catch(e){console.warn('recipe',e);}
  }
  var subStr=el.dataset.subrecipe;
  if(subStr){
    try{renderSubRecipes(el,JSON.parse(subStr),newStk);}catch(e){console.warn('subrecipe',e);}
  }
  if(newStk<=0){el.classList.add('ist-null');}
  else{el.classList.remove('ist-null');}
}

function fmtGramm(v){
  v=Math.round(v*10)/10;
  return (v===Math.floor(v))?Math.floor(v)+'g':v+'g';
}

function renderRecipe(el,rcp,stk){
  var rec=el.querySelector('.recipe');
  if(!rec || !rcp.sections) return;
  var ul=rec.querySelector('.zutaten');
  if(!ul) return;
  var html='';
  rcp.sections.forEach(function(sec){
    if(rcp.sections.length>1 && sec.name){
      html+='<div class="rzp-sec">'+sec.name+'</div>';
    }
    sec.zutaten.forEach(function(z){
      if(z.e==='ref'){
        html+='<li><span class="zutat">'+z.z+'</span><span class="menge">'+z.m+'  (vorh. backen/vorbereiten)</span></li>';
      }else{
        html+='<li><span class="zutat">'+z.z+'</span><span class="menge">'+fmtGramm(z.m*stk)+'</span></li>';
      }
    });
  });
  ul.innerHTML=html;
  var t=rec.querySelector('.recipe-title');
  if(t){
    var mx=rcp.max_charge||1;
    var nMal=mx>0?Math.ceil(stk/mx):1;
    var batch=Math.min(mx,stk);
    if(stk<=0){t.innerHTML='Rezept (Ist-Menge: 0)';return;}
    var title=(batch>1?'Rezept fuer '+batch+' Stk':'Grundrezept (1 Stk)');
    title+=' &mdash; '+(nMal>1?'heute '+nMal+'x ansetzen':'1x ansetzen');
    t.innerHTML=title;
  }
}

function renderSubRecipes(el,subs,stk){
  Object.keys(subs).forEach(function(key){
    var rcp=subs[key];
    el.querySelectorAll('.sub-recipe').forEach(function(b){
      var title=b.querySelector('.sub-recipe-title');
      if(!title) return;
      if(title.textContent.toUpperCase().indexOf(key.toUpperCase())<0) return;
      var ul=b.querySelector('.zutaten');
      if(!ul) return;
      var html='';
      rcp.sections.forEach(function(sec){
        sec.zutaten.forEach(function(z){
          html+='<li><span class="zutat">'+z.z+'</span><span class="menge">'+fmtGramm(z.m*stk)+'</span></li>';
        });
      });
      ul.innerHTML=html;
    });
  });
}

function applyIst(){
  document.querySelectorAll('.task[data-tag][data-key]').forEach(function(el){
    var rec=getIst(el.dataset.tag,el.dataset.key);
    if(rec.erledigt){el.classList.add('done','collapsed');}
    if(typeof rec.ist_menge==='number' && rec.ist_menge!==parseFloat(el.dataset.planMenge)){
      setStkDisplay(el,rec.ist_menge);
    }
    if(rec.verschoben_nach){
      el.classList.add('verschoben');
      addVerschobenBanner(el,'auf '+rec.verschoben_nach);
    }
    if(rec.verschoben_von){
      addVerschobenBanner(el,'von '+rec.verschoben_von);
    }
  });
}

async function loadIst(){
  KW=getKW();
  if(!ghReady()){
    var indicator=document.getElementById('ip-status');
    if(indicator){indicator.textContent='⚠ GitHub-Login fehlt — Status wird nicht gespeichert';indicator.className='ip-status warn';}
    return;
  }
  try{
    var r=await fetch(ghBase()+'/contents/Wochenplan/ist_produktion_KW'+KW+'.json',{
      headers:{Authorization:'Bearer '+cfg('token'),Accept:'application/vnd.github+json'}
    });
    if(r.status===404){
      var ind=document.getElementById('ip-status');
      if(ind){ind.textContent='Tracking aktiv (KW '+KW+')';ind.className='ip-status ok';}
      return;
    }
    if(!r.ok) throw new Error('HTTP '+r.status);
    var j=await r.json();
    IST._sha=j.sha;
    var content=atob(j.content.replace(/\\s/g,''));
    var data=JSON.parse(decodeURIComponent(escape(content)));
    IST.kw=data.kw||KW;
    IST.tage=data.tage||{};
    applyIst();
    var ind2=document.getElementById('ip-status');
    if(ind2){ind2.textContent='Tracking aktiv (KW '+KW+')';ind2.className='ip-status ok';}
  }catch(e){
    console.warn('ist_produktion laden:',e);
    var ind=document.getElementById('ip-status');
    if(ind){ind.textContent='⚠ Fehler beim Laden: '+e.message;ind.className='ip-status warn';}
  }
}

var _saveDebounce=null;
async function saveIst(){
  clearTimeout(_saveDebounce);
  _saveDebounce=setTimeout(_doSaveIst,400);
}

// Holt die aktuelle SHA der ist_produktion-Datei. Bewusst minimalistische
// Headers (gleich wie loadIst), kein ref=main (nutzt Default-Branch), keine
// Cache-Buster — Mobile-Browser/Provider reagieren teils auf "If-None-Match"
// oder unbekannte Headers mit fehlerhaften CORS-Preflights.
// Returns: SHA-String wenn Datei existiert, null wenn 404, undefined bei Fehler.
async function _fetchCurrentSha(){
  try{
    var r=await fetch(ghBase()+'/contents/Wochenplan/ist_produktion_KW'+KW+'.json',{
      headers:{Authorization:'Bearer '+cfg('token'),Accept:'application/vnd.github+json'}
    });
    if(r.ok){var j=await r.json();return j.sha;}
    if(r.status===404) return null;
  }catch(e){console.warn('_fetchCurrentSha:',e);}
  return undefined;
}

async function _doSaveIst(retried){
  if(!ghReady()){showToast('GitHub-Login fehlt','error');return;}
  // Wenn IST._sha vom letzten erfolgreichen PUT vorhanden ist, vertrauen wir
  // dem. Sonst holen wir die SHA frisch — und blocken solange, bis wir
  // entweder eine SHA haben oder sicher wissen, dass die Datei noch nicht
  // existiert (404). Andernfalls riskieren wir, ohne SHA an eine existierende
  // Datei zu PUTten -> "sha wasn't supplied".
  if(!IST._sha){
    var sha=await _fetchCurrentSha();
    if(sha===undefined){
      // GET hat geknallt (Netz/CORS). Statt blind PUT zu senden, kurz
      // backoff und einmal nachfassen.
      await new Promise(function(res){setTimeout(res,500);});
      sha=await _fetchCurrentSha();
    }
    if(sha!==undefined) IST._sha=sha;  // string oder null
  }
  var payload={kw:KW,_format:1,_updated:new Date().toISOString(),tage:IST.tage};
  var jsonStr=JSON.stringify(payload,null,2);
  var b64=btoa(unescape(encodeURIComponent(jsonStr)));
  var body={message:'ist_produktion KW '+KW+' aktualisiert',content:b64};
  if(IST._sha) body.sha=IST._sha;
  try{
    var r=await fetch(ghBase()+'/contents/Wochenplan/ist_produktion_KW'+KW+'.json',{
      method:'PUT',
      headers:{Authorization:'Bearer '+cfg('token'),Accept:'application/vnd.github+json','Content-Type':'application/json'},
      body:JSON.stringify(body)
    });
    if(!r.ok){
      var er;try{er=await r.json();}catch(_){er={};}
      // Retry bei SHA-Konflikten und "wasn't supplied" — alles, was darauf
      // hinweist, dass die SHA aus irgendeinem Grund nicht stimmt.
      var shaErr = er.message && /sha|does not match|wasn'?t supplied|not supplied/i.test(er.message);
      if(!retried && (r.status===409 || r.status===422 || shaErr)){
        IST._sha=null;  // erzwingt fresh GET im naechsten Lauf
        await new Promise(function(res){setTimeout(res,500);});
        return _doSaveIst(true);
      }
      throw new Error(er.message||'HTTP '+r.status);
    }
    var rj=await r.json();IST._sha=rj.content.sha;
    showToast('Gespeichert','success');
  }catch(e){
    showToast('Speichern: '+e.message,'error');
  }
}

function _initOnce(){
  // Idempotent: in Dashboard.html laeuft das Script einmal pro Tab —
  // dieser Guard verhindert, dass jeder Task 5x Buttons bekommt.
  if(window.__ip_init_done) return;
  window.__ip_init_done = true;

  // Status-Indikator
  var status=document.createElement('div');
  status.id='ip-status';
  status.className='ip-status';
  status.textContent='Verbinde...';
  var h1=document.querySelector('h1');
  if(h1 && h1.parentNode){h1.parentNode.insertBefore(status,h1.nextSibling);}

  // Tasks mit Buttons + Panel ausstatten — Doppelt-Init pro Karte verhindern
  document.querySelectorAll('.task[data-tag][data-key]').forEach(function(el){
    if(el.dataset.ipReady==='1') return;
    el.dataset.ipReady='1';
    var hdr=el.querySelector('.task-header');
    if(hdr && !hdr.querySelector('.done-btn')){
      var dbtn=document.createElement('button');
      dbtn.className='done-btn';dbtn.title='Erledigt';dbtn.innerHTML='&#10003;';
      dbtn.onclick=function(e){markDone(e,dbtn);};
      var gbtn=document.createElement('button');
      gbtn.className='gear-btn';gbtn.title='Einstellungen';gbtn.innerHTML='&#9881;';
      gbtn.onclick=function(e){openGear(e,gbtn);};
      var toggle=hdr.querySelector('.task-toggle');
      if(toggle){
        hdr.insertBefore(gbtn,toggle);
        hdr.insertBefore(dbtn,toggle);
      }else{
        hdr.appendChild(dbtn);hdr.appendChild(gbtn);
      }
    }
    if(!el.querySelector('.gear-panel')){
      var panel=document.createElement('div');
      panel.className='gear-panel';
      panel.style.display='none';
      panel.onclick=function(e){e.stopPropagation();};
      var planM=el.dataset.planMenge||'0';
      var curTag=el.dataset.tag;
      var dayOpts=DAYS.map(function(d){
        return '<option value="'+d+'"'+(d===curTag?' selected':'')+'>'+d+'</option>';
      }).join('');
      panel.innerHTML=
        '<div class="gear-row"><label>Stueckzahl: <input type="number" class="stk-input" min="0" step="1" value="'+planM+'"></label> <span class="gear-hint">(Plan: '+planM+')</span></div>'+
        '<div class="gear-row"><label>Verschieben auf: <select class="day-select">'+dayOpts+'</select></label></div>'+
        '<div class="gear-row gear-actions"><button class="gear-save" onclick="saveGearChanges(this)">Übernehmen</button>'+
        '<button class="gear-cancel" onclick="cancelGear(this)">Abbrechen</button></div>';
      // Panel oben in den Task-Body — direkt unter dem Header,
      // damit das Rezept nach unten rutscht und nicht uebersprungen werden muss.
      var body=el.querySelector('.task-body');
      if(body){ body.insertBefore(panel, body.firstChild); }
      else{ el.appendChild(panel); }
    }
  });
  loadIst();
}

if(document.readyState==='loading'){
  document.addEventListener('DOMContentLoaded',_initOnce);
}else{
  _initOnce();
}
})();
</script>"""
    done_btn_css = (
        ".done-btn,.gear-btn{background:none;border:1.5px solid #ccc;border-radius:50%;width:28px;height:28px;"
        "font-size:14px;color:#888;cursor:pointer;flex-shrink:0;margin-left:4px;padding:0;line-height:1;"
        "display:flex;align-items:center;justify-content:center;transition:all .15s}"
        ".done-btn:hover,.gear-btn:hover{background:#f5f5f5}"
        ".done-btn:active{background:#e8f5e9;border-color:#43a047;color:#43a047}"
        ".task.done .done-btn{background:#e8f5e9;border-color:#43a047;color:#43a047}"
        ".gear-btn{font-size:15px}"
        ".gear-panel{background:#f8f9fa;border:1px solid #dee2e6;border-radius:8px;padding:12px;margin-top:10px}"
        ".gear-row{margin-bottom:8px;font-size:14px}"
        ".gear-row label{display:flex;align-items:center;gap:8px;font-weight:600}"
        ".gear-row input,.gear-row select{padding:6px 10px;border:1.5px solid #ccc;border-radius:6px;font-size:15px;background:white}"
        ".gear-row input{width:80px;text-align:center;font-weight:700}"
        ".gear-hint{font-size:12px;color:#888;margin-left:6px}"
        ".gear-actions{display:flex;gap:8px;margin-top:10px}"
        ".gear-save,.gear-cancel{flex:1;padding:8px;border-radius:6px;font-weight:700;cursor:pointer;border:none;font-size:14px}"
        ".gear-save{background:#2E75B6;color:white}"
        ".gear-cancel{background:#e9ecef;color:#495057}"
        ".task-header{cursor:pointer}"
        ".task.verschoben{border-left-color:#9b59b6!important;opacity:0.75}"
        ".verschoben-banner{background:#f4ecf7;color:#6c3483;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600;margin-bottom:8px}"
        ".task.ist-null{opacity:0.5}"
        ".task.ist-null .recipe{opacity:0.5}"
        ".ip-status{text-align:center;font-size:12px;padding:5px 12px;border-radius:6px;margin:0 auto 8px;max-width:300px;color:#666;background:#f5f5f5}"
        ".ip-status.ok{background:#d4edda;color:#155724}"
        ".ip-status.warn{background:#fff3cd;color:#856404}"
        ".ip-toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%) translateY(100px);background:#333;color:white;padding:10px 18px;border-radius:8px;font-size:14px;font-weight:600;transition:transform .3s;z-index:1000;box-shadow:0 4px 12px rgba(0,0,0,.2)}"
        ".ip-toast.show{transform:translateX(-50%) translateY(0)}"
        ".ip-toast.success{background:#28a745}"
        ".ip-toast.error{background:#dc3545}"
    )
    html=(
        "<!DOCTYPE html>\n"
        '<html lang="de"><head>'
        '<meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1.0">\n'
        "<title>{tag}</title><style>{css}{done_btn_css}</style></head><body>\n"
        "<h1>{tag}</h1>"
        '<div class="datum">KW {kw} &middot; {datum}</div>\n'
        "{heute_banner}"
        "{zeitbadge}"
        "{inhalt}\n"
        "{einkauf_html}\n"
        '<div class="footer">Kuchenproduktion &middot; KW {kw}</div>\n'
        "{toggle_js}"
        "</body></html>"
    ).format(tag=tag,kw=kw,datum=datum,inhalt=inhalt,css=HTML_CSS,zeitbadge=zeitbadge,heute_banner=heute_banner,einkauf_html=einkauf_html,toggle_js=toggle_js,done_btn_css=done_btn_css)

    fname=OUTPUT_PFAD/"{}.html".format(tag)
    fname.write_text(html,encoding="utf-8")
    print("Tagesplan: {}".format(fname.name))
    return fname

# ============================================================
# ZWISCHENPRODUKT-CHECK (Midweek-Neuplanung)
# ==================

def auto_archiviere_und_lerne():
    """Archiv-Workflow direkt in planner.py — laeuft auch in GitHub Action.

    Aufgaben:
      1. Bei Montag-Lauf: Vorwochen-Archiv suchen (~7 Tage alt) und
         per lerne_woche.py die Vorwoche in Verkaufshistorie.json eintragen.
      2. Aktuellen lagerbestand.json ins Archiv kopieren (= zukuenftiger
         Vorwochen-Stand fuer naechste Woche).

    Muss VOR berechne_wochenbedarf laufen, damit der Forecast den neuen
    Wocheneintrag direkt nutzt. Idempotent: gleiches Datum/KW wird
    ueberschrieben statt dupliziert.
    """
    import shutil, subprocess
    from datetime import timedelta

    lager_json = BASE / "lagerbestand.json"
    archiv_dir = Path(__file__).parent / "lager_archiv"
    archiv_dir.mkdir(exist_ok=True)
    lerne_script = Path(__file__).parent / "lerne_woche.py"

    # --- Schritt 1: Montag-Lernen ---
    if datetime.now().weekday() == 0 and lerne_script.exists():
        # Vorwochen-Archiv suchen (~7 Tage alt, +/- 2 Tage Toleranz)
        ziel_datum = datetime.now() - timedelta(days=7)
        bester = None; bester_abstand = None
        for f in archiv_dir.glob("lagerbestand_*.json"):
            try:
                d = datetime.strptime(f.stem.split("_")[1], "%Y%m%d")
            except (IndexError, ValueError):
                continue
            ab = abs((d - ziel_datum).total_seconds())
            if ab <= 2 * 86400 and (bester_abstand is None or ab < bester_abstand):
                bester = f; bester_abstand = ab
        # Vorwochen-Plan suchen (juengste Wochenuebersicht_*.xlsx ohne heutiges Datum)
        heute_str = datetime.now().strftime("%Y%m%d")
        plaene = sorted(OUTPUT_PFAD.glob("Wochenuebersicht_*.xlsx"))
        plaene = [p for p in plaene if heute_str not in p.name]
        plan = plaene[-1] if plaene else None

        if bester and plan:
            iso = datetime.now().isocalendar()
            kw_label = "{}-{:02d}".format(iso[0], iso[1] - 1 if iso[1] > 1 else 52)
            print("Auto-Lernen Vorwoche {}: altlager={}, plan={}".format(
                kw_label, bester.name, plan.name))
            try:
                result = subprocess.run(
                    [sys.executable, str(lerne_script),
                     "--altlager", str(bester),
                     "--plan",     str(plan),
                     "--kw",       kw_label],
                    cwd=str(BASE), capture_output=True, text=True,
                    encoding="utf-8", timeout=30,
                )
                if result.returncode == 0:
                    print("  Auto-Lernen erfolgreich.")
                else:
                    print("  Auto-Lernen FEHLER (Exit {}): {}".format(
                        result.returncode, result.stderr[:200]))
            except Exception as e:
                print("  Auto-Lernen FEHLER: {}".format(e))
        else:
            print("Auto-Lernen: Montag, aber {} fehlt — uebersprungen".format(
                "Vorwochen-Archiv" if not bester else "Vorwochen-Plan"))

    # --- Schritt 2: Aktuellen Stand archivieren ---
    if lager_json.exists():
        heute_str = datetime.now().strftime("%Y%m%d")
        ziel = archiv_dir / "lagerbestand_{}.json".format(heute_str)
        try:
            shutil.copy2(str(lager_json), str(ziel))
            print("Archiv: {}".format(ziel.name))
        except Exception as e:
            print("Archiv FEHLER: {}".format(e))


import sys  # falls noch nicht importiert


def erstelle_dashboard_html():
    """Aggregiert alle Wochenplan-HTMLs in eine Single-Page Dashboard.html.

    Liest die vorher generierten Dateien (Wochenuebersicht.html, Montag.html, ...,
    Bestellliste.html) ein, extrahiert pro Datei Style + Body und packt sie in
    Tabs. Beim Oeffnen wird der heutige Wochentag automatisch ausgewaehlt
    (Samstag/Sonntag = Wochenuebersicht).

    Datei wird ueberschrieben, ist self-contained und funktioniert offline.
    """
    import re

    TABS = [
        ("ueber",  "&#128202; Woche",   "Wochenuebersicht.html"),
        ("mo",     "Mo",                "Montag.html"),
        ("di",     "Di",                "Dienstag.html"),
        ("mi",     "Mi",                "Mittwoch.html"),
        ("do",     "Do",                "Donnerstag.html"),
        ("fr",     "Fr",                "Freitag.html"),
        ("bestell","&#128722; Einkauf", "Bestellliste.html"),
    ]
    WOCHENTAG_ZU_TAB = {0:"mo", 1:"di", 2:"mi", 3:"do", 4:"fr"}  # Sa/So -> ueber

    panels    = []
    styles    = []
    seen_css  = set()

    for tab_id, label, fname in TABS:
        f = OUTPUT_PFAD / fname
        if not f.exists():
            panels.append((tab_id, label, '<div style="padding:20px;color:#888">Datei {} nicht gefunden</div>'.format(fname)))
            continue
        try:
            html = f.read_text(encoding='utf-8', errors='replace').rstrip('\x00 \t\r\n')
        except Exception as e:
            panels.append((tab_id, label, '<div style="padding:20px;color:#c00">Fehler: {}</div>'.format(e)))
            continue
        # Style extrahieren (alle <style> Bloecke)
        for m in re.finditer(r'<style[^>]*>(.*?)</style>', html, re.S):
            css = m.group(1).strip()
            if css and css not in seen_css:
                seen_css.add(css)
                styles.append(css)
        # Body extrahieren
        mb = re.search(r'<body[^>]*>(.*?)</body>', html, re.S)
        body = mb.group(1).strip() if mb else html
        panels.append((tab_id, label, body))

    aktueller_tab = WOCHENTAG_ZU_TAB.get(datetime.now().weekday(), "ueber")

    # Tabs-CSS und Skelett
    dashboard_css = """
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f0f2f5;padding-bottom:60px}
.tab-bar{position:sticky;top:0;z-index:100;background:#2E75B6;display:flex;overflow-x:auto;
        scrollbar-width:none;-webkit-overflow-scrolling:touch;box-shadow:0 2px 6px rgba(0,0,0,.15)}
.tab-bar::-webkit-scrollbar{display:none}
.tab-btn{flex:0 0 auto;padding:14px 16px;background:none;border:none;color:#cbe0f0;
         font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap;
         border-bottom:3px solid transparent;transition:all .15s}
.tab-btn.active{color:#fff;border-bottom-color:#fff;background:rgba(255,255,255,.08)}
.tab-btn:hover{color:#fff}
.panel{display:none}
.panel.active{display:block}
.panel > .header{display:none}  /* eingebettete Wochenuebersicht-Header ausblenden */
.panel > h1{display:none}        /* eingebettete Tagesplan-Titel ausblenden */
.panel > .datum{display:none}
.dash-meta{padding:6px 14px;font-size:11px;color:#888;text-align:right}
"""

    html_out = []
    html_out.append('<!DOCTYPE html><html lang="de"><head>')
    html_out.append('<meta charset="UTF-8">')
    html_out.append('<meta name="viewport" content="width=device-width,initial-scale=1.0">')
    html_out.append('<title>Kuchenproduktion Dashboard</title>')
    html_out.append('<style>')
    html_out.append(dashboard_css)
    for css in styles:
        html_out.append(css)
    html_out.append('</style></head><body>')

    # Tab-Leiste
    html_out.append('<div class="tab-bar">')
    for tab_id, label, _ in panels:
        cls = "tab-btn active" if tab_id == aktueller_tab else "tab-btn"
        html_out.append('<button class="{}" data-tab="{}" onclick="showTab(\'{}\')">{}</button>'.format(cls, tab_id, tab_id, label))
    html_out.append('</div>')
    html_out.append('<div class="dash-meta">'
                    'Stand: {} &nbsp;&middot;&nbsp; '
                    '<a href="../eingabe.html" '
                    'style="color:#2E75B6;text-decoration:none;font-weight:600">'
                    '&#9998; Lager eintragen</a>'
                    '</div>'.format(datetime.now().strftime("%d.%m.%Y %H:%M")))

    # Panels
    for tab_id, label, body in panels:
        cls = "panel active" if tab_id == aktueller_tab else "panel"
        html_out.append('<div class="{}" id="panel-{}">{}</div>'.format(cls, tab_id, body))

    # JavaScript: Tab-Wechsel + LocalStorage fuer letzten Tab
    html_out.append('<script>')
    html_out.append('function showTab(id){')
    html_out.append('  document.querySelectorAll(".tab-btn").forEach(b=>b.classList.toggle("active",b.dataset.tab===id));')
    html_out.append('  document.querySelectorAll(".panel").forEach(p=>p.classList.toggle("active",p.id==="panel-"+id));')
    html_out.append('  try{localStorage.setItem("dashboard_tab",id);}catch(e){}')
    html_out.append('  window.scrollTo(0,0);')
    html_out.append('}')
    # Beim Laden: wenn LocalStorage einen Tab hat und der noch existiert, dort hin
    html_out.append('document.addEventListener("DOMContentLoaded",function(){')
    html_out.append('  try{')
    html_out.append('    var saved=localStorage.getItem("dashboard_tab");')
    html_out.append('    if(saved && document.getElementById("panel-"+saved)){showTab(saved);}')
    html_out.append('  }catch(e){}')
    html_out.append('});')
    html_out.append('</script>')
    html_out.append('</body></html>')

    fname = OUTPUT_PFAD / "Dashboard.html"
    fname.write_text("\n".join(html_out), encoding='utf-8')
    print("Dashboard HTML: {}".format(fname.name))
    return fname


def frage_vorbereitungen(bedarf, wplan, vtage, sub_rezepte):
    import sys
    if vtage == list(ARBEITSTAGE):
        return wplan
    vergangene = [t for t in ARBEITSTAGE if t not in vtage]
    if not vergangene:
        return wplan
    def noch_geplant(prod_keys):
        return any(
            a.get("produkt_key") == pk and a.get("menge", 0) > 0
            for pk in prod_keys for tag in vtage for a in wplan.get(tag, [])
        )
    VORBEREITUNGEN = [
        {"id":"murbeteig","check_tag":"Dienstag",
         "check_prods":["Kaesekuchen","Kaese-Rhabarber Schnitte"],
         "frage":"Murbeteigboeden bereits fertig?","aufgabe":"MURBETEIG nachbacken",
         "notiz":"Boeden fuer KK + KR","prio":1,"aktiv_min":35,"passiv_min":10},
        {"id":"streusel","check_tag":"Montag",
         "check_prods":["Kaese-Rhabarber Schnitte"],
         "frage":"Streusel (Mo) bereits hergestellt?","aufgabe":"STREUSEL nachstellen",
         "notiz":"Fuer verbleibende KR-Schnitte","prio":2,"aktiv_min":15,"passiv_min":0},
        {"id":"pistazien_prep","check_tag":"Mittwoch",
         "check_prods":["Pistazien Toertchen"],
         "frage":"Pistazien-Vorbereitung bereits erledigt?","aufgabe":"PISTAZIEN VORBEREITUNG nacharbeiten",
         "notiz":"Brownieboden + Ganache","prio":3,"aktiv_min":45,"passiv_min":60},
    ]
    relevante = [v for v in VORBEREITUNGEN
                 if v["check_tag"] in vergangene and noch_geplant(v["check_prods"])]
    if not relevante or not sys.stdin.isatty():
        return wplan
    for v in relevante:
        try:
            antwort = input("\n{} [j/N]: ".format(v["frage"])).strip().lower()
        except EOFError:
            return wplan
        if antwort != "j":
            erster = vtage[0]
            wplan.setdefault(erster, []).insert(0, {
                "produkt":v["aufgabe"],"produkt_key":None,"notiz":v["notiz"],
                "menge":0,"prio":v.get("prio",5),"aktiv_min":v.get("aktiv_min",0),
                "passiv_min":v.get("passiv_min",0),"excel_only":False,
            })
    return wplan


def main():
    print("\nKuchenproduktion Wochenplaner")
    print("========================================")
    print("Lade Konfiguration..."); lade_konfiguration(); _init_murbeteig()
    print("Lade Rezepte..."); rezepte = lade_alle_rezepte()
    print("  {} Rezepte geladen".format(len(rezepte)))
    sub_rezepte = lade_sub_rezepte()
    print("  {} Sub-Rezepte geladen".format(len(sub_rezepte)))
    print("Lade Preise..."); preise = lade_preise()
    auto_archiviere_und_lerne()
    print("Lese Verkaufszahlen..."); verkauf = lese_verkaufszahlen()
    print("Lese Lagerbestand..."); lager = lese_lagerbestand()
    plan_ab_morgen = bool(lager.get('_plan_ab_morgen', False))
    print("Berechne Wochenbedarf...")
    cafe2_erledigt = cafe2_bereits_abgeholt()
    bedarf = berechne_wochenbedarf(verkauf, lager, cafe2_erledigt)
    print("Berechne Muerbeteig-Bedarf...")
    murt = berechne_murbeteig(bedarf, lager)
    print("Erstelle Wochenplan...")
    verbleibende_tage = ab_heute_tage(ab_morgen=plan_ab_morgen)
    vtage, wplan = erstelle_wochenplan(bedarf, murt, rezepte, sub_rezepte,
                                       verbleibende_tage=verbleibende_tage)
    wplan = frage_vorbereitungen(bedarf, wplan, vtage, sub_rezepte)
    # Wplan-Snapshot fuer diese KW persistieren — wird in der Wochenmatrix
    # gelesen, damit vergangene Tage bei mid-week-Laeufen sichtbar bleiben.
    speichere_wplan_kw(wplan, vtage=vtage)
    print("Schreibe Excel..."); erstelle_excel(bedarf, murt, wplan)
    print("Erstelle HTML Bestellliste...")
    erstelle_html_bestellliste(bedarf, wplan, vtage, verkauf, lager, rezepte, sub_rezepte, preise)
    print("Berechne Lagerprognose...")
    rolling, tages_info, engpaesse = berechne_rolling_inventory(bedarf, lager, wplan)
    if engpaesse:
        print("  WARNUNG Engpaesse: {}".format(engpaesse))
    print("Erstelle HTML Wochenuebersicht...")
    erstelle_html_wochenuebersicht(bedarf, murt, wplan, rolling, tages_info, vtage=vtage)
    print("Erstelle HTML Tagesplaene...")
    for tag, aufgaben in wplan.items():
        # Vergangene Tage NICHT ueberschreiben: Wenn der Planer mitten in der
        # Woche laeuft (z.B. Di Lager aktualisiert), bleibt Montag.html mit
        # allen bereits gesetzten Haken und per Zahnrad geaenderten Stueckzahlen
        # unangetastet. Nachtraegliches Korrigieren via Zahnrad bleibt moeglich,
        # weil die ist_produktion_KW*.json unabhaengig weiterlaeuft.
        if vtage and tag not in vtage:
            print("  {} uebersprungen (vergangener Tag) — bestehende HTML bleibt.".format(tag))
            continue
        ist_heut = (tag == vtage[0]) if vtage else False
        lager_info = tages_info.get(tag, [])
        erstelle_html(tag, aufgaben, murt, rezepte, sub_rezepte=sub_rezepte,
                      lager_info=lager_info, ist_heut=ist_heut, preise=preise)
    print("Aggregiere Single-Page Dashboard...")
    erstelle_dashboard_html()
    print("\nFertig! Ordner: {}".format(OUTPUT_PFAD))


if __name__ == "__main__":
    main()
