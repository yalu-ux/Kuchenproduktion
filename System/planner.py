#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Kuchenproduktion Wochenplaner"""
import struct, zlib, xml.etree.ElementTree as ET
import openpyxl, math, re
from pathlib import Path
from datetime import datetime
from collections import defaultdict

BASE                 = Path(__file__).parent.parent
OUTPUT_PFAD          = BASE / "Wochenplan"
LAGERBESTAND_DATEI   = BASE / "Lagerbestand.xlsx"
VERKAUFSZAHLEN_DATEI = BASE / "Verkaufszahlen1.xlsx"
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

def ab_heute_tage():
    """Gibt die noch verbleibenden Arbeitstage zurueck (ab heute inkl.).
    Samstag/Sonntag → komplette Woche (Planung fuer naechste Woche).
    """
    wochentag_idx = datetime.now().weekday()  # 0=Mo, 1=Di, ..., 6=So
    if wochentag_idx >= 5:   # Wochenende → volle Woche planen
        return list(ARBEITSTAGE)
    return ARBEITSTAGE[wochentag_idx:]

FROSTER_SCHWELLEN = {
    "Beeren Tartelette":               45,
    "Mango-Passionsfrucht Tartelette": 45,
    "Pistazien Toertchen":             30,
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

VERKAUF_ZU_CONFIG = {
    "Beeren Tartelette":"Beeren Tartelette","Bienenstich":"Bienenstich",
    "Donauwelle":"Donauwelle","Kaese-Rhabarber Schnitte":"Kaese-Rhabarber Schnitte",
    "Kaesekuchen":"Kaesekuchen","Lauch-Speck Quiche":"Lauch-Speck Quiche",
    "Mango-Passionsfrucht Tartelette":"Mango-Passionsfrucht Tartelette",
    "Pistazien Toertchen":"Pistazien Toertchen",
    "Käse-Rhabarber Schnitte":"Kaese-Rhabarber Schnitte",
    "Käsekuchen":"Kaesekuchen","Pistazien Törtchen":"Pistazien Toertchen",
}

MURBETEIG_PORTIONEN = {"Tartelette":40,"30x20 Schnitte":200,"Kuchenform":350,"Murbeteig rund":180}
MURBETEIG_REZEPT    = {"Zucker":250,"Butter":500,"Salz":1,"Zitronenaroma":1,"Vanillearoma":2,"Vollei":100,"Weizenmehl":750}
MURBETEIG_GESAMT    = sum(MURBETEIG_REZEPT.values())

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
SUB_REZEPT_DATEIEN = {
    "Salzmasse":                "Salzmasse.xlsx",
    "Dinkelhefeteig":           "Dinkelhefeteig.xlsx",
    "Sandmasse hell":           "Sandmasse hell.xlsx",
    "Sandmasse dunkel":         "Sandmasse dunkel.xlsx",
    "Deutsche Buttercreme":     "Deutsche Buttercreme.xlsx",
    "Überzugsganache":          "Überzugsganache.xlsx",
    "Amerikanische Käsemasse":  "Amerikanische Käsemasse.xlsx",
    "Streusel":                 "Streusel.xlsx",
    "Mandel Sandmasse":         "Mandel Sandmasse.xlsx",
    "Vanillecreme":             "Vanillecreme.xlsx",
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
    return {'name':name,'max_charge_num':max_charge_num,'max_charge_text':max_charge_text,
            'sections':sections,'aktiv_min':aktiv_min,'passiv_min':passiv_min}

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
    """Parst eine einfache Sub-Rezept-Datei (Zutatenliste + Yield)."""
    zeilen = _lade_rezept_zeilen(pfad)
    name_val = Path(pfad).stem
    for b, c in zeilen:
        if b: name_val = b; break

    skip = ['maximal', 'herstellung', 'haltbarkeit', 'hinweis', 'zeitaufwand']
    zutaten = []; yield_gramm = 0.0; in_block = False
    aktiv_min = passiv_min = 0; yield_info = None

    for b, c in zeilen:
        if not b: continue
        bl = b.lower()
        # Zeitaufwand
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
        # Abbruch bei Skip-Schlüsselwörtern
        if any(w in bl for w in skip): break
        # Zutaten-Block
        if bl in ('zutat', 'rezept:', 'rezept', 'rezept :'): in_block = True; continue
        if c:
            try:
                menge = float(c.replace(',', '.'))
                zutaten.append({'zutat': b, 'menge': menge, 'einheit': 'g'})
                yield_gramm += menge
                in_block = True
            except: pass

    return {'name': str(name_val), 'zutaten': zutaten, 'yield_gramm': yield_gramm,
            'yield_info': yield_info, 'aktiv_min': aktiv_min, 'passiv_min': passiv_min}

def lade_sub_rezepte():
    """Laedt alle Sub-Rezept-Dateien."""
    sub = {}
    if not REZEPTE_PFAD.exists(): return sub
    for key, datei in SUB_REZEPT_DATEIEN.items():
        pfad = REZEPTE_PFAD / datei
        if not pfad.exists(): continue
        r = parse_sub_rezept(pfad)
        if r: sub[key] = r
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
    """Erstellt HTML-Block fuer ein einzelnes Sub-Rezept."""
    if n_mal <= 1:
        title = 'Sub-Rezept: {} &mdash; 1x Grundrezept ({:.0f}g)'.format(srz['name'], srz['yield_gramm'])
    else:
        title = 'Sub-Rezept: {} &mdash; {}x Grundrezept ({:.0f}g gesamt)'.format(
            srz['name'], n_mal, srz['yield_gramm'] * n_mal)
    zhtml = ''.join(
        '<li><span class="zutat">{}</span><span class="menge">{:.0f}g</span></li>'.format(
            z['zutat'], z['menge'] * n_mal)
        for z in srz['zutaten']
    )
    return (
        '<div class="sub-recipe">'
        '<div class="sub-recipe-title">{}</div>'
        '<ul class="zutaten">{}</ul>'
        '</div>'
    ).format(title, zhtml)

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
        print("Konfiguration: Ofen={}, Schwellen={}".format(
            {k:v['max_ofen'] for k,v in PRODUKT_CONFIG.items()}, FROSTER_SCHWELLEN))
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
    return verkauf

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
            netto=max(0,gesamt-stock); mc=cfg["min_charge"]
            prod=math.ceil(netto/mc)*mc if netto>0 else 0; grund=""
        cafe2_lief=0 if cafe2_erledigt else (min(math.ceil(c2),int(stock)) if cfg.get("cafe2") else 0)
        bedarf[key]={"cafe1_woche":round(c1,1),"cafe2_woche":round(c2,1),"cafe2_erledigt":cafe2_erledigt,
                     "gesamt_woche":round(gesamt,1),"im_lager":stock,
                     "netto_bedarf":round(max(0,gesamt-stock),1),
                     "zu_produzieren":prod,"cafe2_lieferung":cafe2_lief,"grund":grund}
    return bedarf

def berechne_murbeteig(bedarf):
    portionen={}; gesamt_g=0
    for key,b in bedarf.items():
        menge=b["zu_produzieren"]
        if not menge: continue
        murt=PRODUKT_CONFIG[key].get("murbeteig_typ")
        if not murt: continue
        g=menge*MURBETEIG_PORTIONEN.get(murt,0)
        if murt not in portionen: portionen[murt]={"anzahl":0,"gramm":0}
        portionen[murt]["anzahl"]+=menge; portionen[murt]["gramm"]+=g; gesamt_g+=g
    gr=math.ceil(gesamt_g/MURBETEIG_GESAMT) if MURBETEIG_GESAMT else 0
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
            tv = bd.get("gesamt_woche", 0.0) / 7.0          # Tagesverbrauch
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
    tv         = bd.get("gesamt_woche", 0.0) / 7.0  # Tagesverbrauch
    start_lager = float(bd.get("im_lager", 0))

    if n_total <= 0 or not verfuegbare_tage:
        return []
    if tv <= 0:
        # Kein laufender Verbrauch: einfach gleichmaessig verteilen
        return verteile_produktionstage(math.ceil(n_total / per_charge), verfuegbare_tage)

    n_chargen = math.ceil(n_total / per_charge)
    tage = list(verfuegbare_tage)
    n = len(tage)

    # Verbrauch an Tagen VOR dem ersten verfuegbaren Tag beruecksichtigen.
    # z.B. Donauwelle startet erst Dienstag → Montag-Verbrauch wurde schon abgezogen.
    n_pre = sum(1 for t in ARBEITSTAGE if t != tage[0] and ARBEITSTAGE.index(t) < ARBEITSTAGE.index(tage[0]))
    start_lager = start_lager - tv * n_pre

    placed = {}     # {tag_index: menge_produziert}
    result  = []

    for _ in range(n_chargen):
        # --- Rollierenden Bestand mit bisher platzierten Chargen simulieren ---
        s = start_lager
        stock_end = []
        for i, t in enumerate(tage):
            s += placed.get(i, 0)
            s -= tv
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

        # --- Zeitoptimal: geringste aktive Arbeitszeit unter Kandidaten ---
        best_i = min(kandidaten, key=lambda i: aktiv_plan.get(tage[i], 0))
        placed[best_i] = per_charge
        result.append(tage[best_i])

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
    vtage = list(verbleibende_tage) if verbleibende_tage else list(ARBEITSTAGE)
    def _filter(tage_liste):
        """Filtert eine Tagesliste auf die verbleibenden Tage."""
        return [t for t in tage_liste if t in vtage]

    def add(tag,produkt,menge,notiz="",prio=10,excel_only=False,produkt_key=None,aktiv_min=0,passiv_min=0):
        plan[tag].append({"produkt":produkt,"menge":menge,"notiz":notiz,
                          "prio":prio,"excel_only":excel_only,
                          "produkt_key":produkt_key or produkt,
                          "aktiv_min":aktiv_min,"passiv_min":passiv_min})
        if not excel_only and aktiv_min: aktiv_plan[tag] += aktiv_min
    def info(tag,produkt,notiz="",prio=10,excel_only=False,aktiv_min=0,passiv_min=0):
        plan[tag].append({"produkt":produkt,"menge":0,"notiz":notiz,
                          "prio":prio,"excel_only":excel_only,"produkt_key":None,
                          "aktiv_min":aktiv_min,"passiv_min":passiv_min})
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
    n_tage_kk = math.ceil(rem_kk / kk_per_day) if rem_kk > 0 else 0
    tage_kk = verteile_produktionstage(n_tage_kk, _filter(["Dienstag","Mittwoch","Donnerstag","Freitag"]))
    for tag in tage_kk:
        if rem_kk<=0: break
        heute=0
        while rem_kk>0 and heute<kk_per_day:
            ch=min(kk_max,rem_kk)
            add(tag,"Kaesekuchen",ch,"Charge {}/{}".format(cn,kk_ges),
                prio=PRIO["KK"],produkt_key="Kaesekuchen",aktiv_min=kk_a,passiv_min=kk_p)
            rem_kk-=ch; cn+=1; heute+=ch
    murt["kk_montag_reserve"]=min(kk_per_day,kk_menge)

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
        streusel_aktiv = streusel_srz.get("aktiv_min", 10) * streusel_gr
        if "Montag" in vtage: info("Montag","STREUSEL HERSTELLEN",
             "{}x Grundrezept ({:.0f}g) - fuer die ganze Woche".format(streusel_gr, streusel_gesamt),
             prio=PRIO["STREUSEL"], aktiv_min=streusel_aktiv)

    # --- BIENENSTICH ---
    b_menge=bedarf.get("Bienenstich",{}).get("zu_produzieren",0)
    b_max=PRODUKT_CONFIG["Bienenstich"]["max_ofen"]
    b_ges=math.ceil(b_menge/b_max) if b_menge else 0
    rem,c=b_menge,1
    b_a, b_p = berechne_zeitaufwand("Bienenstich", rezepte, sub_rezepte, b_max)
    tage_b = verteile_bestandssicher("Bienenstich", bedarf, aktiv_plan,
                                     _filter(["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]), b_max)
    for tag in tage_b:
        if rem<=0: break
        ch=min(rem,b_max)
        add(tag,"Bienenstich",ch,"Charge {}/{}".format(c,b_ges),prio=PRIO["BIEN"],
            produkt_key="Bienenstich",aktiv_min=b_a,passiv_min=b_p)
        rem-=ch; c+=1

    # --- QUICHE ---
    q_menge=bedarf.get("Lauch-Speck Quiche",{}).get("zu_produzieren",0)
    q_max=PRODUKT_CONFIG["Lauch-Speck Quiche"]["max_ofen"]
    q_ges=math.ceil(q_menge/q_max) if q_menge else 0
    rem,c=q_menge,1
    q_a, q_p = berechne_zeitaufwand("Lauch-Speck Quiche", rezepte, sub_rezepte, q_max)
    tage_q = verteile_produktionstage(q_ges, _filter(["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]))
    for tag in tage_q:
        if rem<=0: break
        ch=min(rem,q_max)
        add(tag,"Lauch-Speck Quiche",ch,"Charge {}/{}".format(c,q_ges),prio=PRIO["QUICHE"],
            produkt_key="Lauch-Speck Quiche",aktiv_min=q_a,passiv_min=q_p)
        rem-=ch; c+=1

    # --- DONAUWELLE ---
    d_menge=bedarf.get("Donauwelle",{}).get("zu_produzieren",0)
    d_max=PRODUKT_CONFIG["Donauwelle"]["max_ofen"]
    d_ges=math.ceil(d_menge/d_max) if d_menge else 0
    rem,c=d_menge,1
    d_a, d_p = berechne_zeitaufwand("Donauwelle", rezepte, sub_rezepte, d_max)
    tage_d = verteile_bestandssicher("Donauwelle", bedarf, aktiv_plan,
                                     _filter(["Dienstag","Mittwoch","Donnerstag","Freitag"]), d_max)
    for tag in tage_d:
        if rem<=0: break
        ch=min(rem,d_max)
        add(tag,"Donauwelle",ch,"Charge {}/{}".format(c,d_ges),prio=PRIO["DONA"],
            produkt_key="Donauwelle",aktiv_min=d_a,passiv_min=d_p)
        rem-=ch; c+=1

    # --- KAESE-RHABARBER ---
    kr_menge=bedarf.get("Kaese-Rhabarber Schnitte",{}).get("zu_produzieren",0)
    kr_max=PRODUKT_CONFIG["Kaese-Rhabarber Schnitte"]["max_ofen"]
    kr_ges=math.ceil(kr_menge/kr_max) if kr_menge else 0
    rem,c=kr_menge,1
    kr_a, kr_p = berechne_zeitaufwand("Kaese-Rhabarber Schnitte", rezepte, sub_rezepte, kr_max)
    tage_kr = verteile_produktionstage(kr_ges, _filter(["Dienstag","Mittwoch","Donnerstag","Freitag"]))
    for tag in tage_kr:
        if rem<=0: break
        ch=min(rem,kr_max)
        add(tag,"Kaese-Rhabarber Schnitte",ch,"Charge {}/{}".format(c,kr_ges),prio=PRIO["KR"],
            produkt_key="Kaese-Rhabarber Schnitte",aktiv_min=kr_a,passiv_min=kr_p)
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

    return plan

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
    "box-shadow:0 1px 4px rgba(0,0,0,.08);border-left:4px solid #2E75B6}"
    ".task.special{border-left-color:#375623}"
    ".task.prep{border-left-color:#9B59B6;background:#F9F0FF}"
    ".task-header{font-weight:700;font-size:16px;margin-bottom:4px}"
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
)

def erstelle_html(tag, aufgaben, murt, rezepte, sub_rezepte=None, lager_info=None, ist_heut=False):
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

    # Sammle Pro-Tag-Mengen pro Produkt (fuer Rezept-Skalierung)
    tagesmengen=defaultdict(int)
    for a in aufgaben_html:
        if a.get("produkt_key") and a["menge"]>0 and not a.get("excel_only"):
            tagesmengen[a["produkt_key"]]+=int(a["menge"])

    # Bereits gezeigte Rezepte (je Produkt nur einmal pro Tag)
    rezept_gezeigt=set()

    for a in aufgaben_html:
        if a.get("excel_only"): continue
        p=a["produkt"]; n=a["notiz"]; m=a["menge"]; pk=a.get("produkt_key")

        # --- Muerbeteig ansetzen ---
        a_min = a.get("aktiv_min", 0)
        p_min = a.get("passiv_min", 0)

        if "MUERBETEIG ANSETZEN" in p:
            gr=murt["grundrezepte"]
            zutaten_html="".join(
                '<li><span class="zutat">{}</span><span class="menge">{}g</span></li>'.format(z,mg)
                for z,mg in MURBETEIG_REZEPT.items()
            )
            inhalt+=(
                '<div class="task special">'
                '<div class="task-header">Muerbeteig ansetzen &mdash; {}x Grundrezept</div>'
                '<div class="task-note">{}</div>'
                '{}'
                '<div class="recipe">'
                '<div class="recipe-title">Grundrezept (1x ansetzen, {}x wiederholen)</div>'
                '<ul class="zutaten">{}</ul>'
                '</div></div>'
            ).format(gr,n,_zeit_badge(a_min,p_min),gr,zutaten_html)
            continue

        # --- Muerbeteig backen ---
        if "MUERBETEIG BACKEN" in p:
            portionen_html="".join(
                '<li><span class="zutat">{}x {}</span><span class="menge">{}g</span></li>'.format(
                    v["anzahl"],t,int(v["gramm"]))
                for t,v in murt["portionen"].items()
            )
            res=murt.get("kk_montag_reserve",0)
            if res:
                portionen_html+='<li><span class="zutat">{}x Murbeteig rund (fuer naechsten Montag)</span><span class="menge">{}g</span></li>'.format(
                    res,res*MURBETEIG_PORTIONEN["Murbeteig rund"])
            inhalt+=(
                '<div class="task special">'
                '<div class="task-header">Muerbeteig backen &mdash; alle Boeden heute</div>'
                '<div class="task-note">180&deg;C, 7&ndash;9 Minuten</div>'
                '{}'
                '<div class="recipe"><div class="recipe-title">Boeden aufteilen</div>'
                '<ul class="zutaten">{}</ul></div>'
                '</div>'
            ).format(_zeit_badge(a_min,p_min),portionen_html)
            continue

        # --- Pistazien ---
        if "PISTAZIEN" in p.upper():
            cls="prep"
            inhalt+='<div class="task {}">' \
                    '<div class="task-header">{}{}</div>' \
                    '{}<div class="task-note">{}</div></div>'.format(
                        cls, p, " &mdash; {} Stk".format(int(m)) if m else "",
                        _zeit_badge(a_min,p_min), n)
            continue

        # --- Streusel herstellen ---
        if "STREUSEL HERSTELLEN" in p:
            sr_key = "Streusel"
            srz = (sub_rezepte or {}).get(sr_key)
            n_mal_str = n.split("x ")[0] if "x " in n else "1"
            try: n_mal_s = int(n_mal_str.strip())
            except: n_mal_s = 1
            streusel_recipe_html = ''
            if srz:
                streusel_recipe_html = format_sub_rezept_html(sr_key, srz, n_mal_s)
            inhalt += (
                '<div class="task special">'
                '<div class="task-header">Streusel herstellen &mdash; {}x Grundrezept</div>'
                '<div class="task-note">{}</div>'
                '{}'
                '{}'
                '</div>'
            ).format(n_mal_s, n, _zeit_badge(a_min,p_min), streusel_recipe_html)
            continue

        # --- Normale Produktionsaufgabe ---
        tagesm=tagesmengen.get(pk,m) if pk else m
        # Notiz ohne "Murbeteigboden von letzter Woche" sauber anzeigen
        notiz_extra=""
        if "letzter Woche" in n:
            notiz_extra = ("Murbeteigboeden von dieser Woche verwenden (Dienstag gebacken)"
                           if ist_heut else "Murbeteigboeden von letzter Woche verwenden")

        header="{}".format(p.replace("Kaesekuchen","Kaesekuchen").replace("Kaese-Rhabarber Schnitte","Kaese-Rhabarber"))
        charge_info=n.split(" - ")[0] if " - " in n else n

        # Zeitangabe fuer diese Aufgabe
        a_min = a.get("aktiv_min", 0)
        p_min = a.get("passiv_min", 0)
        zeit_html = ""
        if a_min:
            z_parts = ["&#9998; {}".format(fmt_min(a_min))]
            if p_min: z_parts.append("&#9201; {} passiv".format(fmt_min(p_min)))
            zeit_html = '<div class="task-zeit">{}</div>'.format(" &nbsp;|&nbsp; ".join(z_parts))

        if m and m > 0:
            inhalt+='<div class="task"><div class="task-header">{} &mdash; {} Stk</div>'.format(p,int(m))
        else:
            inhalt+='<div class="task"><div class="task-header">{}</div>'.format(p)
        if charge_info: inhalt+='<div class="task-note">{}</div>'.format(charge_info)
        if notiz_extra: inhalt+='<div class="task-note" style="color:#8B4513">{}</div>'.format(notiz_extra)
        if zeit_html: inhalt+=zeit_html

        # Rezept (nur beim ersten Vorkommen des Produkts an diesem Tag)
        if pk and pk not in rezept_gezeigt:
            rzp=rezepte.get(pk)
            if rzp:
                rezept_html=format_rezept_block_html(rzp,tagesm)
                inhalt+=rezept_html
                # Sub-Rezepte zeigen
                if sub_rezepte:
                    # Streusel wird Montags fuer die ganze Woche hergestellt
                    skip = {"Streusel"} if tag != "Montag" else set()
                    inhalt+=format_alle_sub_rezepte_html(rzp,tagesm,sub_rezepte,skip_keys=skip)
            rezept_gezeigt.add(pk)

        inhalt+='</div>'

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
    html=(
        "<!DOCTYPE html>\n"
        '<html lang="de"><head>'
        '<meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1.0">\n'
        "<title>{tag}</title><style>{css}</style></head><body>\n"
        "<h1>{tag}</h1>"
        '<div class="datum">KW {kw} &middot; {datum}</div>\n'
        "{heute_banner}"
        "{zeitbadge}"
        "{inhalt}\n"
        '<div class="footer">Kuchenproduktion &middot; KW {kw}</div>\n'
        "</body></html>"
    ).format(tag=tag,kw=kw,datum=datum,inhalt=inhalt,css=HTML_CSS,zeitbadge=zeitbadge,heute_banner=heute_banner)

    fname=OUTPUT_PFAD/"{}.html".format(tag)
    fname.write_text(html,encoding="utf-8")
    print("Tagesplan: {}".format(fname.name))
    return fname

# ============================================================
# ZWISCHENPRODUKT-CHECK (Midweek-Neuplanung)
# ============================================================
def frage_vorbereitungen(bedarf, wplan, vtage, sub_rezepte):
    """Fragt bei Midweek-Neustart nach dem Status vergangener Vorbereitungen.
    Wird nur aufgerufen wenn vtage != volle Woche UND stdin ein Terminal ist.
    Gibt (evtl. ergaenztes) wplan zurueck.
    """
    import sys
    if vtage == list(ARBEITSTAGE):
        return wplan                    # volle Woche — nichts zu pruefen

    vergangene = [t for t in ARBEITSTAGE if t not in vtage]
    erster_tag = vtage[0]

    def noch_geplant(prod_keys):
        """True wenn mindestens ein Produkt aus prod_keys noch in vtage eingeplant ist."""
        return any(
            a.get("produkt_key") == pk and a.get("menge", 0) > 0
            for pk in prod_keys
            for tag in vtage
            for a in wplan.get(tag, [])
        )

    # ── Alle moeglichen Vorbereitungen ─────────────────────────
    # Jeder Eintrag: {
    #   "id"          : eindeutiger Schluessel,
    #   "frage"       : Frage an den Nutzer,
    #   "check_tag"   : Dieser Tag muss in vergangene liegen,
    #   "check_prods" : Produkt-Keys die noch geplant sein muessen,
    #   "aufgabe"     : Aufgabenname wenn nicht erledigt,
    #   "notiz"       : Notiz zur Aufgabe,
    #   "prio"        : Sortierprioритaet im Tagesplan,
    #   "aktiv_min"   : Geschaetzte aktive Minuten,
    #   "passiv_min"  : Geschaetzte passive Minuten,
    # }
    VORBEREITUNGEN = [
        {
            "id":          "murbeteig",
            "check_tag":   "Dienstag",   # backen war Di (ansetzen Mo irrelevant wenn Di ok)
            "check_prods": ["Kaesekuchen", "Kaese-Rhabarber Schnitte"],
            "frage":       "Murbeteigboeden (Montag angesetzt, Dienstag gebacken) — bereits fertig?",
            "aufgabe":     "MURBETEIG nachbacken",
            "notiz":       "Boeden fuer Kaesekuchen + KR-Schnitte — Kuehlschrank-Teig oder frisch ansetzen "
                           "(mind. 30 Min Ruhezeit), dann backen 180 Grad 7-9 Min",
            "prio":        1,
            "aktiv_min":   35,
            "passiv_min":  10,
        },
        {
            "id":          "streusel",
            "check_tag":   "Montag",
            "check_prods": ["Kaese-Rhabarber Schnitte"],
            "frage":       "Streusel (Montag fuer die ganze Woche) — bereits hergestellt?",
            "aufgabe":     "STREUSEL nachstellen",
            "notiz":       "Fuer verbleibende Kaese-Rhabarber Schnitte",
            "prio":        2,
            "aktiv_min":   15,
            "passiv_min":  0,
        },
        {
            "id":          "pistazien_prep",
            "check_tag":   "Mittwoch",
            "check_prods": ["Pistazien Toertchen"],
            "frage":       "Pistazien-Vorbereitung (Brownieboden + Ganache ansetzen) — bereits erledigt?",
            "aufgabe":     "PISTAZIEN VORBEREITUNG nacharbeiten",
            "notiz":       "Brownieboden backen + auskuehlen, Pistazienganache ansetzen (Uebernacht)",
            "prio":        3,
            "aktiv_min":   45,
            "passiv_min":  60,
        },
        {
            "id":          "vanillecreme_beeren",
            "check_tag":   "Dienstag",   # Vanillecreme fuer Beeren Tartelette am Vorabend (Di abends)
            "check_prods": ["Beeren Tartelette"],
            "frage":       "Vanillecreme fuer Beeren Tartelette (Vorabend angesetzt, Uebernacht kuehl) — fertig?",
            "aufgabe":     "VANILLECREME nachstellen — Beeren Tartelette",
            "notiz":       "Vanillecreme ansetzen, mind. 4h kuehl stellen (besser Uebernacht). "
                           "Tartelette-Produktion ggf. auf naechsten Tag verschieben.",
            "prio":        2,
            "aktiv_min":   20,
            "passiv_min":  240,
        },
    ]

    # Filtere: nur relevante Vorbereitungen fragen
    relevante = [
        v for v in VORBEREITUNGEN
        if v["check_tag"] in vergangene and noch_geplant(v["check_prods"])
    ]

    if not relevante:
        return wplan

    # ── Interaktive Abfrage ─────────────────────────────────────
    nicht_interaktiv = not sys.stdin.isatty()
    print()
    print("=" * 52)
    print("  ZWISCHENPRODUKT-CHECK  (Midweek-Neuplanung ab {})".format(erster_tag))
    print("=" * 52)

    for v in relevante:
        if nicht_interaktiv:
            # Im nicht-interaktiven Modus (z.B. scheduled task): alles als erledigt annehmen
            print("  [auto] {} -> als erledigt angenommen".format(v["id"]))
            continue

        while True:
            try:
                antwort = input("  {} (j/n): ".format(v["frage"])).strip().lower()
            except (EOFError, KeyboardInterrupt):
                antwort = "j"
            if antwort in ("j", "ja", "y", "yes", ""):
                print("    Gut — wird als erledigt behandelt.")
                break
            elif antwort in ("n", "nein", "no"):
                print("    -> Wird auf {} vorgezogen.".format(erster_tag))
                # Aufgabe mit hoher Prioritaet vorne einfuegen
                aufgabe = {
                    "produkt":     v["aufgabe"],
                    "menge":       0,
                    "notiz":       v["notiz"],
                    "prio":        v["prio"],
                    "excel_only":  False,
                    "produkt_key": None,
                    "aktiv_min":   v["aktiv_min"],
                    "passiv_min":  v["passiv_min"],
                }
                wplan[erster_tag].append(aufgabe)
                # Neu sortieren (wie am Ende von erstelle_wochenplan)
                wplan[erster_tag].sort(key=lambda x: x["prio"])
                break
            else:
                print("    Bitte j (ja) oder n (nein) eingeben.")
    print()
    return wplan

# ============================================================
# MAIN
# ============================================================
def main():
    print("\nKuchenproduktion Wochenplaner")
    print("="*40)
    print("Lade Konfiguration...")
    lade_konfiguration()
    print("Lade Rezepte...")
    rezepte=lade_alle_rezepte()
    print("  {} Rezepte geladen: {}".format(len(rezepte),list(rezepte.keys())))
    sub_rezepte=lade_sub_rezepte()
    print("  {} Sub-Rezepte geladen: {}".format(len(sub_rezepte),list(sub_rezepte.keys())))
    print("Lese Verkaufszahlen...")
    verkauf=lese_verkaufszahlen()
    print("Lese Lagerbestand...")
    lager=lese_lagerbestand()
    print("  Lager: {}".format(lager))
    print("Berechne Wochenbedarf...")
    c2_erledigt = cafe2_bereits_abgeholt()
    if c2_erledigt:
        print("  Cafe-2-Abholung bereits erledigt — Lager ist post-Abgabe")
    bedarf=berechne_wochenbedarf(verkauf,lager,cafe2_erledigt=c2_erledigt)
    print("Berechne Muerbeteig...")
    murt=berechne_murbeteig(bedarf)
    vtage = ab_heute_tage()
    ist_midweek = vtage != list(ARBEITSTAGE)
    if ist_midweek:
        print("  Ab-heute Modus: plane ab {} ({} Tage)".format(vtage[0], len(vtage)))
    print("Erstelle Wochenplan...")
    wplan=erstelle_wochenplan(bedarf,murt,rezepte,sub_rezepte,verbleibende_tage=vtage)
    print("Berechne Rolling Inventory...")
    rolling, tages_info, engpaesse = berechne_rolling_inventory(bedarf, lager, wplan)
    if engpaesse:
        print("  !! ENGPAESSE ERKANNT:")
        for tag, key, bestand in engpaesse:
            print("     {} - {}: {:.1f} Stk (Lager leer!)".format(tag, key, bestand))
    else:
        print("  Lagerbestand OK fuer alle Tage")
    if ist_midweek:
        wplan = frage_vorbereitungen(bedarf, wplan, vtage, sub_rezepte)
        # Rolling Inventory mit ggf. angepasstem wplan neu berechnen
        rolling, tages_info, engpaesse = berechne_rolling_inventory(bedarf, lager, wplan)
    print("\nErstelle Ausgabedateien...")
    erstelle_excel(bedarf,murt,wplan)
    for tag in vtage:
        erstelle_html(tag,wplan.get(tag,[]),murt,rezepte,sub_rezepte,
                      lager_info=tages_info.get(tag), ist_heut=ist_midweek)
    print("\n"+"="*40+"  WOCHENBEDARF:")
    for key,bd in bedarf.items():
        s="-> {} Stk".format(bd["zu_produzieren"]) if bd["zu_produzieren"]>0 else "-> Lager reicht"
        print("  {}: {}".format(key,s))
        if bd.get("grund"): print("     {}".format(bd["grund"]))
    print("\nMuerbeteig: {}x Grundrezept ({:.0f}g)".format(murt["grundrezepte"],murt["gesamt_gramm"]))
    print("Fertig! Ordner: {}".format(OUTPUT_PFAD))

main()
