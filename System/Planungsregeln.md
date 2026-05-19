# Planungsregeln — Kuchenproduktion

Diese Datei beschreibt die Verteilungslogik für den Wochenplan.
Der Planer liest diese Regeln und wendet sie automatisch an.

---

## Allgemeines

- Arbeitstage: Montag, Dienstag, Mittwoch, Donnerstag, Freitag
- Café 1: bekommt frische Ware aus laufender Produktion
- Café 2: bekommt Dienstags Ware aus dem Froster (Vorwoche)
- Produktion immer in vollen Chargen (= max_ofen Stück)

---

## Tiebreaker-Reihenfolge (gilt für alle Produkte)

Wenn mehrere Tage gleich viele Chargen haben und einer bevorzugt werden muss:

1. Freitag
2. Montag
3. Mittwoch
4. Donnerstag
5. Dienstag

Hintergrund: Freitag sichert Wochenend-Frische, Montag und Mittwoch sind wochenmittig wichtig.

---

## Produkt: Käsekuchen

**Charge = 2 Stück (max_ofen = 2), max. 4 Stück pro Tag (2 Chargen)**

### Feste Mindestchargen pro Tag

| Tag       | Mindestchargen | Mindeststück |
|-----------|---------------|--------------|
| Montag    | 1             | 2            |
| Mittwoch  | 1             | 2            |
| Freitag   | 2             | 4            |
| Dienstag  | 0             | —            |
| Donnerstag| 0             | —            |

Hintergrund:
- Montag: erster Arbeitstag nach dem Frei, frische Ware wichtig
- Mittwoch: Wochenmitte, Lager auffüllen
- Freitag: zwei Chargen sichern frische Ware für das Wochenende

### Verteilung der Restchargen

Nachdem die Mindestchargen verteilt sind, werden übrige Chargen wie folgt verteilt:

1. **Schritt:** Tage mit den wenigsten Chargen bekommen Priorität (beginne mit 0, dann 1, dann 2, ...)
2. **Tiebreaker:** Bei Gleichstand gilt die allgemeine Tiebreaker-Reihenfolge (Fr → Mo → Mi → Do → Di)
3. Wiederhole bis alle Chargen verteilt sind

**Beispiel** (14 Stück benötigt = 7 Chargen):

| Schritt | Aktion                                      | Mo | Di | Mi | Do | Fr |
|---------|---------------------------------------------|----|----|----|----|----|
| Start   | Mindestchargen verteilen                    | 1  | 0  | 1  | 0  | 2  |
| +1      | Di=0, Do=0 → Tiebreaker → Do zuerst        | 1  | 0  | 1  | 1  | 2  |
| +1      | Nur Di=0 → Di bekommt Charge               | 1  | 1  | 1  | 1  | 2  |
| +1      | Alle ≥1, Fr=2 hat mehr → Mo/Di/Mi/Do alle =1 → Mo zuerst | 2 | 1 | 1 | 1 | 2 |
| **Ende**| **7 Chargen verteilt**                      | **2** | **1** | **1** | **1** | **2** |

Bei 11 Chargen (alle Tage auf 2 aufgefüllt, dann Fr als erstes auf 3):
Mo=2, Di=2, Mi=2, Do=2, Fr=3

### Tagesreihenfolge

Wenn an einem Tag Käsekuchen steht, ist das **immer die erste Aufgabe des Morgens**.
Begründung: lange Backzeit → während der Ofen läuft, kann man anderes vorbereiten.

### Böden-Herkunft nach Tag

| Tag        | Produkt          | Böden kommen von                     |
|------------|------------------|--------------------------------------|
| Montag     | alle             | Vorwoche (Dienstag der Vorwoche)     |
| Dienstag   | KK               | Vorwoche — KK wird **vor** dem Ausrollen hergestellt |
| Dienstag   | andere Produkte  | Diese Woche — nach dem Ausrollen verfügbar |
| Mittwoch   | alle             | Diese Woche (Dienstag dieser Woche)  |
| Donnerstag | alle             | Diese Woche (Dienstag dieser Woche)  |
| Freitag    | alle             | Diese Woche (Dienstag dieser Woche)  |

---

## Mürbeteig-Zyklus

### Feste Tage

- **Montag:** Mürbeteig herstellen und über Nacht kalt stellen
- **Dienstag:** Mürbeteig ausrollen und backen (alle Böden für die Woche)

### Mengenkalkulation

Der Dienstags-Vorrat muss reichen von **Dienstag (nach dem Ausrollen) bis zum nächsten Dienstag (vor dem Ausrollen)**.

Das bedeutet: der Dienstag-Vorrat versorgt:
- Dienstag dieser Woche — Produkte **nach** dem Ausrollen (nicht KK)
- Mittwoch, Donnerstag, Freitag dieser Woche (alle Produkte mit Mürbeteig)
- Montag nächste Woche (alle Produkte mit Mürbeteig)
- Dienstag nächste Woche — **nur KK** (wird vor dem Ausrollen hergestellt; andere Dienstags-Produkte nutzen dann den nächsten frischen Mürbeteig)

→ Reserve für nächste Woche: geschätzt wie diese Woche Mo (alle) + Di KK

---

---

## Produkt: Streusel

- Streusel ist immer auf der Karte (Käse-Streusel Schnitte ist ein Dauerbrenner)
- Wird **immer Montags** hergestellt, für die gesamte Woche (Mo–Fr)
- Menge: reicht für alle Käse-Streusel Schnitte, die diese Woche produziert werden
- In den Tagesplänen Di–Fr wird das Streusel-Rezept **nicht** angezeigt (bereits fertig)

---

## Fester Tagesstart: Montag

Montag beginnt immer mit dieser fixen Reihenfolge:

| Reihenfolge | Aufgabe            | Notiz                                              |
|-------------|--------------------|----------------------------------------------------|
| 1           | Lager              | Tagesplan anschauen, benötigte Sachen aus dem Froster / Lager holen |
| 2           | Käsekuchen         | Erste Charge in den Ofen                           |
| 3           | Streusel           | Wird hergestellt während KK im Ofen ist            |
| 4           | Mürbeteig          | Wird hergestellt während KK im Ofen ist, über Nacht kalt stellen |

Danach kommen alle weiteren Produktionen des Tages.

---

---

## Produkt: Vanillecreme

- Wird **immer Dienstags** hergestellt
- Haltbarkeit: muss bis **Donnerstag** verarbeitet sein
- Konsequenz: Alle Produkte, die Vanillecreme enthalten, können nur an **Dienstag, Mittwoch oder Donnerstag** produziert werden

---

## Produkt: Tartelettes

### Grundregeln
- Es gibt immer **zwei Sorten** auf der Karte
- Produktion findet **immer Mittwochs** statt
- Es wird immer die Sorte produziert, von der **weniger im Lager ist**
- Produktion findet nur statt, wenn **mindestens eine Sorte unter 45 Stück** liegt (Schwelle = 45)
- Produktionsmenge: **mindestens 60, maximal 90 Stück**

### Zwischenprodukte mit Übernacht-Kühlung
Wenn ein Tartelette eine Creme enthält, die über Nacht kühlen muss, wird diese **Dienstags** hergestellt und am Mittwoch fertiggestellt.

Aktuelle Beispiele:
- **Beeren Tartelette:** Vanillecreme → Dienstag herstellen, Mittwoch Schlagsahne unterheben und fertigstellen
- **Mango-Passionsfrucht Tartelette:** Mango-Joghurt Fond (alles außer Schlagsahne) → Dienstag kochen und kühlen, Mittwoch Schlagsahne unterheben

---

---

## Produkt: Pistazien-Schoko Törtchen

### Grundregeln
- Produktion immer **Donnerstags**
- Nur produzieren wenn Lager **unter 30 Stück** (Schwelle = 30)
- Menge: immer **60 Stück** (1 Grundrezept, 60 Mousseformen)

### Übernacht-Vorbereitung (Mittwoch)
Zwei Komponenten werden am Mittwoch vorbereitet:

| Komponente       | Mittwoch                              | Donnerstag              |
|-----------------|---------------------------------------|-------------------------|
| Brownieboden    | Herstellen, backen, auskühlen lassen  | Ausstechen              |
| Pistazienganache| Herstellen, über Nacht kalt stellen   | Aufschlagen, aufdressieren |

Pistazienkern und Schoko-Mousse werden Donnerstags frisch hergestellt.

### Produktionsablauf Donnerstag (Reihenfolge)
1. Kern herstellen und schockfrosten (mind. 1 Stunde)
2. Brownieboden ausstechen
3. Schoko-Mousse herstellen
4. Mousse in Formen füllen, Kern eindrücken, Boden auflegen
5. Schockfrosten (mind. 2 Stunden)
6. Aus Form lösen, Pistazienganache aufdressieren

---

---

## Produktkategorie: Käse-Frucht-Streusel Schnitte

Gilt für alle saisonalen Varianten dieser Schnitte (z.B. Käse-Rhabarber, Käse-Mandarin, etc.).
Das aktuelle Rezept im System wird entsprechend ausgetauscht, die Regeln bleiben gleich.

- Charge = immer **6 Stück** (1 Blech)

### Feste Mindestchargen pro Tag

| Tag       | Mindestchargen |
|-----------|---------------|
| Donnerstag| 1             |
| alle anderen | 0          |

### Verteilung der Restchargen

1. Tage mit den wenigsten Chargen bekommen Priorität
2. Tiebreaker bei Gleichstand: **Do → Mo → Di → Mi → Fr**

**Beispiel** (2 Chargen benötigt):
- Nach Mindestcharge: Do=1, Mo=0, Di=0, Mi=0, Fr=0
- Alle anderen Tage haben 0 → Priorität gegenüber Do
- Tiebreaker unter Mo/Di/Mi/Fr → Mo bekommt die 2. Charge
- Ergebnis: Do=1, Mo=1

**Beispiel** (6 Chargen benötigt):
- Nach Mindestcharge: Do=1, Rest=0 → Mo, Di, Mi, Fr bekommen je eine (Tiebreaker: Do→Mo→Di→Mi→Fr, aber Do schon bei 1)
- Nach 5 Chargen: Do=1, Mo=1, Di=1, Mi=1, Fr=1 → alle gleich → 6. Charge nach Tiebreaker → Do
- Ergebnis: Do=2, Mo=1, Di=1, Mi=1, Fr=1

### Zusammenspiel mit Streusel

Streusel wird immer Montags für die gesamte Woche hergestellt (siehe Regel oben).
Die Streuselmenge richtet sich nach der Gesamtproduktion dieser Schnitte in der Woche.

---

## Implementierungsregel: Tageszeitbeschränkung durch Zutaten (Option A)

Der Planer leitet erlaubte Produktionstage automatisch aus den Zutaten ab:

1. Für jedes Produkt werden alle Zutaten und Sub-Rezepte durchsucht
2. Enthält eine Zutat ein zeitkritisches Zwischenprodukt (z.B. Vanillecreme), werden die erlaubten Produktionstage automatisch eingeschränkt
3. Die Haltbarkeitsregel des Zwischenprodukts ist maßgeblich (z.B. Vanillecreme: hergestellt Di, verarbeitet bis Do → Produkt nur Di/Mi/Do möglich)
4. Gilt automatisch für alle zukünftigen Rezepte, die diese Zutat verwenden — keine manuelle Pflege nötig

---

## Einfrierempfehlung

### Grundregel: Verfügbarkeit ab nächstem Tag
Da tagsüber produziert wird, ist ein Produkt in der Regel **erst ab dem Folgetag** im Verkauf verfügbar.
- Produktion Dienstag → Verkauf ab Mittwoch
- Produktion Donnerstag → Verkauf ab Freitag
- Ausnahme: wenn ein Produkt ausnahmsweise noch am selben Tag in den Verkauf geht, aber der Planer rechnet konservativ mit dem Folgetag.

### Berechnung "frisch lassen" pro Produktionstag

Für jeden Backtag gilt:
> **Frisch lassen = Tagesverkaufsschnitt × Tage bis das nächste Batch verfügbar ist**
> **Einfrieren = Produktion − frisch lassen**
> **Obergrenze frisch lassen = Haltbarkeit (Kühlung in Tagen) × Tagesverkaufsschnitt**

**Tage bis nächstes Batch verfügbar** = Abstand in Tagen vom Folgetag der aktuellen Produktion bis zum Folgetag der nächsten Produktion desselben Produkts (oder bis Wochenende wenn kein weiterer Backtag).

**Beispiel** (Käse-Rhabarber Schnitte, Backtage Di + Do, Tagesverkauf = 2 Stk, Haltbarkeit Kühlung = 4 Tage):
- Dienstag produziert 6 Stk → verfügbar ab Mi → nächste Charge Do, verfügbar ab Fr → Fenster = Mi + Do = 2 Tage → frisch lassen = 2 × 2 = 4 Stk → einfrieren = 2 Stk
- Donnerstag produziert 6 Stk → verfügbar ab Fr → nächste Charge Di nächste Woche, verfügbar ab Mi → Fenster = Fr + Sa + So + Mo + Di = 5 Tage → frisch lassen wäre 5 × 2 = 10, aber Haltbarkeit nur 4 Tage → Obergrenze = 4 × 2 = 8 → da nur 6 produziert: alle 6 frisch lassen, 0 einfrieren

Die Empfehlung erscheint direkt im Tagesplan, z.B.:
*"6 Stk backen → 4 frisch (reicht bis Do), 2 einfrieren"*

---

---

## Produktkategorie: Quiche

Gilt für alle Varianten (z.B. Lauch-Speck, Spinat-Feta, etc.).
Das aktuelle Rezept im System wird ausgetauscht, die Regeln bleiben gleich.

- Charge = immer **5 Stück** (1 Dinkelhefeteig-Grundrezept = 5 Böden)

### Feste Mindestchargen pro Tag

| Tag       | Mindestchargen |
|-----------|---------------|
| Montag    | 1             |
| alle anderen | 0          |

### Verteilung der Restchargen

1. Tage mit den wenigsten Chargen bekommen Priorität
2. Tiebreaker bei Gleichstand: **Mo → Do → Di → Mi → Fr**

**Beispiel** (3 Chargen benötigt):
- Nach Mindestcharge: Mo=1, Do=0, Di=0, Mi=0, Fr=0
- Tiebreaker unter Do/Di/Mi/Fr → Do bekommt 2. Charge
- Noch eine übrig, Do=1 wie Di/Mi/Fr → Tiebreaker → Di
- Ergebnis: Mo=1, Do=1, Di=1

---

## Weitere Produkte

*(werden ergänzt)*

