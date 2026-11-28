#!/usr/bin/env python3
"""Prüft die ganze Excel-Kette (Vorlagen-Prinzip, Stand V8) mit LibreOffice.

Aufbau: Ordner „Schichtplanung“ (offen: Platz 01 … , 001 MA-Tagesplan, 002 Schichtplan) und „Steuerung“ (privat:
Stundenzettel 01 … , 000 Inhaltsverzeichnis). Platz-/Stundenzettel-Dateien sind unveränderte KOPIEN je einer Vorlage und
erkennen ihre Nummer am Dateinamen. Namen stehen nur im Inhaltsverzeichnis.
Ablauf (bildet Excel nach – Excel liest bei geschlossenen Dateien die GESPEICHERTEN Werte):
  1. Paket mit 4 Plätzen erzeugen (Anna, Ben mit Testwerten, Cora ohne, Platz 04 frei)
  1b. Verknüpfungen liefern Fehler → kein #BEZUG!, sondern „keine Vorgabe“/leer; falscher Dateiname → Warnung
  2. Eingabe-Dateien und Stundenzettel einmal speichern (noch ohne Aktualisieren) → „noch keine Vorgabe“
  3. Inhaltsverzeichnis aktualisieren (Vorgaben der Stundenzettel) → speichern
  4. Schichtplan (Namen aus dem Inhaltsverzeichnis, nach Nachname) und Tagesplan (nach Platz-Nummer) aktualisieren
  5. Stundenzettel aus dem gespeicherten Schichtplan → muss genau wie die Hauptvorlage rechnen; 057-Wert
  6. Inhaltsverzeichnis erneut → Eingabe-Dateien holen Name und Vorgaben (Reststunden = Stundenzettel, Urlaubswarnung)
  7. Vertragsende: Ben ohne Vertrag im Januar → Eingabe „kein Vertrag“ → Schichtplan ohne Zeiten
Die echte Verknüpfung wird simuliert: jeder Bezug [n]Blatt!… zeigt auf ein lokales Blatt mit den gespeicherten Werten
der Quelldatei (auch Bereiche für INDEX). LibreOffice selbst verwechselt beim Aktualisieren Werte zwischen Dateien.

Start: python3 werkzeuge/plaene.test.py      Ende: „ERGEBNIS: alles in Ordnung“ (ca. 3 Min)
"""

import os, re, shutil, subprocess, sys, tempfile, time
import openpyxl
import uno
from com.sun.star.beans import PropertyValue

HIER = os.path.dirname(os.path.abspath(__file__))
TMP = tempfile.mkdtemp(prefix='plaene-')
ORDNER = os.path.join(TMP, 'Abteilung')
PROFIL = os.path.join(TMP, 'profil')
os.makedirs(os.path.join(PROFIL, 'user'))
open(os.path.join(PROFIL, 'user', 'registrymodifications.xcu'), 'w').write('''<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>0</value></prop></item>
</oor:items>''')
MA = ['Beispiel, Anna', 'Test, Ben', 'Abel, Cora']           # Platz 01, 02, 03; Platz 04 frei
subprocess.run([sys.executable, os.path.join(HIER, 'plaene-erzeugen.py'), '--ordner', ORDNER, '--plaetze', '4',
                '--wochen', '1', '--tage', '5', '--mitarbeiter', *MA, '--kopien', '--beispieldaten'], check=True)
IV, TP, SP = '000 Inhaltsverzeichnis.xlsx', '001 MA-Tagesplan.xlsx', '002 Schichtplan.xlsx'
PL = [f'Platz {i:02d}.xlsx' for i in range(1, 5)]
SZ = [f'Stundenzettel {i:02d}.xlsx' for i in range(1, 5)]


def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p


PORT = 2100 + os.getpid() % 800
proc = subprocess.Popen(['soffice', '--headless', '--norestore', f'-env:UserInstallation=file://{PROFIL}',
                         f'--accept=socket,host=localhost,port={PORT};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for _ in range(120):
    try:
        lc = uno.getComponentContext()
        ctx = lc.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', lc).resolve(
            f'uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext'); break
    except Exception:
        time.sleep(0.5)
desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
OFFEN, PRIVAT = 'Schichtplanung', 'Steuerung'
pfad = lambda n: os.path.join(ORDNER, n)


def oeffnen(name, links=False):
    return desk.loadComponentFromURL(uno.systemPathToFileUrl(pfad(name)), '_blank', 0,
                                     (pv('Hidden', True), pv('UpdateDocMode', 3 if links else 0)))


def speichern(doc, name):
    doc.storeToURL(uno.systemPathToFileUrl(pfad(name)), (pv('FilterName', 'Calc MS Excel 2007 XML'),))


def wert(doc, blatt, zelle):
    return doc.Sheets.getByName(blatt).getCellRangeByName(zelle).getString()


fehler = 0
def pruefe(text, ist, soll):
    global fehler
    ok = ist == soll
    fehler += not ok
    print(f"{'OK  ' if ok else 'FEHLER'} {text}: {ist!r}" + ('' if ok else f' (erwartet {soll!r})'), flush=True)


REF = re.compile(r"\[(\d+)\]([A-Za-z]+)!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?")
from openpyxl.utils import column_index_from_string as CI
from urllib.parse import unquote
ORIGINAL = os.path.join(TMP, 'mit-verknuepfung')
for wurzel, _, dateien in os.walk(ORDNER):          # Ausgangsstand mit Verknüpfungen aufheben
    for f in dateien:
        rel = os.path.relpath(os.path.join(wurzel, f), ORDNER)
        os.makedirs(os.path.dirname(os.path.join(ORIGINAL, rel)), exist_ok=True)
        shutil.copy(os.path.join(wurzel, f), os.path.join(ORIGINAL, rel))


def aktualisieren_wie_excel(datei):
    """Wie Excel „Werte aktualisieren“: jeder Bezug [n]Blatt!… liest die GESPEICHERTEN Werte der Quelldatei n
    (leer → 0). Ausgangspunkt ist immer der Stand mit Verknüpfungen; lokale Änderungen im Test gehen dabei verloren."""
    wb = openpyxl.load_workbook(os.path.join(ORIGINAL, datei))
    basis = os.path.dirname(pfad(datei))
    ziele = [os.path.normpath(os.path.join(basis, unquote(lk.file_link.Target))) for lk in wb._external_links]
    quellen, boxen = {}, {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    for m in REF.finditer(c.value):
                        n, bl = int(m[1]), m[2]
                        c1, r1 = CI(m[3]), int(m[4]); c2, r2 = (CI(m[5]), int(m[6])) if m[5] else (c1, r1)
                        b = boxen.get((n, bl), (c1, r1, c2, r2))
                        boxen[(n, bl)] = (min(b[0], c1), min(b[1], r1), max(b[2], c2), max(b[3], r2))
    for (n, bl), (c1, r1, c2, r2) in boxen.items():
        if n not in quellen: quellen[n] = openpyxl.load_workbook(ziele[n - 1], data_only=True)
        q = quellen[n][bl]; lok = wb.create_sheet(f'X{n}{bl}')
        for r in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                v = q.cell(r, cc).value
                lok.cell(r, cc).value = 0 if v is None else v
        lok.sheet_state = 'hidden'
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('=') and REF.search(c.value):
                    c.value = re.sub(r"\[(\d+)\]([A-Za-z]+)!", lambda m: f"X{m[1]}{m[2]}!", c.value)
    wb._external_links = []
    wb.save(pfad(datei))


def verknuepfung_fehler(datei, ziel):
    """Kopie von datei, in der jede Verknüpfung einen Fehler liefert (wie Excel mit #BEZUG!, wenn die Quelle
    noch nie gespeichert wurde oder nicht erreichbar ist)."""
    wb = openpyxl.load_workbook(pfad(datei))
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and REF.search(c.value): c.value = '=1/0'
    wb._external_links = []
    wb.save(pfad(ziel))


def eingabe_speichern(datei):
    """Eingabe-Datei „in Excel“ öffnen, rechnen, speichern. ZELLE("Dateiname") wird wie in Excel eingesetzt."""
    d = oeffnen(datei)
    d.Sheets.getByName('Listen').getCellRangeByName('O2').setString('C:\\Abteilung\\[' + os.path.basename(datei) + ']Dateneingabe')
    d.calculateAll(); speichern(d, datei); return d


def rechnen_speichern(datei):
    d = oeffnen(datei); d.calculateAll(); speichern(d, datei); return d


def keine_fehler(doc, blatt, bereich):
    alle = doc.Sheets.getByName(blatt).getCellRangeByName(bereich).getDataArray()
    return sum(1 for z in alle for v in z if isinstance(v, str) and (v.startswith('#') or v.startswith('Err:')))


O, P = OFFEN + '/', PRIVAT + '/'

# 1b. Verknüpfungen liefern Fehler → Dateien rechnen trotzdem
print('--- Verknüpfung liefert Fehler')
verknuepfung_fehler(O + PL[0], O + 'fehler-eingabe.xlsx')
d = oeffnen(O + 'fehler-eingabe.xlsx'); d.calculateAll()
pruefe('Eingabe: Start vorbelegt', wert(d, 'Dateneingabe', 'D7'), '11:00')
pruefe('Eingabe: Pause vorbelegt', wert(d, 'Dateneingabe', 'E7'), '30 Min')
pruefe('Eingabe: Gilt Start', wert(d, 'Dateneingabe', 'G7'), '11:00')
pruefe('Eingabe: Hinweis statt Fehler', wert(d, 'Dateneingabe', 'H3'), 'noch keine Vorgabe vom Steuerer')
pruefe('Eingabe: Titel mit Jahr', wert(d, 'Dateneingabe', 'A1'), 'Dateneingabe 2027')
pruefe('Eingabe: keine Fehlerwerte', keine_fehler(d, 'Dateneingabe', 'A1:N267'), 0)
d.close(True)
verknuepfung_fehler(P + SZ[0], P + 'fehler-sz.xlsx')
d = oeffnen(P + 'fehler-sz.xlsx'); d.calculateAll()
pruefe('Stundenzettel: Name leer statt Fehler', wert(d, 'Stammdaten', 'B3'), '')
pruefe('Stundenzettel: Tagesliste leer statt Fehler', wert(d, 'Tagesliste', 'D7') + wert(d, 'Tagesliste', 'E7') + wert(d, 'Tagesliste', 'F7'), '')
pruefe('Stundenzettel: keine Fehlerwerte', keine_fehler(d, 'Tagesliste', 'A1:J267'), 0)
d.close(True)
for datei in (O + SP, O + TP):
    verknuepfung_fehler(datei, O + 'fehler-x.xlsx')
    d = oeffnen(O + 'fehler-x.xlsx'); d.calculateAll()
    pruefe(f'{datei}: kein Fehler', wert(d, 'Mitarbeiter', 'C4') + wert(d, 'Mitarbeiter', 'J4'), '')
    d.close(True)

print('--- Falscher Dateiname')
shutil.copy(pfad(O + PL[0]), pfad(O + 'Platz 01 - Kopie.xlsx'))
d = eingabe_speichern(O + 'Platz 01 - Kopie.xlsx')
pruefe('Eingabe „Platz 01 - Kopie“: Warnung', wert(d, 'Dateneingabe', 'A4')[:30], 'Achtung: Dateiname falsch oder')
pruefe('Eingabe „Platz 01 - Kopie“: Nr. 0', wert(d, 'Listen', 'O10'), '0')
d.close(True)
shutil.copy(pfad(P + SZ[0]), pfad(P + 'Stundenzettel 1a.xlsx'))
d = rechnen_speichern(P + 'Stundenzettel 1a.xlsx')
pruefe('Stundenzettel „1a“: Warnung', wert(d, 'Stammdaten', 'D3')[:15], 'Dateiname falsc')
d.close(True)
for f in (O + 'Platz 01 - Kopie.xlsx', P + 'Stundenzettel 1a.xlsx'): os.remove(pfad(f))

# 2. Vorlagen-Kopien einmal speichern (wie der Nutzer: öffnen, speichern) – noch ohne Aktualisieren
print('--- Erstes Speichern (ohne Aktualisieren)')
for i, f in enumerate(PL):
    d = eingabe_speichern(O + f)
    pruefe(f'{f}: Platz-Nr. aus Dateiname', wert(d, 'Listen', 'O10'), str(i + 1))
    if i == 0:
        pruefe('Eingabe: Blätter', ','.join(d.Sheets.ElementNames), 'Musterwoche,Dateneingabe,Urlaub,Listen,Export,Vorgaben')
        pruefe('Eingabe ohne Vorgaben: Reststunden-Hinweis', wert(d, 'Dateneingabe', 'H3'), 'noch keine Vorgabe vom Steuerer')
        pruefe('Eingabe ohne Vorgaben: keine Warnung', wert(d, 'Dateneingabe', 'A4'), '')
        pruefe('Eingabe ohne Vorgaben: Vertrag angenommen', wert(d, 'Dateneingabe', 'G7'), '11:00')
        pruefe('Eingabe ohne Vorgaben: Titel ohne Namen', wert(d, 'Dateneingabe', 'A1'), 'Dateneingabe 2027')
    d.close(True)
for i, f in enumerate(SZ):
    d = rechnen_speichern(P + f)
    pruefe(f'{f}: Nr. aus Dateiname', wert(d, 'Daten', 'B2'), str(i + 1))
    if i == 0:
        pruefe('Stundenzettel: Hinweis „noch kein Name“', wert(d, 'Stammdaten', 'D3')[:15], 'Noch kein Name ')
        pruefe('Stundenzettel ohne Anleitung', 'Anleitung' in d.Sheets.ElementNames, False)
    d.close(True)

# 3. Inhaltsverzeichnis: Vorgaben aus den Stundenzetteln
print('--- Inhaltsverzeichnis')
ivl = openpyxl.load_workbook(pfad(P + IV))['Inhaltsverzeichnis']
pruefe('Inhaltsverzeichnis: Klick-Link', ivl['D5'].hyperlink.target, 'Stundenzettel 02.xlsx')
pruefe('Inhaltsverzeichnis: Name-Feld entsperrt', ivl['B4'].protection.locked, False)
aktualisieren_wie_excel(P + IV)
d = rechnen_speichern(P + IV)
for i, n in enumerate(MA + ['']):
    pruefe(f'Inhaltsverzeichnis Zeile {i + 1}', ' | '.join(wert(d, 'Inhaltsverzeichnis', f'{c}{4 + i}') for c in 'ABCD'),
           f'{i + 1} | {n} | Platz {i + 1:02d} | Stundenzettel {i + 1:02d}')
pruefe('Vorgaben Zeile 2: Name / Jahr / Urlaub', ' / '.join(wert(d, 'Vorgaben', f'{c}3') for c in 'BCD'), 'Test, Ben / 2027 / 30')
pruefe('Vorgaben Zeile 4 (frei): Name leer', wert(d, 'Vorgaben', 'B5'), '')
d.close(True)

# 4. Schichtplan und Tagesplan aus den Eingabe-Dateien + Namen aus dem Inhaltsverzeichnis
aktualisieren_wie_excel(O + SP)
sp = rechnen_speichern(O + SP)                    # Steuerer speichert den Schichtplan (Quelle der Stundenzettel)
W = '04.01.2027'
print('--- Schichtplan, Woche', W, '(nach Nachname)')
pruefe('Schichtplan ohne Anleitung', 'Anleitung' in sp.Sheets.ElementNames, False)
pruefe('Zeile 1 Name (Abel)', wert(sp, W, 'A4'), 'Abel, Cora')
pruefe('Zeile 2 Name', wert(sp, W, 'A5'), 'Beispiel, Anna')
pruefe('Zeile 3 Name', wert(sp, W, 'A6'), 'Test, Ben')
pruefe('freier Platz ausgeblendet', wert(sp, W, 'A7'), '')
pruefe('freier Platz Mo leer', wert(sp, W, 'B7'), '')
pruefe('Cora Mo Beginn (Vorlage)', wert(sp, W, 'B4'), '11:00')
pruefe('Anna Mo Beginn', wert(sp, W, 'B5'), '11:00')
pruefe('Anna Mo Pause', wert(sp, W, 'C5'), '30')
pruefe('Anna Mo Anmerkung (Musterwoche)', wert(sp, W, 'D5'), 'Homeoffice')
pruefe('Anna Di Urlaub', wert(sp, W, 'E5'), 'Urlaub')
pruefe('Anna Di Pause leer', wert(sp, W, 'F5'), '')
pruefe('Anna Di Anmerkung leer', wert(sp, W, 'G5'), '')
pruefe('Anna Fr Beginn', wert(sp, W, 'N5'), '12:00')
pruefe('Ben Mo Beginn', wert(sp, W, 'B6'), '13:00')
pruefe('Ben Mo Pause', wert(sp, W, 'C6'), '45')
pruefe('Ben Mi abwesend (Gutschrift) → abwesend', wert(sp, W, 'H6'), 'abwesend')
pruefe('Ben Mi keine Anmerkung (kein Grund)', wert(sp, W, 'J6'), '')
pruefe('Ben Do Teammentor*in', wert(sp, W, 'M6'), 'Teammentor*in')
pruefe('Ben Fr frei', wert(sp, W, 'N6'), 'Frei')
pruefe('Datum Kopf', wert(sp, W, 'B2')[:6], 'Montag')
if os.environ.get('BILDER'):
    sp.storeToURL(uno.systemPathToFileUrl(os.path.join(os.environ['BILDER'], 'schichtplan.pdf')), (pv('FilterName', 'calc_pdf_Export'),))
sp.close(True)

aktualisieren_wie_excel(O + TP)
tp = oeffnen(O + TP); tp.calculateAll()
print('--- MA-Tagesplan (nach Platz-Nummer)')
T = lambda d, z: wert(tp, d, z)
pruefe('04.01. Zeile 1 = Platz 01 Vorname', T('04.01.2027', 'B3'), 'Anna')
pruefe('04.01. Zeile 1 Nachname', T('04.01.2027', 'C3'), 'Beispiel')
pruefe('04.01. Zeile 2 = Platz 02', T('04.01.2027', 'C4'), 'Test')
pruefe('04.01. Zeile 3 = Platz 03 (nicht alphabetisch)', T('04.01.2027', 'C5'), 'Abel')
pruefe('04.01. Zeile 4 frei', T('04.01.2027', 'C6'), '')
pruefe('04.01. Anna Beginn', T('04.01.2027', 'D3'), '11:00')
pruefe('04.01. Anna HO', T('04.01.2027', 'F3'), 'HO')
pruefe('04.01. Ben Beginn', T('04.01.2027', 'D4'), '13:00')
pruefe('04.01. Ben Pause', T('04.01.2027', 'E4'), '45')
pruefe('05.01. Anna Urlaub (ausgegraut)', T('05.01.2027', 'F3'), 'Urlaub')
pruefe('05.01. Anna kein Beginn', T('05.01.2027', 'D3'), '')
pruefe('05.01. Anna Name bleibt stehen', T('05.01.2027', 'C3'), 'Beispiel')
pruefe('06.01. Ben abwesend', T('06.01.2027', 'F4'), 'abwesend')
pruefe('07.01. Ben Teammentor*in', T('07.01.2027', 'F4'), 'Teammentor*in')
pruefe('08.01. Ben frei', T('08.01.2027', 'F4'), 'Frei')
pruefe('08.01. Anna 12:00', T('08.01.2027', 'D3'), '12:00')
if os.environ.get('BILDER'):
    tp.storeToURL(uno.systemPathToFileUrl(os.path.join(os.environ['BILDER'], 'tagesplan.pdf')), (pv('FilterName', 'calc_pdf_Export'),))
tp.close(True)

# Sollwerte: Hauptvorlage mit denselben Eingaben wie Ben
ref = os.path.join(TMP, 'referenz.xlsx')
wbr = openpyxl.load_workbook(os.path.join(HIER, '..', 'excel', 'Arbeitszeiterfassung.xlsx'))
import datetime as _dt
for r in range(5, 10): wbr['Musterwoche'][f'B{r}'] = _dt.time(13, 0); wbr['Musterwoche'][f'C{r}'] = 45
wbr['Dateneingabe']['F9'] = 'abwesend (Gutschrift)'; wbr['Dateneingabe']['F10'] = 'Teammentor*in'; wbr['Dateneingabe']['D11'] = 'Frei'
wbr.save(ref)
rd = desk.loadComponentFromURL(uno.systemPathToFileUrl(ref), '_blank', 0, (pv('Hidden', True),)); rd.calculateAll()
soll = {z: wert(rd, *z.split('!')) for z in ('Auswertung!C26', 'Auswertung!C27', 'Auswertung!C28', 'Urlaub!B4', 'Auswertung!B5', 'Stammdaten!B27')}
rd.close(True)

# 5. Stundenzettel aus dem gespeicherten Schichtplan (eigene Spalten über die Nummer im Dateinamen)
print('--- Stundenzettel')
for f in SZ:
    aktualisieren_wie_excel(P + f)
sz = oeffnen(P + SZ[1]); sz.calculateAll()
pruefe('Name aus dem Inhaltsverzeichnis (über Schichtplan)', wert(sz, 'Stammdaten', 'B3'), 'Test, Ben')
pruefe('kein Hinweis mehr', wert(sz, 'Stammdaten', 'D3'), '')
for blatt, titel in (('Stammdaten', 'Stammdaten'), ('Tagesliste', 'Tagesliste 2027'), ('Urlaub', 'Urlaubsübersicht & Abgleich'),
                     ('Auswertung', 'Auswertung & Jahresbilanz 2027')):
    pruefe(f'Stundenzettel Titel {blatt}', wert(sz, blatt, 'A1'), titel + ' – Test, Ben')
pruefe('Tagesliste 04.01. Beginn', wert(sz, 'Tagesliste', 'G7'), '13:00')
pruefe('Tagesliste 06.01. abwesend, Zeit bleibt', wert(sz, 'Tagesliste', 'G9') + ' / ' + wert(sz, 'Tagesliste', 'I9'), '13:00 / abwesend (Gutschrift)')
pruefe('Tagesliste 07.01. Teammentor*in', wert(sz, 'Tagesliste', 'I10'), 'Teammentor*in')
pruefe('Tagesliste 08.01. Frei', wert(sz, 'Tagesliste', 'G11'), 'Frei')
for z, v in soll.items():
    pruefe(f'wie Hauptvorlage {z}', wert(sz, *z.split('!')), v)
sz.Sheets.getByName('Auswertung').getCellRangeByName('C5').setValue(100 / 24); sz.calculateAll()    # 057 Januar
ben_rest_sek = sz.Sheets.getByName('Auswertung').getCellRangeByName('G28').getValue()
pruefe('Vorgaben: Jahr / Anspruch', wert(sz, 'Vorgaben', 'A1') + ' / ' + wert(sz, 'Vorgaben', 'B1'), '2027 / 30')
speichern(sz, P + SZ[1]); sz.close(True)
d = oeffnen(P + SZ[0])
pruefe('Stundenzettel 01: Name Anna', wert(d, 'Stammdaten', 'B3') if d.calculateAll() is None else '', 'Beispiel, Anna')
pruefe('Stundenzettel 01: 05.01. Urlaub', wert(d, 'Tagesliste', 'G8'), 'Urlaub')
d.Sheets.getByName('Stammdaten').getCellRangeByName('B5').setValue(0.25); d.calculateAll()    # Anna: Anspruch 0,25 Tage → Warnung
speichern(d, P + SZ[0]); d.close(True)
d = oeffnen(P + SZ[2]); d.calculateAll()
pruefe('Stundenzettel 03: Name Cora', wert(d, 'Stammdaten', 'B3'), 'Abel, Cora')
speichern(d, P + SZ[2]); d.close(True)
d = oeffnen(P + SZ[3]); d.calculateAll()
pruefe('Stundenzettel 04 (frei): Name leer', wert(d, 'Stammdaten', 'B3'), '')
speichern(d, P + SZ[3]); d.close(True)

# 6. Inhaltsverzeichnis erneut, dann holen die Eingabe-Dateien Name und Vorgaben
print('--- Eingabe-Datei mit Name und Vorgaben')
aktualisieren_wie_excel(P + IV); rechnen_speichern(P + IV).close(True)
for f in PL:
    aktualisieren_wie_excel(O + f); eingabe_speichern(O + f).close(True)
d = oeffnen(O + PL[1]); d.calculateAll()
pruefe('Ben: Name im Export', wert(d, 'Export', 'B1'), 'Test, Ben')
pruefe('Ben: Titel Dateneingabe', wert(d, 'Dateneingabe', 'A1'), 'Dateneingabe 2027 – Test, Ben')
pruefe('Ben: Titel Musterwoche', wert(d, 'Musterwoche', 'A1'), 'Musterwoche (Montag bis Freitag) – Test, Ben')
pruefe('Ben: Titel Urlaub', wert(d, 'Urlaub', 'A1'), 'Urlaubsübersicht – Test, Ben')
pruefe('Ben: Reststunden-Text', wert(d, 'Dateneingabe', 'H3').split(':')[0], 'noch zu leisten' if ben_rest_sek > 0 else 'zu viel geleistet')
pruefe('Ben: Reststunden = Stundenzettel (Sek.)', round(d.Sheets.getByName('Listen').getCellRangeByName('O9').getValue()), round(ben_rest_sek))
pruefe('Ben: Urlaub Anspruch', wert(d, 'Urlaub', 'B4'), '30 Tage')
if os.environ.get('BILDER'):
    d.storeToURL(uno.systemPathToFileUrl(os.path.join(os.environ['BILDER'], 'eingabe.pdf')), (pv('FilterName', 'calc_pdf_Export'),))
d.close(True)
d = oeffnen(O + PL[0]); d.calculateAll()
pruefe('Anna: Urlaubswarnung (Anspruch 0,25)', wert(d, 'Dateneingabe', 'A4').replace(',', '.'), 'Achtung: 1 Urlaubstage eingetragen – Anspruch: 0.25 Tage!')
d.close(True)
d = oeffnen(O + PL[3]); d.calculateAll()
pruefe('Platz 04 (frei): Titel ohne Namen', wert(d, 'Dateneingabe', 'A1'), 'Dateneingabe 2027')
d.close(True)

# 7. Vertragsende: Ben im Januar ohne Vertrag → Eingabe-Datei „kein Vertrag“ → Schichtplan ohne Zeiten
print('--- Vertragsende')
sz = oeffnen(P + SZ[1])
sz.Sheets.getByName('Stammdaten').getCellRangeByName('B14').setValue(2028 / 24)
sz.Sheets.getByName('Stammdaten').getCellRangeByName('B13').setValue(0); sz.calculateAll()
pruefe('Vorgabe Januar ohne Vertrag = 1', wert(sz, 'Vorgaben', 'A3'), '1')
speichern(sz, P + SZ[1]); sz.close(True)
aktualisieren_wie_excel(P + IV); rechnen_speichern(P + IV).close(True)
aktualisieren_wie_excel(O + PL[1]); d = eingabe_speichern(O + PL[1])
pruefe('Eingabe Ben 04.01. „kein Vertrag“', wert(d, 'Dateneingabe', 'G7'), 'kein Vertrag'); d.close(True)
aktualisieren_wie_excel(O + SP)
sp = oeffnen(O + SP); sp.calculateAll()
pruefe('Ben steht noch in der Liste', wert(sp, W, 'A6'), 'Test, Ben')
pruefe('Ben Mo ohne Zeit (kein Vertrag)', wert(sp, W, 'B6'), '')
pruefe('Ben Do ohne Anmerkung (kein Vertrag)', wert(sp, W, 'M6'), '')
pruefe('Anna unverändert', wert(sp, W, 'B5'), '11:00')
sp.close(True)
try:
    desk.terminate()
except Exception:
    pass
proc.wait(timeout=30)
print('Testordner:', ORDNER)
print('ERGEBNIS: alles in Ordnung' if fehler == 0 else f'ERGEBNIS: {fehler} Fehler')
sys.exit(1 if fehler else 0)
