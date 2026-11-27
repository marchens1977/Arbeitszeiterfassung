#!/usr/bin/env python3
"""Prüft Schichtplan und MA-Tagesplan (Kette ohne Makros) mit LibreOffice.

Ablauf (bildet Excel nach), Ordner „Schichtplanung“ (offen) und „Steuerung“ (privat):
  1. Mini-Paket erzeugen (3 Plätze: „Beispiel, Anna“ und „Test, Ben“ mit Testwerten, 1 Platzhalter)
  2. Eingabe-Dateien öffnen, rechnen, speichern (wie ein Mitarbeiter in Excel) – noch ohne Vorgaben des Steuerers
  3./4. Schichtplan und Tagesplan direkt aus den Eingabe-Dateien aktualisieren („Werte aktualisieren“ wie Excel:
     Excel liest bei geschlossenen Dateien die GESPEICHERTEN Werte) → prüfen
  5. Stundenzettel aktualisieren → muss genauso rechnen wie die Hauptvorlage mit denselben Eingaben; 057-Wert eintragen
  6. Eingabe-Dateien holen die Vorgaben (Reststunden = Stundenzettel, Urlaubswarnung)
  1b. Verknüpfung liefert Fehler (Quelle nie gespeichert/nicht erreichbar) → kein #BEZUG!, sondern „keine Vorgabe“
  5b. Inhaltsverzeichnis im Ordner Steuerung: Nr. → Name aus dem Stundenzettel
  7. Vertragsende: Ben ohne Vertrag im Januar → Eingabe „kein Vertrag“ → Schichtplan ohne Zeiten
Warum nicht die echte Verknüpfung in LibreOffice? LibreOffice rechnet verknüpfte Dateien beim Aktualisieren
neu, verwechselt dabei Werte zwischen mehreren Dateien und stürzt teils ab – Excel liest nur gespeicherte Werte.
Dass die Verknüpfung grundsätzlich funktioniert (relativer Pfad, gleicher Ordner), wurde einmalig von Hand geprüft;
in Excel prüft es der Nutzer mit dem Mini-Paket.

Start: python3 werkzeuge/plaene.test.py      Ende: „ERGEBNIS: alles in Ordnung“
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
MA = ['Beispiel, Anna', 'Test, Ben', 'Mustermann, Max 03']
subprocess.run([sys.executable, os.path.join(HIER, 'plaene-erzeugen.py'), '--ordner', ORDNER, '--plaetze', '3',
                '--wochen', '1', '--tage', '5', '--mitarbeiter', MA[0], MA[1], '--beispieldaten'], check=True)


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


EXT = re.compile(r"^=\[(\d+)\]([A-Za-z]+)!\$?([A-Z]+)\$?(\d+)$")


def aktualisieren_wie_excel(datei):
    """Ersetzt Verknüpfungsformeln =[n]Blatt!Zelle durch die gespeicherten Werte der Quelldatei n."""
    # Excel behält die Verknüpfungen; der Test ersetzt sie durch Werte. Damit man eine Datei mehrmals „aktualisieren“
    # kann, wird beim ersten Mal eine Kopie mit Verknüpfungen aufgehoben und danach immer von ihr ausgegangen.
    kopie = os.path.join(TMP, 'mit-verknuepfung', datei)
    if not os.path.exists(kopie):
        os.makedirs(os.path.dirname(kopie), exist_ok=True); shutil.copy(pfad(datei), kopie)
    wb = openpyxl.load_workbook(kopie)
    from urllib.parse import unquote
    basis = os.path.dirname(pfad(datei))
    ziele = [os.path.normpath(os.path.join(basis, unquote(lk.file_link.Target))) for lk in wb._external_links]
    cache = {i + 1: openpyxl.load_workbook(z, data_only=True) for i, z in enumerate(ziele)}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    m = EXT.match(c.value)
                    if m:
                        v = cache[int(m.group(1))][m.group(2)][f'{m.group(3)}{m.group(4)}'].value
                        c.value = 0 if v is None else v      # leere Zelle → 0 (wie Excel)
    wb._external_links = []
    wb.save(pfad(datei))


def verknuepfung_fehler(datei, ziel):
    """Kopie von datei, in der jede Verknüpfung einen Fehler liefert (wie Excel mit #BEZUG!, wenn die Quelle
    noch nie gespeichert wurde oder nicht erreichbar ist)."""
    wb = openpyxl.load_workbook(pfad(datei))
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and EXT.match(c.value): c.value = '=1/0'
    wb._external_links = []
    wb.save(pfad(ziel))


# 1b. Verknüpfung liefert Fehler (Beobachtung des Nutzers: ganze Dateneingabe #BEZUG!) → Dateien rechnen trotzdem
print('--- Verknüpfung liefert Fehler')
verknuepfung_fehler(f'{OFFEN}/{MA[0]}.xlsx', 'fehler-eingabe.xlsx')
d = oeffnen('fehler-eingabe.xlsx'); d.calculateAll()
pruefe('Eingabe: Start vorbelegt', wert(d, 'Dateneingabe', 'D7'), '11:00')
pruefe('Eingabe: Pause vorbelegt', wert(d, 'Dateneingabe', 'E7'), '30 Min')
pruefe('Eingabe: Gilt Start', wert(d, 'Dateneingabe', 'G7'), '11:00')
pruefe('Eingabe: Hinweis statt Fehler', wert(d, 'Dateneingabe', 'H3'), 'noch keine Vorgabe vom Steuerer')
pruefe('Eingabe: Titel mit Jahr', wert(d, 'Dateneingabe', 'A1')[:17], 'Dateneingabe 2027')
alle = d.Sheets.getByName('Dateneingabe').getCellRangeByName('A1:N267').getDataArray()
pruefe('Eingabe: keine Fehlerwerte', sum(1 for z in alle for v in z if isinstance(v, str) and (v.startswith('#') or v.startswith('Err:'))), 0)
d.close(True)
verknuepfung_fehler(f'{PRIVAT}/Stundenzettel 01.xlsx', 'fehler-sz.xlsx')
d = oeffnen('fehler-sz.xlsx'); d.calculateAll()
pruefe('Stundenzettel: Name leer statt Fehler', wert(d, 'Stammdaten', 'B3'), '')
pruefe('Stundenzettel: Tagesliste leer statt Fehler', wert(d, 'Tagesliste', 'D7') + wert(d, 'Tagesliste', 'E7') + wert(d, 'Tagesliste', 'F7'), '')
d.close(True)
for datei, blatt, zelle, soll_ in ((f'{OFFEN}/Schichtplan.xlsx', 'Mitarbeiter', 'C4', ''),
                                    (f'{PRIVAT}/Inhaltsverzeichnis.xlsx', 'Inhaltsverzeichnis', 'C4',
                                     '– noch kein Name (Stundenzettel aktualisieren und speichern) –')):
    verknuepfung_fehler(datei, 'fehler-x.xlsx')
    d = oeffnen('fehler-x.xlsx'); d.calculateAll()
    pruefe(f'{os.path.basename(datei)}: kein Fehler', wert(d, blatt, zelle), soll_)
    d.close(True)

# 2. Eingabe-Dateien „in Excel“ speichern (ZELLE("Dateiname") wie in Excel fest eingesetzt)
SZN = {n: f'Stundenzettel {i + 1:02d}.xlsx' for i, n in enumerate(MA)}
def eingabe_speichern(n, setze_namen=True):
    d = oeffnen(f'{OFFEN}/{n}.xlsx')
    if setze_namen:
        d.Sheets.getByName('Listen').getCellRangeByName('O2').setString('C:\\Abteilung\\[' + n + '.xlsx]Dateneingabe')
    d.calculateAll(); speichern(d, f'{OFFEN}/{n}.xlsx'); return d
for n in MA:
    d = eingabe_speichern(n)
    pruefe(f'{n}: Name im Export', wert(d, 'Export', 'B1'), n)
    if n == MA[0]:
        pruefe('Eingabe: Blätter (ohne Start/Stammdaten/Auswertung)', ','.join(d.Sheets.ElementNames),
               'Musterwoche,Dateneingabe,Urlaub,Listen,Export,Vorgaben')
        pruefe('Eingabe: Titel Dateneingabe', wert(d, 'Dateneingabe', 'A1'), 'Dateneingabe 2027 – Beispiel, Anna')
        pruefe('Eingabe: Titel Musterwoche', wert(d, 'Musterwoche', 'A1'), 'Musterwoche (Montag bis Freitag) – Beispiel, Anna')
        pruefe('Eingabe: Titel Urlaub', wert(d, 'Urlaub', 'A1'), 'Urlaubsübersicht – Beispiel, Anna')
        pruefe('Eingabe ohne Vorgaben: Reststunden-Hinweis', wert(d, 'Dateneingabe', 'H3'), 'noch keine Vorgabe vom Steuerer')
        pruefe('Eingabe ohne Vorgaben: keine Urlaubswarnung', wert(d, 'Dateneingabe', 'A4'), '')
        pruefe('Eingabe ohne Vorgaben: Vertrag angenommen', wert(d, 'Dateneingabe', 'G7'), '11:00')
    d.close(True)

# 3. Schichtplan direkt aus den Eingabe-Dateien (Stundenzettel noch nie gespeichert)
shutil.copy(pfad(f'{OFFEN}/Schichtplan.xlsx'), os.path.join(TMP, 'Schichtplan-original.xlsx'))
aktualisieren_wie_excel(f'{OFFEN}/Schichtplan.xlsx')
sp = oeffnen(f'{OFFEN}/Schichtplan.xlsx'); sp.calculateAll()
W = '04.01.2027'
print('--- Schichtplan, Woche', W, '(direkt aus den Eingabe-Dateien)')
pruefe('Schichtplan ohne Anleitung', 'Anleitung' in sp.Sheets.ElementNames, False)
pruefe('Zeile 1 Name', wert(sp, W, 'A4'), 'Beispiel, Anna')
pruefe('Zeile 2 Name', wert(sp, W, 'A5'), 'Test, Ben')
pruefe('Platzhalter ausgeblendet', wert(sp, W, 'A6'), '')
pruefe('Platzhalter Mo leer', wert(sp, W, 'B6'), '')
pruefe('Anna Mo Beginn', wert(sp, W, 'B4'), '11:00')
pruefe('Anna Mo Pause', wert(sp, W, 'C4'), '30')
pruefe('Anna Mo Anmerkung (Musterwoche)', wert(sp, W, 'D4'), 'Homeoffice')
pruefe('Anna Di Urlaub', wert(sp, W, 'E4'), 'Urlaub')
pruefe('Anna Di Pause leer', wert(sp, W, 'F4'), '')
pruefe('Anna Di Anmerkung leer', wert(sp, W, 'G4'), '')
pruefe('Anna Fr Beginn', wert(sp, W, 'N4'), '12:00')
pruefe('Ben Mo Beginn', wert(sp, W, 'B5'), '13:00')
pruefe('Ben Mo Pause', wert(sp, W, 'C5'), '45')
pruefe('Ben Mi abwesend (Gutschrift) → abwesend', wert(sp, W, 'H5'), 'abwesend')
pruefe('Ben Mi keine Anmerkung (kein Grund)', wert(sp, W, 'J5'), '')
pruefe('Ben Do Teammentor*in', wert(sp, W, 'M5'), 'Teammentor*in')
pruefe('Ben Fr frei', wert(sp, W, 'N5'), 'Frei')
pruefe('Datum Kopf', wert(sp, W, 'B2')[:6], 'Montag')
if os.environ.get('BILDER'):
    sp.storeToURL(uno.systemPathToFileUrl(os.path.join(os.environ['BILDER'], 'schichtplan.pdf')), (pv('FilterName', 'calc_pdf_Export'),))
sp.close(True)

# 4. Tagesplan direkt aus den Eingabe-Dateien (unabhängig vom Schichtplan)
aktualisieren_wie_excel(f'{OFFEN}/MA-Tagesplan.xlsx')
tp = oeffnen(f'{OFFEN}/MA-Tagesplan.xlsx'); tp.calculateAll()
print('--- MA-Tagesplan (direkt aus den Eingabe-Dateien)')
T = lambda d, z: wert(tp, d, z)
pruefe('04.01. Zeile 1 Vorname', T('04.01.2027', 'B3'), 'Anna')
pruefe('04.01. Zeile 1 Nachname', T('04.01.2027', 'C3'), 'Beispiel')
pruefe('04.01. Zeile 2 Nachname', T('04.01.2027', 'C4'), 'Test')
pruefe('04.01. Anna Beginn', T('04.01.2027', 'D3'), '11:00')
pruefe('04.01. Anna HO', T('04.01.2027', 'F3'), 'HO')
pruefe('04.01. Ben Beginn', T('04.01.2027', 'D4'), '13:00')
pruefe('04.01. Ben Pause', T('04.01.2027', 'E4'), '45')
pruefe('04.01. Platzhalter leer', T('04.01.2027', 'C5'), '')
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

# 5. Stundenzettel aktualisieren wie Excel
print('--- Stundenzettel')
for n in MA:
    aktualisieren_wie_excel(f'{PRIVAT}/{SZN[n]}')
sz = oeffnen(f'{PRIVAT}/{SZN[MA[1]]}'); sz.calculateAll()
pruefe('Name aus der Eingabe-Datei', wert(sz, 'Stammdaten', 'B3'), 'Test, Ben')
pruefe('Stundenzettel ohne Anleitung', 'Anleitung' in sz.Sheets.ElementNames, False)
for blatt, titel in (('Stammdaten', 'Stammdaten'), ('Tagesliste', 'Tagesliste 2027'), ('Urlaub', 'Urlaubsübersicht & Abgleich'),
                     ('Auswertung', 'Auswertung & Jahresbilanz 2027')):
    pruefe(f'Stundenzettel Titel {blatt}', wert(sz, blatt, 'A1'), titel + ' – Test, Ben')
pruefe('Tagesliste 04.01. Beginn', wert(sz, 'Tagesliste', 'G7'), '13:00')
pruefe('Tagesliste 06.01. abwesend, Zeit bleibt', wert(sz, 'Tagesliste', 'G9') + ' / ' + wert(sz, 'Tagesliste', 'I9'), '13:00 / abwesend (Gutschrift)')
pruefe('Tagesliste 08.01. Frei', wert(sz, 'Tagesliste', 'G11'), 'Frei')
for z, v in soll.items():
    pruefe(f'wie Hauptvorlage {z}', wert(sz, *z.split('!')), v)
# Stundenabrechnung 057 für Januar eintragen → muss auch in den Reststunden der Eingabe-Datei ankommen
sz.Sheets.getByName('Auswertung').getCellRangeByName('C5').setValue(100 / 24); sz.calculateAll()
ben_rest = wert(sz, 'Auswertung', 'C28'); ben_rest_sek = sz.Sheets.getByName('Auswertung').getCellRangeByName('G28').getValue()
pruefe('Vorgaben: Jahr / Anspruch', wert(sz, 'Vorgaben', 'A1') + ' / ' + wert(sz, 'Vorgaben', 'B1'), '2027 / 30')
speichern(sz, f'{PRIVAT}/{SZN[MA[1]]}'); sz.close(True)
d = oeffnen(f'{PRIVAT}/{SZN[MA[0]]}')
d.Sheets.getByName('Stammdaten').getCellRangeByName('B5').setValue(0.25); d.calculateAll()    # Anna: Anspruch 0,25 Tage → Warnung
speichern(d, f'{PRIVAT}/{SZN[MA[0]]}'); d.close(True)
d = oeffnen(f'{PRIVAT}/{SZN[MA[2]]}'); d.calculateAll(); speichern(d, f'{PRIVAT}/{SZN[MA[2]]}'); d.close(True)

# 5b. Inhaltsverzeichnis (Ordner Steuerung): Nr. → Name aus den gespeicherten Stundenzetteln
print('--- Inhaltsverzeichnis')
iv_links = openpyxl.load_workbook(pfad(f'{PRIVAT}/Inhaltsverzeichnis.xlsx'))['Inhaltsverzeichnis']
pruefe('Inhaltsverzeichnis: Klick-Link', iv_links['B5'].hyperlink.target, 'Stundenzettel 02.xlsx')
aktualisieren_wie_excel(f'{PRIVAT}/Inhaltsverzeichnis.xlsx')
iv = oeffnen(f'{PRIVAT}/Inhaltsverzeichnis.xlsx'); iv.calculateAll()
for i, n in enumerate(MA):
    pruefe(f'Inhaltsverzeichnis Zeile {i + 1}', wert(iv, 'Inhaltsverzeichnis', f'B{4 + i}') + ' → ' + wert(iv, 'Inhaltsverzeichnis', f'C{4 + i}'),
           f'Stundenzettel {i + 1:02d} → {n}')
iv.close(True)

# 6. Eingabe-Dateien holen die Vorgaben aus dem Stundenzettel (Steuerer öffnet, aktualisiert, speichert)
print('--- Eingabe-Datei mit Vorgaben')
for n in MA[:2]:
    aktualisieren_wie_excel(f'{OFFEN}/{n}.xlsx')
    eingabe_speichern(n).close(True)
d = oeffnen(f'{OFFEN}/{MA[1]}.xlsx'); d.calculateAll()
h3 = wert(d, 'Dateneingabe', 'H3')
pruefe('Ben: Reststunden-Text', h3.split(':')[0], 'noch zu leisten' if ben_rest_sek > 0 else 'zu viel geleistet')
pruefe('Ben: Reststunden = Stundenzettel (Sek.)', round(d.Sheets.getByName('Listen').getCellRangeByName('O9').getValue()), round(ben_rest_sek))
pruefe('Ben: Urlaub Anspruch', wert(d, 'Urlaub', 'B4'), '30 Tage')
if os.environ.get('BILDER'):
    d.storeToURL(uno.systemPathToFileUrl(os.path.join(os.environ['BILDER'], 'eingabe.pdf')), (pv('FilterName', 'calc_pdf_Export'),))
d.close(True)
d = oeffnen(f'{OFFEN}/{MA[0]}.xlsx'); d.calculateAll()
pruefe('Anna: Urlaubswarnung (Anspruch 0,25)', wert(d, 'Dateneingabe', 'A4').replace(',', '.'), 'Achtung: 1 Urlaubstage eingetragen – Anspruch: 0.25 Tage!')
d.close(True)

# 7. Vertragsende: Ben im Januar ohne Vertrag → Eingabe-Datei „kein Vertrag“ → Schichtplan ohne Zeiten
print('--- Vertragsende')
sz = oeffnen(f'{PRIVAT}/{SZN[MA[1]]}')
sz.Sheets.getByName('Stammdaten').getCellRangeByName('B14').setValue(2028 / 24)
sz.Sheets.getByName('Stammdaten').getCellRangeByName('B13').setValue(0); sz.calculateAll()
pruefe('Vorgabe Januar ohne Vertrag = 1', wert(sz, 'Vorgaben', 'A3'), '1')
speichern(sz, f'{PRIVAT}/{SZN[MA[1]]}'); sz.close(True)
aktualisieren_wie_excel(f'{OFFEN}/{MA[1]}.xlsx'); d = eingabe_speichern(MA[1])
pruefe('Eingabe Ben 04.01. „kein Vertrag“', wert(d, 'Dateneingabe', 'G7'), 'kein Vertrag'); d.close(True)
shutil.copy(os.path.join(TMP, 'Schichtplan-original.xlsx'), pfad(f'{OFFEN}/Schichtplan.xlsx'))
aktualisieren_wie_excel(f'{OFFEN}/Schichtplan.xlsx')
sp = oeffnen(f'{OFFEN}/Schichtplan.xlsx'); sp.calculateAll()
pruefe('Ben steht noch in der Liste', wert(sp, W, 'A5'), 'Test, Ben')
pruefe('Ben Mo ohne Zeit (kein Vertrag)', wert(sp, W, 'B5'), '')
pruefe('Ben Do ohne Anmerkung (kein Vertrag)', wert(sp, W, 'M5'), '')
pruefe('Anna unverändert', wert(sp, W, 'B4'), '11:00')
sp.close(True)
try:
    desk.terminate()
except Exception:
    pass
proc.wait(timeout=30)
print('Testordner:', ORDNER)
print('ERGEBNIS: alles in Ordnung' if fehler == 0 else f'ERGEBNIS: {fehler} Fehler')
sys.exit(1 if fehler else 0)
