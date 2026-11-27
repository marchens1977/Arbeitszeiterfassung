#!/usr/bin/env python3
"""Prüft Schichtplan und MA-Tagesplan (Kette ohne Makros) mit LibreOffice.

Ablauf (bildet Excel nach):
  1. Mini-Paket erzeugen (3 Plätze: „Beispiel, Anna“ und „Test, Ben“ mit Testwerten, 1 Platzhalter)
  2. Mitarbeiter-Dateien öffnen, rechnen, speichern (wie ein Mitarbeiter in Excel)
  3. „Werte aktualisieren“ wie Excel: Excel liest bei geschlossenen Dateien die GESPEICHERTEN Werte.
     Der Test setzt genau diese Werte in den Schichtplan ein, rechnet ihn mit LibreOffice und speichert ihn.
  4. Dasselbe für den MA-Tagesplan mit den gespeicherten Werten des Schichtplans.
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
    wb = openpyxl.load_workbook(pfad(datei))
    ziele = [lk.file_link.Target.replace('%20', ' ').replace('%2C', ',') for lk in wb._external_links]
    cache = {i + 1: openpyxl.load_workbook(pfad(z), data_only=True) for i, z in enumerate(ziele)}
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


# 2. Mitarbeiter-Dateien „in Excel“ speichern (ZELLE("Dateiname") wie in Excel fest eingesetzt)
for n in MA:
    d = oeffnen(n + '.xlsx')
    d.Sheets.getByName('Stammdaten').getCellRangeByName('G3').setString('C:\\Abteilung\\[' + n + '.xlsx]Stammdaten')
    d.calculateAll()
    pruefe(f'{n}: Name im Export', wert(d, 'Export', 'B1'), n)
    speichern(d, n + '.xlsx'); d.close(True)

# 3. Schichtplan aktualisieren wie Excel
aktualisieren_wie_excel('Schichtplan.xlsx')
sp = oeffnen('Schichtplan.xlsx'); sp.calculateAll()
W = '04.01.2027'
print('--- Schichtplan, Woche', W)
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
pruefe('Ben Mi krank → abwesend', wert(sp, W, 'H5'), 'abwesend')
pruefe('Ben Mi keine Anmerkung (kein „krank“)', wert(sp, W, 'J5'), '')
pruefe('Ben Do Teammentor*in', wert(sp, W, 'M5'), 'Teammentor*in')
pruefe('Ben Fr frei', wert(sp, W, 'N5'), 'Frei')
pruefe('Datum Kopf', wert(sp, W, 'B2')[:6], 'Montag')
speichern(sp, 'Schichtplan.xlsx'); sp.close(True)

# 4. Tagesplan aktualisieren wie Excel (Quelle: gespeicherter Schichtplan)
aktualisieren_wie_excel('MA-Tagesplan.xlsx')
tp = oeffnen('MA-Tagesplan.xlsx'); tp.calculateAll()
print('--- MA-Tagesplan')
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
if os.environ.get('BILDER'):
    sp = oeffnen('Schichtplan.xlsx')
    sp.storeToURL(uno.systemPathToFileUrl(os.path.join(os.environ['BILDER'], 'schichtplan.pdf')), (pv('FilterName', 'calc_pdf_Export'),))
    sp.close(True)
try:
    desk.terminate()
except Exception:
    pass
proc.wait(timeout=30)
print('Testordner:', ORDNER)
print('ERGEBNIS: alles in Ordnung' if fehler == 0 else f'ERGEBNIS: {fehler} Fehler')
sys.exit(1 if fehler else 0)
