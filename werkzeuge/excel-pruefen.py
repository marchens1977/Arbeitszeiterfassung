#!/usr/bin/env python3
"""Rechnet die Excel-Datei mit LibreOffice neu und gibt die wichtigsten Werte aus.

Aufruf:  python3 werkzeuge/excel-pruefen.py [datei.xlsx] [--jahresbasis WERT]
  WERT wird in Stammdaten!B13 eingetragen (so, wie die Datei es erwartet), um eine
  andere Eingabe zu testen. Ausgabe: Stammdaten (Soll), Regel-Werte, Urlaub, Auswertung.
Zum Vergleich zweier Versionen beide Dateien prüfen und die Ausgaben nebeneinanderlegen.
"""
import os, subprocess, sys, tempfile, shutil
import openpyxl

args = sys.argv[1:]
datei = 'excel/Arbeitszeiterfassung.xlsx'
setze = None
if '--jahresbasis' in args:
    i = args.index('--jahresbasis'); setze = float(args[i + 1]); del args[i:i + 2]
if args: datei = args[0]

tmp = tempfile.mkdtemp()
quelle = os.path.join(tmp, 'ein.xlsx')
if setze is not None:
    wb = openpyxl.load_workbook(datei)
    wb['Stammdaten']['B13'] = setze
    wb.save(quelle)
else:
    shutil.copy(datei, quelle)
aus = os.path.join(tmp, 'aus')
subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', aus, quelle],
               check=True, capture_output=True)
wb = openpyxl.load_workbook(os.path.join(aus, 'ein.xlsx'), data_only=True)
st, ur, au = wb['Stammdaten'], wb['Urlaub'], wb['Auswertung']
print('Stammdaten A28:', st['A28'].value)
for r in range(27, 33):
    print(f'  {st.cell(r,1).value}: {st.cell(r,2).value}  (Sek {st.cell(r,7).value})')
for r in (13, 14, 24):
    print(f'  {st.cell(r,1).value}: Sek/Jahr {st.cell(r,7).value}, Regel-Tag Sek {st.cell(r,8).value}, '
          f'Regel-Woche {st.cell(r,5).value}, Regel-Monat {st.cell(r,4).value}')
print('Urlaub:', [ur[c].value for c in ('B4', 'B5', 'B6', 'B24', 'B25', 'B26')])
print('Auswertung A21:', au['A21'].value)
for r in range(5, 29):
    werte = [au.cell(r, c).value for c in range(1, 8)]
    if any(v is not None for v in werte): print('  ', werte)
shutil.rmtree(tmp)
