#!/usr/bin/env python3
"""Rechnet die Excel-Datei mit LibreOffice neu und gibt die wichtigsten Werte aus.

Aufruf:  python3 werkzeuge/excel-pruefen.py [datei.xlsx] [--jahresbasis WERT] [--anspruch TAGE]
                                          [--ab-monat N] [--urlaub K]
  --jahresbasis WERT  wird in Stammdaten!B13 eingetragen (so, wie die Datei es erwartet)
  --anspruch TAGE     Urlaubsanspruch volles Jahr (Stammdaten!B5)
  --ab-monat N        Ausscheiden: Jahres-Basis ab Monat N (1-12) auf 0
  --urlaub K          K Urlaubstage im Januar eintragen (ab 4.1., Dateneingabe Spalte D) Ausgabe: Stammdaten (Soll), Regel-Werte, Urlaub, Auswertung.
Zum Vergleich zweier Versionen beide Dateien prüfen und die Ausgaben nebeneinanderlegen.
"""
import os, subprocess, sys, tempfile, shutil
import openpyxl

args = sys.argv[1:]
datei = 'excel/Arbeitszeiterfassung.xlsx'
opt = {}
for name in ('--jahresbasis', '--anspruch', '--ab-monat', '--urlaub'):
    if name in args:
        i = args.index(name); opt[name] = float(args[i + 1]); del args[i:i + 2]
if args: datei = args[0]
setze = opt or None

tmp = tempfile.mkdtemp()
quelle = os.path.join(tmp, 'ein.xlsx')
if setze is not None:
    wb = openpyxl.load_workbook(datei)
    st0 = wb['Stammdaten']
    if '--jahresbasis' in opt: st0['B13'] = opt['--jahresbasis']
    if '--anspruch' in opt: st0['B5'] = opt['--anspruch']
    if '--ab-monat' in opt:
        for m in range(int(opt['--ab-monat']), 13): st0.cell(12 + m, 2).value = 0
    if '--urlaub' in opt:
        for r in range(7, 7 + int(opt['--urlaub'])): wb['Dateneingabe'][f'D{r}'] = 'Urlaub'
    wb.save(quelle)
else:
    shutil.copy(datei, quelle)
aus = os.path.join(tmp, 'aus')
subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', aus, quelle],
               check=True, capture_output=True)
wb = openpyxl.load_workbook(os.path.join(aus, 'ein.xlsx'), data_only=True)
st, ur, au = wb['Stammdaten'], wb['Urlaub'], wb['Auswertung']
print('Mitarbeiter (aus Dateiname, hier immer „ein“):', st['B3'].value, '|', st['C3'].value)
print('Stammdaten C5:', st['C5'].value, '| Hilfswerte G35-G39:', [st[f'G{r}'].value for r in range(35, 40)])
de = wb['Dateneingabe']
print('Dateneingabe Juli (erste Zeile):', next(([de.cell(r,c).value for c in (1,7,11)] for r in range(6,268) if de.cell(r,12).value == 7), None))
print('Urlaub C6:', ur['C6'].value)
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
