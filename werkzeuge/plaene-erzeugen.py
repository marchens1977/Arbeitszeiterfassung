#!/usr/bin/env python3
"""Erzeugt Schichtplan.xlsx und MA-Tagesplan.xlsx (ohne Makros, nur Excel-Verknüpfungen)
und die Mitarbeiter-Dateien (Kopien von excel/Arbeitszeiterfassung.xlsx) in EINEM Ordner.

Datenfluss (Aktualisieren in Excel: Daten → Links bearbeiten → Werte aktualisieren):
  „Nachname, Vorname.xlsx“ (Blatt Export)  →  Schichtplan.xlsx (Blatt Daten)  →  MA-Tagesplan.xlsx

Aufruf:
  python3 werkzeuge/plaene-erzeugen.py --ordner AUSGABE [--plaetze 35] [--wochen 52] [--tage alle]
         [--mitarbeiter "Beispiel, Anna" "Test, Ben" ...] [--beispieldaten]
  --plaetze N       Anzahl Mitarbeiter-Plätze (fehlende Namen werden „Mustermann, Max NN“)
  --wochen N        Anzahl Wochenblätter ab 04.01.2027 (Standard: alle 52)
  --tage N          Anzahl Tagesblätter (Werktage ohne Feiertage) ab 04.01.2027 (Standard: alle)
  --beispieldaten   trägt in die ersten Mitarbeiter-Dateien Testwerte ein (nur für das Mini-Paket)
"""
import argparse, datetime as dt, os, shutil
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.external_link.external import ExternalLink, ExternalBook, ExternalSheetNames, ExternalSheetDataSet, ExternalSheetData
from openpyxl.packaging.relationship import Relationship
from openpyxl.utils import get_column_letter as L
from urllib.parse import quote

JAHR = 2027
ZEILEN = 262            # Werktage-Zeilen im Export-Blatt (Zeile 4 … 265)
ERSTE, LETZTE = 4, 3 + ZEILEN
BLAU, ORANGE, GELB = 'FF002D72', 'FFFF5910', 'FFFFF2CC'
DATUM_LANG = '[$-407]dddd, d. mmmm yyyy'
VORLAGE = os.path.join(os.path.dirname(__file__), '..', 'excel', 'Arbeitszeiterfassung.xlsx')

duenn = Side(style='thin', color='FFBFBFBF')
rahmen = Border(left=duenn, right=duenn, top=duenn, bottom=duenn)
kopf_fill = PatternFill('solid', fgColor=BLAU)
kopf_font = Font(bold=True, color='FFFFFFFF')
mitte = Alignment(horizontal='center', vertical='center', wrap_text=True)


def ostern(j):
    a, b, c = j % 19, j // 100, j % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25; g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mon = (h + l - 7 * m + 114) // 31; tag = (h + l - 7 * m + 114) % 31 + 1
    return dt.date(j, mon, tag)


def feiertage_hessen(j):
    o = ostern(j); t = dt.timedelta
    return {dt.date(j, 1, 1), o - t(2), o + t(1), dt.date(j, 5, 1), o + t(39), o + t(50), o + t(60),
            dt.date(j, 10, 3), dt.date(j, 12, 25), dt.date(j, 12, 26)}


def werktage(j):
    d, out = dt.date(j, 1, 1), []
    while d.year == j:
        if d.weekday() < 5: out.append(d)
        d += dt.timedelta(days=1)
    return out


def externer_link(wb, dateiname, blaetter):
    """Fügt eine Verknüpfung auf eine Datei im selben Ordner hinzu; Rückgabe: Nummer für Formeln [n]Blatt!A1."""
    link = ExternalLink(externalBook=ExternalBook(sheetNames=ExternalSheetNames(sheetName=blaetter), id='rId1',
        sheetDataSet=ExternalSheetDataSet(sheetData=[ExternalSheetData(sheetId=i) for i in range(len(blaetter))])))
    link.file_link = Relationship(type='externalLinkPath', Target=quote(dateiname), TargetMode='External', Id='rId1')
    wb._external_links.append(link)
    return len(wb._external_links)


def kopf(c, text=None):
    if text is not None: c.value = text
    c.fill = kopf_fill; c.font = kopf_font; c.alignment = mitte; c.border = rahmen


def regel(ws, bereich, formel, fill=None, font=None):
    dx = DifferentialStyle(fill=PatternFill(bgColor=fill, fgColor=fill, fill_type='solid') if fill else None,
                           font=font)
    ws.conditional_formatting.add(bereich, Rule(type='expression', dxf=dx, formula=[formel], stopIfTrue=True))


# ---------------------------------------------------------------- Mitarbeiter-Dateien
def mitarbeiter_dateien(ordner, namen, beispieldaten):
    for i, n in enumerate(namen):
        ziel = os.path.join(ordner, n + '.xlsx')
        shutil.copy(VORLAGE, ziel)
        if beispieldaten and i < 2:
            wb = openpyxl.load_workbook(ziel)
            m, d = wb['Musterwoche'], wb['Dateneingabe']
            if i == 0:   # Anna: montags Homeoffice, 05.01. Urlaub, Fr 12:00
                m['D5'] = 'Homeoffice'; d['D8'] = 'Urlaub'; m['B9'] = dt.time(12, 0)
            else:        # Ben: 06.01. krank, 07.01. Teammentor*in, 08.01. frei, Beginn 13:00 mit 45 Min
                for r in range(5, 10): m[f'B{r}'] = dt.time(13, 0); m[f'C{r}'] = 45
                d['F9'] = 'krank'; d['F10'] = 'Teammentor*in'; d['D11'] = 'Frei'
            wb.save(ziel)


# ---------------------------------------------------------------- Schichtplan
def schichtplan(ordner, namen, n_wochen):
    N = len(namen)
    wb = openpyxl.Workbook()
    an = wb.active; an.title = 'Anleitung'
    ma = wb.create_sheet('Mitarbeiter')
    da = wb.create_sheet('Daten')

    # --- Daten (versteckt): Verknüpfungen auf das Blatt Export jeder Mitarbeiter-Datei
    da['A1'] = 'Name'; da['A2'] = 'aktiv'; da['A3'] = 'Datum'
    for k in range(1, ZEILEN + 1):
        r = 3 + k
        da[f'A{r}'] = f'=WORKDAY(DATE({JAHR},1,1)-1,{k})'; da[f'A{r}'].number_format = 'DD.MM.YYYY'
    for s, n in enumerate(namen, start=1):
        nr = externer_link(wb, n + '.xlsx', ['Export'])
        c0 = 2 + (s - 1) * 3
        da.cell(1, c0).value = f'=[{nr}]Export!$B$1'
        da.cell(2, c0).value = f'=[{nr}]Export!$D$1'
        for j, h in enumerate(('Beginn', 'Pause', 'Anm')): da.cell(3, c0 + j).value = f'{h} {s:02d}'
        for r in range(ERSTE, LETZTE + 1):
            da.cell(r, c0).value = f'=[{nr}]Export!B{r}'
            da.cell(r, c0 + 1).value = f'=[{nr}]Export!C{r}'
            da.cell(r, c0 + 2).value = f'=[{nr}]Export!D{r}'
            da.cell(r, c0).number_format = 'hh:mm'
    letzte_sp = L(1 + 3 * N)
    wb.defined_names['SPDatum'] = DefinedName('SPDatum', attr_text=f'Daten!$A${ERSTE}:$A${LETZTE}')
    wb.defined_names['SPWerte'] = DefinedName('SPWerte', attr_text=f'Daten!$B${ERSTE}:${letzte_sp}${LETZTE}')
    da.sheet_state = 'hidden'

    # --- Mitarbeiter: Liste aller Plätze + Sortierung nach Nachname
    ma['A1'] = 'Mitarbeiter im Schichtplan'; ma['A1'].font = Font(bold=True, size=14, color=BLAU)
    ma['A2'] = ('Jede Zeile ist mit einer Datei im selben Ordner verknüpft. Neuer Mitarbeiter: Datei umbenennen '
                '(„Nachname, Vorname.xlsx“), dann hier Daten → Links bearbeiten → alte Datei markieren → „Quelle ändern“ '
                '→ neue Datei wählen. Platzhalter „Mustermann, Max …“ werden nicht angezeigt.')
    ma['A2'].alignment = Alignment(wrap_text=True, vertical='top'); ma.merge_cells('A2:L2'); ma.row_dimensions[2].height = 48
    for i, h in enumerate(['Nr.', 'Datei (bei Erstellung)', 'Name (aus der Datei)', 'aktiv', 'Sortierschlüssel', 'Rang',
                           '', 'Position', 'Platz-Nr.', 'Name sortiert', 'Vorname', 'Nachname']):
        if h: kopf(ma.cell(3, 1 + i), h)
    for s in range(1, N + 1):
        r = 3 + s
        ma[f'A{r}'] = s; ma[f'B{r}'] = namen[s - 1] + '.xlsx'
        c0 = L(2 + (s - 1) * 3)
        ma[f'C{r}'] = f'=IF(ISNUMBER(Daten!{c0}$1),"",Daten!{c0}$1)'
        ma[f'D{r}'] = f'=IF(N(Daten!{c0}$2)=1,1,0)'
        ma[f'E{r}'] = f'=IF(D{r}=1,LOWER(C{r}),"zzzz"&TEXT(A{r},"00"))'
        ma[f'F{r}'] = f'=COUNTIF($E$4:$E${3 + N},"<"&E{r})+COUNTIF($E$4:E{r},E{r})'
        ma[f'H{r}'] = s
        ma[f'I{r}'] = f'=MATCH(H{r},$F$4:$F${3 + N},0)'
        ma[f'J{r}'] = f'=IF(INDEX($D$4:$D${3 + N},I{r})=1,INDEX($C$4:$C${3 + N},I{r}),"")'
        ma[f'K{r}'] = f'=IF(J{r}="","",IFERROR(TRIM(MID(J{r},FIND(",",J{r})+1,99)),""))'
        ma[f'L{r}'] = f'=IF(J{r}="","",IFERROR(TRIM(LEFT(J{r},FIND(",",J{r})-1)),J{r}))'
        for c in 'ABCDEFHIJKL': ma[f'{c}{r}'].border = rahmen
    for c, w in zip('ABCDEFGHIJKL', (5, 26, 26, 6, 18, 6, 2, 8, 8, 26, 16, 18)): ma.column_dimensions[c].width = w
    for c in 'EF': ma.column_dimensions[c].hidden = True
    ma.freeze_panes = 'A4'

    # --- Wochenblätter
    montag = dt.date(JAHR, 1, 4)
    for w in range(n_wochen):
        mo = montag + dt.timedelta(weeks=w)
        ws = wb.create_sheet(mo.strftime('%d.%m.%Y'))
        ws['A1'] = f'Schichtplan – KW {mo.isocalendar()[1]} ({mo:%d.%m.} – {mo + dt.timedelta(days=4):%d.%m.%Y})'
        ws['A1'].font = Font(bold=True, size=14, color=BLAU); ws.merge_cells('A1:P1')
        ws['A1'].alignment = Alignment(horizontal='center')
        kopf(ws['A2'], 'Mitarbeiter'); kopf(ws['A3'], 'Nachname, Vorname')
        for t in range(5):
            b = 2 + 3 * t
            c = ws.cell(2, b); c.value = mo + dt.timedelta(days=t); kopf(c); c.number_format = DATUM_LANG
            ws.merge_cells(start_row=2, start_column=b, end_row=2, end_column=b + 2)
            for j, h in enumerate(('Arbeits-\nbeginn', 'Pause\n(Min)', 'An-\nmerkung')): kopf(ws.cell(3, b + j), h)
        ws.row_dimensions[3].height = 30
        for p in range(1, N + 1):
            r = 3 + p
            ws[f'Q{r}'] = f'=Mitarbeiter!I{r}'                     # Platz-Nr. (versteckt)
            ws[f'A{r}'] = f'=Mitarbeiter!J{r}'
            ws[f'A{r}'].border = rahmen
            for t in range(5):
                b = 2 + 3 * t; dz = L(b) + '$2'
                zeile = f'MATCH({dz},SPDatum,0)'
                aus = lambda off: f'INDEX(SPWerte,{zeile},($Q{r}-1)*3+{off})'
                ws.cell(r, b).value = f'=IF($A{r}="","",IFERROR(IF({aus(1)}=0,"",{aus(1)}),""))'
                ws.cell(r, b + 1).value = f'=IF(ISNUMBER({L(b)}{r}),IFERROR({aus(2)},""),"")'
                ws.cell(r, b + 2).value = f'=IF($A{r}="","",IFERROR(IF({aus(3)}=0,"",{aus(3)}),""))'
                ws.cell(r, b).number_format = 'hh:mm'
                for j in range(3):
                    ws.cell(r, b + j).border = rahmen; ws.cell(r, b + j).alignment = Alignment(horizontal='center')
        letzte = 3 + N
        for t in range(5):
            b = L(2 + 3 * t); e = L(4 + 3 * t); an_ = L(4 + 3 * t)
            blk = f'{b}4:{e}{letzte}'
            regel(ws, blk, f'OR(${b}4="Urlaub",${b}4="Feiertag")', 'FF92D050')
            regel(ws, blk, f'${b}4="abwesend"', 'FFD9D9D9', Font(color='FF595959', italic=True))
            regel(ws, blk, f'${b}4="Frei"', 'FFF2F2F2', Font(color='FF7F7F7F'))
            regel(ws, blk, f'${an_}4="Homeoffice"', 'FFFFFF00')
            regel(ws, blk, f'${an_}4="Teammentor*in"', 'FFBDD7EE')
        ws.column_dimensions['A'].width = 26
        for c in range(2, 17): ws.column_dimensions[L(c)].width = 9 if (c - 2) % 3 != 2 else 16
        ws.column_dimensions['Q'].hidden = True
        ws.freeze_panes = 'B4'
        ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.protection.sheet = True

    # --- Anleitung
    zeilen = [
        ('Schichtplan – Anleitung', True),
        ('Was ist das?', True),
        ('Der Schichtplan zeigt für jede Woche, wer wann arbeitet. Die Werte kommen automatisch aus den '
         'Arbeitszeiterfassungen der Mitarbeiter („Nachname, Vorname.xlsx“) im selben Ordner.', False),
        ('Aktualisieren (ohne Makros)', True),
        ('1. Datei öffnen. Erscheint oben ein gelber Balken „Sicherheitswarnung … Verknüpfungen“: „Inhalt aktivieren“ klicken.', False),
        ('2. Sonst: Menü Daten → Links bearbeiten (bzw. Verknüpfungen bearbeiten) → alle markieren → „Werte aktualisieren“.', False),
        ('3. Speichern – erst dann bekommt der MA-Tagesplan die neuen Werte.', False),
        ('Wichtig: Es zählen die zuletzt GESPEICHERTEN Werte der Mitarbeiter-Dateien.', False),
        ('Farben', True),
        ('Grün = Urlaub oder Feiertag · Grau = abwesend (Grund wird nicht angezeigt) · Hellgrau = frei · '
         'Gelb = Homeoffice · Hellblau = Teammentor*in', False),
        ('Mitarbeiter wechseln', True),
        ('Datei des Platzhalters bzw. alten Mitarbeiters umbenennen in „Nachname, Vorname.xlsx“, darin Stammdaten anpassen, '
         'speichern. Dann im Schichtplan: Daten → Links bearbeiten → alten Dateinamen markieren → „Quelle ändern“ → '
         'neue Datei wählen → Speichern. Der MA-Tagesplan muss NICHT angepasst werden.', False),
        ('Schutz', True),
        ('Blätter sind ohne Passwort geschützt. Gegen fremde Änderungen: Datei → Speichern unter → Tools → '
         'Allgemeine Optionen → „Kennwort zum Ändern“ (NICHT „Kennwort zum Öffnen“, sonst klappt das Aktualisieren nicht).', False),
    ]
    for i, (t, fett) in enumerate(zeilen, start=1):
        c = an.cell(i, 1, t); c.alignment = Alignment(wrap_text=True, vertical='top')
        c.font = Font(bold=fett, size=14 if i == 1 else 11, color=BLAU if fett else None)
    an.column_dimensions['A'].width = 110
    an.sheet_properties.tabColor = 'FF5910'
    wb.active = wb.sheetnames.index(wb.sheetnames[3]) if n_wochen else 1
    wb.save(os.path.join(ordner, 'Schichtplan.xlsx'))


# ---------------------------------------------------------------- MA-Tagesplan
def tagesplan(ordner, N, n_tage):
    wb = openpyxl.Workbook()
    ue = wb.active; ue.title = 'Übersicht'
    ma = wb.create_sheet('Mitarbeiter'); da = wb.create_sheet('Daten')
    nr = externer_link(wb, 'Schichtplan.xlsx', ['Mitarbeiter', 'Daten'])
    letzte_sp = L(1 + 3 * N)
    # Daten: Spiegel des Schichtplan-Blatts „Daten“
    da['A3'] = 'Datum'
    for r in range(ERSTE, LETZTE + 1):
        da[f'A{r}'] = f'=WORKDAY(DATE({JAHR},1,1)-1,{r - 3})'; da[f'A{r}'].number_format = 'DD.MM.YYYY'
        for c in range(2, 2 + 3 * N):
            da.cell(r, c).value = f'=[{nr}]Daten!{L(c)}{r}'
    wb.defined_names['TPDatum'] = DefinedName('TPDatum', attr_text=f'Daten!$A${ERSTE}:$A${LETZTE}')
    wb.defined_names['TPWerte'] = DefinedName('TPWerte', attr_text=f'Daten!$B${ERSTE}:${letzte_sp}${LETZTE}')
    da.sheet_state = 'hidden'
    # Mitarbeiter: sortierte Liste aus dem Schichtplan (Platz-Nr., Vorname, Nachname)
    for i, h in enumerate(('Position', 'Platz-Nr.', 'Vorname', 'Nachname')): ma.cell(3, 1 + i, h)
    for p in range(1, N + 1):
        r = 3 + p
        ma[f'A{r}'] = p
        ma[f'B{r}'] = f'=[{nr}]Mitarbeiter!I{r}'              # Verknüpfungen nur als einfache Bezüge
        ma[f'E{r}'] = f'=[{nr}]Mitarbeiter!K{r}'
        ma[f'F{r}'] = f'=[{nr}]Mitarbeiter!L{r}'
        ma[f'C{r}'] = f'=IF(ISNUMBER(E{r}),"",E{r})'          # leere Zelle kommt als 0 an
        ma[f'D{r}'] = f'=IF(ISNUMBER(F{r}),"",F{r})'
    ma.sheet_state = 'hidden'

    tage = [d for d in werktage(JAHR) if d not in feiertage_hessen(JAHR) and d >= dt.date(JAHR, 1, 4)]
    if n_tage: tage = tage[:n_tage]
    ue['A1'] = 'MA-Tagesplan – Übersicht'; ue['A1'].font = Font(bold=True, size=14, color=BLAU)
    ue['A2'] = ('Ein Blatt je Arbeitstag (ohne Wochenenden und hessische Feiertage). Namen, Beginn, Pause und Anmerkung kommen '
                'aus dem Schichtplan (dort zuerst aktualisieren und speichern; hier dann Daten → Links bearbeiten → '
                '„Werte aktualisieren“). Gelbe Kampagnen-Felder trägst du selbst ein. Grau = an dem Tag nicht da.')
    ue['A2'].alignment = Alignment(wrap_text=True, vertical='top'); ue.merge_cells('A2:F2'); ue.row_dimensions[2].height = 60
    ue.column_dimensions['A'].width = 34
    for i, d in enumerate(tage):
        name = d.strftime('%d.%m.%Y')
        c = ue.cell(4 + i, 1, d); c.number_format = DATUM_LANG
        c.hyperlink = f"#'{name}'!A1"; c.font = Font(color='FF0563C1', underline='single')

        ws = wb.create_sheet(name)
        ws['A1'] = d; ws['A1'].number_format = DATUM_LANG
        ws['A1'].font = Font(bold=True, size=14, color='FFFFFFFF'); ws['A1'].fill = kopf_fill; ws['A1'].alignment = mitte
        ws.merge_cells('A1:J1'); ws.row_dimensions[1].height = 26
        ws['L1'] = '=MATCH(A1,TPDatum,0)'                     # Zeile im Datenblatt (versteckt)
        for j, h in enumerate(['Nr.', 'MA\n(Vorname)', 'MA\n(Nachname)', 'Arbeits-\nbeginn', 'geplante\nPause in\nMinuten',
                               'Anm.', 'Kampagne 1\n(11:00 - 13:00)', 'Kampagne 2\n(13:00 - 15:00)',
                               'Kampagne 3\n(15:00 - 17:00)', 'Kampagne 4\n(17:00 - 19:00)']):
            kopf(ws.cell(2, 1 + j), h)
        ws.row_dimensions[2].height = 45
        for p in range(1, N + 1):
            r = 2 + p; mr = 3 + p
            ws[f'L{r}'] = f'=Mitarbeiter!B{mr}'               # Platz-Nr.
            ws[f'M{r}'] = f'=IF(C{r}="","",IFERROR(INDEX(TPWerte,$L$1,(L{r}-1)*3+1),""))'   # Beginn/Status roh
            ws[f'N{r}'] = f'=IF(C{r}="","",IFERROR(INDEX(TPWerte,$L$1,(L{r}-1)*3+3),""))'   # Anmerkung roh
            ws[f'A{r}'] = f'{p}.'
            ws[f'B{r}'] = f'=Mitarbeiter!C{mr}'
            ws[f'C{r}'] = f'=Mitarbeiter!D{mr}'
            ws[f'D{r}'] = f'=IF(ISNUMBER(M{r}),IF(M{r}=0,"",M{r}),"")'
            ws[f'E{r}'] = f'=IF(ISNUMBER(M{r}),IFERROR(INDEX(TPWerte,$L$1,(L{r}-1)*3+2),""),"")'
            ws[f'F{r}'] = (f'=IF(C{r}="","",IF(ISNUMBER(M{r}),IF(N{r}="Homeoffice","HO",IF(N{r}=0,"",N{r})),'
                           f'IF(OR(M{r}="",M{r}=0),"",M{r})))')
            ws[f'D{r}'].number_format = 'hh:mm'
            for c in 'ABCDEFGHIJ':
                ws[f'{c}{r}'].border = rahmen
                ws[f'{c}{r}'].alignment = Alignment(horizontal='center' if c in 'ADEF' else 'left', vertical='center')
            for c in 'GHIJ':
                ws[f'{c}{r}'].fill = PatternFill('solid', fgColor=GELB)
                ws[f'{c}{r}'].protection = openpyxl.styles.Protection(locked=False)
            ws.row_dimensions[r].height = 20
        last = 2 + N
        regel(ws, f'A3:J{last}', '$C3=""', 'FFFFFFFF', Font(color='FFBFBFBF'))                 # leerer Platz
        regel(ws, f'A3:J{last}', 'AND($C3<>"",NOT(ISNUMBER($M3)))', 'FFD9D9D9', Font(color='FF7F7F7F', italic=True))
        regel(ws, f'B3:F{last}', '$F3="HO"', 'FFFFFF00')
        regel(ws, f'B3:F{last}', '$F3="Teammentor*in"', 'FFBDD7EE')
        for c, w in zip('ABCDEFGHIJ', (6, 16, 18, 10, 10, 16, 22, 22, 22, 22)): ws.column_dimensions[c].width = w
        for c in 'LMN': ws.column_dimensions[c].hidden = True
        ws.freeze_panes = 'D3'
        ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.protection.sheet = True
    wb.save(os.path.join(ordner, 'MA-Tagesplan.xlsx'))


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--ordner', required=True)
    ap.add_argument('--plaetze', type=int, default=35)
    ap.add_argument('--wochen', type=int, default=52)
    ap.add_argument('--tage', type=int, default=0)
    ap.add_argument('--mitarbeiter', nargs='*', default=[])
    ap.add_argument('--beispieldaten', action='store_true')
    a = ap.parse_args()
    os.makedirs(a.ordner, exist_ok=True)
    namen = list(a.mitarbeiter) + [f'Mustermann, Max {i:02d}' for i in range(len(a.mitarbeiter) + 1, a.plaetze + 1)]
    mitarbeiter_dateien(a.ordner, namen, a.beispieldaten)
    schichtplan(a.ordner, namen, a.wochen)
    tagesplan(a.ordner, len(namen), a.tage)
    print('Erzeugt in', a.ordner, ':', len(namen), 'Mitarbeiter-Dateien, Schichtplan.xlsx, MA-Tagesplan.xlsx')
