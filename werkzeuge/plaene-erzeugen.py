#!/usr/bin/env python3
"""Erzeugt das Paket für den Abteilungsbetrieb (ohne Makros, nur Excel-Verknüpfungen) unter AUSGABE:

  AUSGABE/Schichtplanung/  (alle Mitarbeiter)  „Platz 01.xlsx“ (Eingabe-Datei, Vorlage), „001 MA-Tagesplan.xlsx“, „002 Schichtplan.xlsx“
  AUSGABE/Steuerung/       (nur Steuerer)      „Stundenzettel 01.xlsx“ (Vorlage), „000 Inhaltsverzeichnis.xlsx“

VORLAGEN-PRINZIP (Wunsch des Nutzers, 28.11.2026): Der Steuerer bekommt nur je EINE Eingabe-Datei und EINEN Stundenzettel und
kopiert sie selbst („Platz 02.xlsx“ …, „Stundenzettel 02.xlsx“ …). Jede Datei liest ihre Nummer aus dem eigenen Dateinamen.
Weil eine Excel-Verknüpfung fest auf EINE Datei zeigt, verbinden sich die Kopien nur mit den zwei Sammel-Dateien:
  Platz NN (Blatt Export, roh)                  → 002 Schichtplan (Daten)  und  → 001 MA-Tagesplan (Daten)
  000 Inhaltsverzeichnis (Namen, vom Steuerer)  → 002 Schichtplan / 001 MA-Tagesplan (Namen)
  002 Schichtplan (Daten, Mitarbeiter)          → Stundenzettel NN (eigene Spalten per INDEX)
  Stundenzettel NN (Blatt Vorgaben)             → 000 Inhaltsverzeichnis (Blatt Vorgaben, Zeile NN)
  000 Inhaltsverzeichnis (Vorgaben, Name)       → Platz NN (eigene Zeile per INDEX)
Namen stehen NUR im Inhaltsverzeichnis. Niemand wird im laufenden Jahr überschrieben; Neue kommen in die nächste freie Zeile.

Aufruf:
  python3 werkzeuge/plaene-erzeugen.py --ordner AUSGABE [--plaetze 35] [--wochen 52] [--tage alle]
         [--mitarbeiter "Beispiel, Anna" "Test, Ben" ...] [--testnamen] [--kopien] [--beispieldaten]
  --plaetze N       Anzahl Plätze in Schichtplan, Tagesplan, Inhaltsverzeichnis (höchstens 35)
  --wochen N        Anzahl Wochenblätter ab 04.01.2027 (Standard: alle 52)
  --tage N          Anzahl Tagesblätter (Werktage ohne Feiertage) ab 04.01.2027 (Standard: alle)
  --mitarbeiter     Namen für die ersten Zeilen des Inhaltsverzeichnisses (nur Testnamen – echte Namen trägt der Nutzer ein)
  --testnamen       füllt die übrigen Zeilen mit erfundenen Namen „Mustermann, <Vorname>“
  --kopien          legt zusätzlich alle Kopien Platz 02 … / Stundenzettel 02 … an (für Tests; sonst macht das der Nutzer)
  --beispieldaten   trägt in Platz 01 und 02 Testwerte ein (mit --kopien; nur für den Kettentest)
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
VORLAGEN = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'excel', 'vorlagen')
OFFEN, PRIVAT = 'Schichtplanung', 'Steuerung'          # Ordnernamen (nebeneinander, gleiches Laufwerk)
SZ, PLATZ = 'Stundenzettel ', 'Platz '                  # Dateiname-Anfang (+ Nummer 01 … 35)
IV, TP, SP = '000 Inhaltsverzeichnis.xlsx', '001 MA-Tagesplan.xlsx', '002 Schichtplan.xlsx'   # NICHT umbenennen (Ziel von Verknüpfungen)
MAXP = 35                                              # muss zu vorlagen-bauen.py passen
TESTVORNAMEN = ['Alva', 'Bosse', 'Cleo', 'Dario', 'Elin', 'Fiete', 'Greta', 'Hanno', 'Ida', 'Jarne', 'Kaja', 'Linus', 'Mila',
                'Nele', 'Ole', 'Pia', 'Quirin', 'Rosa', 'Sanne', 'Tjark', 'Una', 'Vito', 'Wanda', 'Xaver', 'Yara', 'Zeno',
                'Amira', 'Bela', 'Carla', 'Delia', 'Emil', 'Frida', 'Gero', 'Hedda', 'Ilvy', 'Juna', 'Kian', 'Lotte',
                'Merle', 'Noah']

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


# ---------------------------------------------------------------- Vorlagen (Platz 01, Stundenzettel 01) und Kopien
def vorlagen_dateien(ordner, n_kopien, beispieldaten, testmuster=False):
    """Platz 01.xlsx und Stundenzettel 01.xlsx (mit Verknüpfung auf die Sammel-Dateien). Mit n_kopien > 1 zusätzlich die
    Kopien 02 … n (wie der Nutzer sie im Explorer anlegt: unveränderte Kopie, nur anderer Dateiname)."""
    sz = openpyxl.load_workbook(os.path.join(VORLAGEN, 'Stundenzettel-Vorlage.xlsx'))
    externer_link(sz, f'../{OFFEN}/{SP}', ['Daten', 'Mitarbeiter'])
    sz.save(os.path.join(ordner, PRIVAT, f'{SZ}01.xlsx'))
    wb = openpyxl.load_workbook(os.path.join(VORLAGEN, 'Eingabe.xlsx'))
    externer_link(wb, f'../{PRIVAT}/{IV}', ['Vorgaben'])
    wb.save(os.path.join(ordner, OFFEN, f'{PLATZ}01.xlsx'))
    for i in range(2, n_kopien + 1):
        shutil.copy(os.path.join(ordner, PRIVAT, f'{SZ}01.xlsx'), os.path.join(ordner, PRIVAT, f'{SZ}{i:02d}.xlsx'))
        shutil.copy(os.path.join(ordner, OFFEN, f'{PLATZ}01.xlsx'), os.path.join(ordner, OFFEN, f'{PLATZ}{i:02d}.xlsx'))
    for i in range(n_kopien if (beispieldaten or testmuster) else 0):
        ziel = os.path.join(ordner, OFFEN, f'{PLATZ}{i + 1:02d}.xlsx')
        wb = openpyxl.load_workbook(ziel)
        m, d = wb['Musterwoche'], wb['Dateneingabe']
        if beispieldaten:
            if i == 0:   # Anna: montags Homeoffice, 05.01. Urlaub, Fr 12:00
                m['D5'] = 'Homeoffice'; d['D8'] = 'Urlaub'; m['B9'] = dt.time(12, 0)
            elif i == 1:  # Ben: 06.01. abwesend (Gutschrift), 07.01. Teammentor*in, 08.01. frei, Beginn 13:00 mit 45 Min
                for r in range(5, 10): m[f'B{r}'] = dt.time(13, 0); m[f'C{r}'] = 45
                d['F9'] = 'abwesend (Gutschrift)'; d['F10'] = 'Teammentor*in'; d['D11'] = 'Frei'
            else:
                continue
        else:
            # unterschiedliche Musterwochen: Beginn 11:00–17:00 (15-Min.-Takt), Pause 30/45/60, mal Homeoffice/Frei
            for t in range(5):
                r = 5 + t
                q = (i * 7 + t * 3) % 25                     # 0 … 24 → 11:00 … 17:00
                m[f'B{r}'] = dt.time(11 + q // 4, (q % 4) * 15)
                m[f'C{r}'] = (30, 45, 60)[(i + t) % 3]
                if (i + t) % 4 == 0: m[f'D{r}'] = 'Homeoffice'
                elif (i + t) % 9 == 0: m[f'D{r}'] = 'Teammentor*in'
            if i % 5 == 0: m[f'B{5 + i % 5}'] = 'Frei'      # jede/r Fünfte montags frei
        wb.save(ziel)


# ---------------------------------------------------------------- Quellen
def quellen(wb, N, nach_name=True):
    """Blätter „Mitarbeiter“ (Plätze, Reihenfolge), „Daten“ (versteckt: je Platz eine Verknüpfung auf das Blatt Export der
    Eingabe-Datei „Platz NN.xlsx“, ROHE Gilt-Werte; Zeile 1 Name aus dem Inhaltsverzeichnis) und „Werte“ (versteckt: daraus
    die Anzeige wie früher Export B–D: „abwesend“ statt Grund, nichts bei „kein Vertrag“). Namen SPDatum/SPWerte.
    Der Schichtplan dient zugleich als Sammelquelle für die Stundenzettel (Blätter Daten und Mitarbeiter).
    nach_name: Schichtplan nach Nachname sortiert; Tagesplan (False) fest nach Platz-Nummer (nichts verrutscht bei Neuen)."""
    ma = wb.create_sheet('Mitarbeiter')
    da = wb.create_sheet('Daten')
    we = wb.create_sheet('Werte')
    iv = externer_link(wb, f'../{PRIVAT}/{IV}', ['Inhaltsverzeichnis'])

    da['A1'] = 'Name'; da['A2'] = 'aktiv'; da['A3'] = 'Datum'
    for k in range(1, ZEILEN + 1):
        r = 3 + k
        for w in (da, we):
            w[f'A{r}'] = f'=WORKDAY(DATE({JAHR},1,1)-1,{k})'; w[f'A{r}'].number_format = 'DD.MM.YYYY'
    for s in range(1, N + 1):
        nr = externer_link(wb, f'{PLATZ}{s:02d}.xlsx', ['Export'])
        c0 = 2 + (s - 1) * 3; B, C, D = L(c0), L(c0 + 1), L(c0 + 2)
        da.cell(1, c0).value = f'=[{iv}]Inhaltsverzeichnis!$B${3 + s}'
        da.cell(2, c0).value = f'=IFERROR(IF(OR(ISNUMBER({B}1),TRIM({B}1)=""),0,1),0)'
        for j, h in enumerate(('Start', 'Pause', 'Anm')): da.cell(3, c0 + j).value = f'{h} roh {s:02d}'
        for j, h in enumerate(('Beginn', 'Pause', 'Anm')): we.cell(3, c0 + j).value = f'{h} {s:02d}'
        for r in range(ERSTE, LETZTE + 1):
            da.cell(r, c0).value = f'=[{nr}]Export!E{r}'
            da.cell(r, c0 + 1).value = f'=[{nr}]Export!F{r}'
            da.cell(r, c0 + 2).value = f'=[{nr}]Export!G{r}'
            da.cell(r, c0).number_format = 'hh:mm'
            # wie früher Export B/C/D der Eingabe-Datei
            we.cell(r, c0).value = (f'=IFERROR(IF(Daten!{B}{r}="kein Vertrag","",IF(Daten!{D}{r}="abwesend (Gutschrift)",'
                                    f'"abwesend",Daten!{B}{r})),"")')
            we.cell(r, c0 + 1).value = f'=IF(ISNUMBER({B}{r}),IFERROR(Daten!{C}{r},""),"")'
            we.cell(r, c0 + 2).value = f'=IFERROR(IF(Daten!{D}{r}="abwesend (Gutschrift)","",Daten!{D}{r}),"")'
            we.cell(r, c0).number_format = 'hh:mm'
    letzte_sp = L(1 + 3 * N)
    wb.defined_names['SPDatum'] = DefinedName('SPDatum', attr_text=f'Werte!$A${ERSTE}:$A${LETZTE}')
    wb.defined_names['SPWerte'] = DefinedName('SPWerte', attr_text=f'Werte!$B${ERSTE}:${letzte_sp}${LETZTE}')
    da.sheet_state = 'hidden'; we.sheet_state = 'hidden'

    # --- Mitarbeiter: Liste aller Plätze + Reihenfolge
    ma['A1'] = 'Mitarbeiter (Plätze)'; ma['A1'].font = Font(bold=True, size=14, color=BLAU)
    ma['A2'] = ('Die Namen kommen aus dem Inhaltsverzeichnis des Steuerers (Ordner „Steuerung“), die Zeiten aus den '
                'Eingabe-Dateien „Platz 01.xlsx“ … (gleiche Nummer). Leere Zeile = freier Platz.')
    ma['A2'].alignment = Alignment(wrap_text=True, vertical='top'); ma.merge_cells('A2:L2'); ma.row_dimensions[2].height = 32
    for i, h in enumerate(['Nr.', 'Eingabe-Datei', 'Name', 'aktiv', 'Sortierschlüssel', 'Rang',
                           '', 'Position', 'Platz-Nr.', 'Name sortiert', 'Vorname', 'Nachname']):
        if h: kopf(ma.cell(3, 1 + i), h)
    for s in range(1, N + 1):
        r = 3 + s
        ma[f'A{r}'] = s; ma[f'B{r}'] = f'{PLATZ}{s:02d}.xlsx'
        c0 = L(2 + (s - 1) * 3)
        ma[f'C{r}'] = f'=IFERROR(IF(N(Daten!{c0}$2)=1,TRIM(Daten!{c0}$1),""),"")'
        ma[f'D{r}'] = f'=IFERROR(IF(N(Daten!{c0}$2)=1,1,0),0)'
        if nach_name:
            ma[f'E{r}'] = f'=IF(D{r}=1,LOWER(C{r}),"zzzz"&TEXT(A{r},"00"))'
            ma[f'F{r}'] = f'=COUNTIF($E$4:$E${3 + N},"<"&E{r})+COUNTIF($E$4:E{r},E{r})'
            ma[f'I{r}'] = f'=MATCH(H{r},$F$4:$F${3 + N},0)'
        else:
            ma[f'I{r}'] = s                                   # fest nach Platz-Nummer
        ma[f'H{r}'] = s
        ma[f'J{r}'] = f'=IF(INDEX($D$4:$D${3 + N},I{r})=1,INDEX($C$4:$C${3 + N},I{r}),"")'
        ma[f'K{r}'] = f'=IF(J{r}="","",IFERROR(TRIM(MID(J{r},FIND(",",J{r})+1,99)),""))'
        ma[f'L{r}'] = f'=IF(J{r}="","",IFERROR(TRIM(LEFT(J{r},FIND(",",J{r})-1)),J{r}))'
        for c in 'ABCDEFHIJKL': ma[f'{c}{r}'].border = rahmen
    for c, w in zip('ABCDEFGHIJKL', (5, 16, 26, 6, 18, 6, 2, 8, 8, 26, 16, 18)): ma.column_dimensions[c].width = w
    for c in 'EF': ma.column_dimensions[c].hidden = True
    ma.freeze_panes = 'A4'
    return ma


# ---------------------------------------------------------------- Schichtplan
def schichtplan(ordner, N, n_wochen):
    wb = openpyxl.Workbook()
    an = wb.active; an.title = 'Anleitung'
    quellen(wb, N)

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

    del wb['Anleitung']                     # keine Anleitung – erklärt der Steuerer bei der Einweisung
    wb.active = 3 if n_wochen else 0          # erstes Wochenblatt (Mitarbeiter, Daten, Werte, Wochen …)
    for ws in wb: ws.sheet_view.tabSelected = (ws is wb.active)
    wb.save(os.path.join(ordner, OFFEN, SP))


# ---------------------------------------------------------------- MA-Tagesplan
def tagesplan(ordner, N, n_tage):
    wb = openpyxl.Workbook()
    ue = wb.active; ue.title = 'Übersicht'
    ma = quellen(wb, N, nach_name=False); ma.sheet_state = 'hidden'

    tage = [d for d in werktage(JAHR) if d not in feiertage_hessen(JAHR) and d >= dt.date(JAHR, 1, 4)]
    if n_tage: tage = tage[:n_tage]
    ue['A1'] = 'MA-Tagesplan – Übersicht'; ue['A1'].font = Font(bold=True, size=14, color=BLAU)
    ue.column_dimensions['A'].width = 34
    for i, d in enumerate(tage):
        name = d.strftime('%d.%m.%Y')
        c = ue.cell(4 + i, 1, d); c.number_format = DATUM_LANG
        c.hyperlink = f"#'{name}'!A1"; c.font = Font(color='FF0563C1', underline='single')

        ws = wb.create_sheet(name)
        ws['A1'] = d; ws['A1'].number_format = DATUM_LANG
        ws['A1'].font = Font(bold=True, size=14, color='FFFFFFFF'); ws['A1'].fill = kopf_fill; ws['A1'].alignment = mitte
        ws.merge_cells('A1:J1'); ws.row_dimensions[1].height = 26
        ws['L1'] = '=MATCH(A1,SPDatum,0)'                     # Zeile im Datenblatt (versteckt)
        for j, h in enumerate(['Nr.', 'MA\n(Vorname)', 'MA\n(Nachname)', 'Arbeits-\nbeginn', 'geplante\nPause in\nMinuten',
                               'Anm.', 'Kampagne 1\n(11:00 - 13:00)', 'Kampagne 2\n(13:00 - 15:00)',
                               'Kampagne 3\n(15:00 - 17:00)', 'Kampagne 4\n(17:00 - 19:00)']):
            kopf(ws.cell(2, 1 + j), h)
        ws.row_dimensions[2].height = 45
        for p in range(1, N + 1):
            r = 2 + p; mr = 3 + p
            ws[f'L{r}'] = f'=Mitarbeiter!I{mr}'               # Platz-Nr.
            ws[f'M{r}'] = f'=IF(C{r}="","",IFERROR(INDEX(SPWerte,$L$1,(L{r}-1)*3+1),""))'   # Beginn/Status roh
            ws[f'N{r}'] = f'=IF(C{r}="","",IFERROR(INDEX(SPWerte,$L$1,(L{r}-1)*3+3),""))'   # Anmerkung roh
            ws[f'A{r}'] = f'{p}.'
            ws[f'B{r}'] = f'=Mitarbeiter!K{mr}'
            ws[f'C{r}'] = f'=Mitarbeiter!L{mr}'
            ws[f'D{r}'] = f'=IF(ISNUMBER(M{r}),IF(M{r}=0,"",M{r}),"")'
            ws[f'E{r}'] = f'=IF(ISNUMBER(M{r}),IFERROR(INDEX(SPWerte,$L$1,(L{r}-1)*3+2),""),"")'
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
    wb.save(os.path.join(ordner, OFFEN, TP))


# ---------------------------------------------------------------- Inhaltsverzeichnis (Ordner Steuerung)
def inhaltsverzeichnis(ordner, N, namen):
    """Steuerung/000 Inhaltsverzeichnis.xlsx – die EINE Quelle für die Namen (Spalte B, gelb, trägt der Steuerer ein) und
    Sammelstelle für die Vorgaben aller Stundenzettel (verstecktes Blatt „Vorgaben“, je Platz eine Verknüpfung auf das Blatt
    Vorgaben von „Stundenzettel NN.xlsx“). Die Eingabe-Dateien holen sich daraus ihre Zeile."""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Inhaltsverzeichnis'
    ws['A1'] = 'Inhaltsverzeichnis – Mitarbeiter'; ws['A1'].font = Font(bold=True, size=14, color=BLAU)
    ws['A2'] = ('Namen hier eintragen (Nachname, Vorname) – nur hier, alle anderen Dateien übernehmen sie. Im laufenden Jahr '
                'niemanden löschen oder überschreiben (auch nicht nach dem Ausscheiden); neue Mitarbeiter in die nächste freie Zeile. '
                'Stundenzettel zum Aktualisieren besser im Explorer per Doppelklick öffnen (über den Link hier klappt „Aktualisieren“ nicht).')
    ws['A2'].alignment = Alignment(wrap_text=True, vertical='top'); ws.merge_cells('A2:D2'); ws.row_dimensions[2].height = 62
    for i, h in enumerate(('Nr.', 'Mitarbeiter (Nachname, Vorname)', 'Eingabe-Datei', 'Stundenzettel')): kopf(ws.cell(3, 1 + i), h)
    vg = wb.create_sheet('Vorgaben')
    for i, h in enumerate(['Nr.', 'Name', 'Jahr', 'Urlaub (Tage)', 'Soll (Sek.)']): vg.cell(1, 1 + i, h)
    for mo in range(12):
        for j, h in enumerate(('ohne Vertrag', 'Regel-Tag Sek.', '057 Sek.', '057 da')):
            vg.cell(1, 6 + 4 * mo + j, f'{mo + 1:02d} {h}')
    for s in range(1, N + 1):
        r = 3 + s; datei = f'{SZ}{s:02d}.xlsx'
        ws[f'A{r}'] = s
        ws[f'B{r}'] = namen[s - 1] if s - 1 < len(namen) and namen[s - 1] else None
        ws[f'B{r}'].fill = PatternFill('solid', fgColor=GELB)
        ws[f'B{r}'].protection = openpyxl.styles.Protection(locked=False)
        ws[f'C{r}'] = f'{PLATZ}{s:02d}'
        ws[f'D{r}'] = datei[:-5]; ws[f'D{r}'].hyperlink = datei; ws[f'D{r}'].font = Font(color='FF0563C1', underline='single')
        for c in 'ABCD': ws[f'{c}{r}'].border = rahmen
        ws[f'A{r}'].alignment = Alignment(horizontal='center')
        # Vorgaben aus dem Stundenzettel (nur einfache Bezüge auf die Verknüpfung)
        nr = externer_link(wb, datei, ['Vorgaben'])
        z = 1 + s
        vg[f'A{z}'] = s
        vg[f'B{z}'] = f'=IF(TRIM(Inhaltsverzeichnis!B{r})="","",TRIM(Inhaltsverzeichnis!B{r}))'
        vg[f'C{z}'] = f'=[{nr}]Vorgaben!A1'; vg[f'D{z}'] = f'=[{nr}]Vorgaben!B1'; vg[f'E{z}'] = f'=[{nr}]Vorgaben!C1'
        for mo in range(12):
            col = chr(65 + mo)
            for j in range(4):
                vg.cell(z, 6 + 4 * mo + j).value = f'=[{nr}]Vorgaben!{col}{3 + j}'
    vg.sheet_state = 'hidden'
    for c, w in zip('ABCD', (6, 40, 16, 22)): ws.column_dimensions[c].width = w
    ws.freeze_panes = 'A4'
    ws.protection.sheet = True
    wb.save(os.path.join(ordner, PRIVAT, IV))


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--ordner', required=True)
    ap.add_argument('--plaetze', type=int, default=MAXP)
    ap.add_argument('--wochen', type=int, default=52)
    ap.add_argument('--tage', type=int, default=0)
    ap.add_argument('--mitarbeiter', nargs='*', default=[])
    ap.add_argument('--beispieldaten', action='store_true')
    ap.add_argument('--testnamen', action='store_true')
    ap.add_argument('--kopien', action='store_true')
    a = ap.parse_args()
    assert 1 <= a.plaetze <= MAXP, f'höchstens {MAXP} Plätze'
    namen = list(a.mitarbeiter)
    if a.testnamen:   # erfundene Namen „Mustermann, <Vorname>“ für die übrigen Zeilen
        frei = [v for v in TESTVORNAMEN if f'Mustermann, {v}' not in namen]
        namen += [f'Mustermann, {v}' for v in frei[:max(0, a.plaetze - len(namen))]]
    for o in (OFFEN, PRIVAT): os.makedirs(os.path.join(a.ordner, o), exist_ok=True)
    vorlagen_dateien(a.ordner, a.plaetze if a.kopien else 1, a.beispieldaten, a.testnamen and a.kopien)
    schichtplan(a.ordner, a.plaetze, a.wochen)
    tagesplan(a.ordner, a.plaetze, a.tage)
    inhaltsverzeichnis(a.ordner, a.plaetze, namen)
    print('Erzeugt in', a.ordner, ':', a.plaetze if a.kopien else 1, 'Eingabe-Datei(en) + Stundenzettel,', IV, TP, SP)
