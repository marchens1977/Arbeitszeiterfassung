#!/usr/bin/env python3
"""Baut aus der Hauptvorlage excel/Arbeitszeiterfassung.xlsx die zwei Vorlagen für den Abteilungsbetrieb:

  excel/vorlagen/Eingabe.xlsx              – OFFENE Eingabe-Datei je Mitarbeiter (Ordner „Schichtplanung“):
                                             Mitarbeiter (Name aus Dateiname, Urlaubsanspruch in Tagen), Musterwoche,
                                             Dateneingabe, Urlaub in Tagen (mit Warnung) – KEINE Vertrags-/Stundendaten.
  excel/vorlagen/Stundenzettel-Vorlage.xlsx – PRIVATER Stundenzettel je Mitarbeiter (Ordner „Steuerung“):
                                             Stammdaten, Urlaub, Auswertung, Tagesliste. Holt die Einträge per Verknüpfung
                                             aus der Eingabe-Datei. Die Verknüpfung setzt plaene-erzeugen.py je Mitarbeiter.

Aufruf: python3 werkzeuge/vorlagen-bauen.py      (nach jeder Änderung an der Hauptvorlage erneut ausführen)
"""
import os, sys
from copy import copy
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import Rule
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.workbook.defined_name import DefinedName
sys.path.insert(0, os.path.dirname(__file__))
from xlsx_spalten import blatt_umbenennen

HIER = os.path.dirname(os.path.abspath(__file__))
HAUPT = os.path.join(HIER, '..', 'excel', 'Arbeitszeiterfassung.xlsx')
ZIEL = os.path.join(HIER, '..', 'excel', 'vorlagen')
GRAU = PatternFill('solid', fgColor='FFD9D9D9')
s = Side(style='thin', color='FFBFBFBF'); RAND = Border(left=s, right=s, top=s, bottom=s)
ROT = DifferentialStyle(font=Font(color='FF9C0006', bold=True),
                        fill=PatternFill(bgColor='FFFFC7CE', fgColor='FFFFC7CE', fill_type='solid'))


def leeren(ws, bereich):
    for row in ws[bereich]:
        for c in row:
            c.value = None; c.style = 'Normal'


def ohne_gueltigkeit(ws, zellen):
    ws.data_validations.dataValidation = [d for d in ws.data_validations.dataValidation
                                          if not any(z in str(d.sqref) for z in zellen)]


def ohne_regeln(ws, bedingung):
    neu = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        for r in cf.rules:
            if not bedingung(str(cf.sqref), r.formula[0] if r.formula else ''):
                neu.add(str(cf.sqref), r)
    ws.conditional_formatting = neu


def anleitung(ws, zeilen):
    for r in range(1, 60):
        for c in range(1, 4): ws.cell(r, c).value = None
    for i, (text, art) in enumerate(zeilen, start=1):
        c = ws.cell(i, 2, text); c.alignment = Alignment(wrap_text=True, vertical='top')
        c.font = Font(bold=art != 't', size=16 if art == 'h' else 12 if art == 'u' else 11, color='FF002D72' if art != 't' else None)


# =============================================================== Eingabe-Datei
def eingabe():
    wb = openpyxl.load_workbook(HAUPT)
    del wb['Auswertung']
    blatt_umbenennen(wb, 'Stammdaten', 'Start')
    st = wb['Start']
    for m in [str(x) for x in st.merged_cells.ranges]:
        if int(m.split(':')[0][1:]) >= 6: st.unmerge_cells(m)
    leeren(st, 'A6:H45')
    ohne_gueltigkeit(st, ['B6', 'B7', 'B13'])
    st['A1'] = '="1. Mitarbeiter"&IF(Mitarbeiter<>""," – "&Mitarbeiter,"")'
    st['A2'] = ('Der Name kommt aus dem Dateinamen („Nachname, Vorname.xlsx“). Den Urlaubsanspruch trägt der Steuerer ein '
                '(bei Ein- oder Austritt bereits anteilig). Vertrag und Stundenkonto stehen nur im Stundenzettel beim Steuerer.')
    st['A5'] = '="Urlaubsanspruch "&Jahr&" (Tage)"'
    st['C5'] = 'trägt der Steuerer ein'
    for n in ('RegelTagSek', 'SollSek', 'UrlaubAnspruchSek'):
        del wb.defined_names[n]
    wb.defined_names['UrlaubAnspruchTage'] = DefinedName('UrlaubAnspruchTage', attr_text='Start!$B$5')

    u = wb['Urlaub']
    leeren(u, 'C8:D21'); leeren(u, 'A23:D26')
    ohne_regeln(u, lambda rng, f: 'D$26' in f)
    u['A2'] = ('Übersicht in Tagen: zustehender Urlaub (trägt der Steuerer ein), eingetragener Urlaub und Rest. Urlaub trägst du in '
               'der Dateneingabe ein (Klappmenü „Urlaub“). Rot = eingetragener Urlaub weicht vom Anspruch ab, Grün = genau gleich.')

    d = wb['Dateneingabe']
    for r in range(6, 268):
        d[f'D{r}'] = f'=IF(A{r}="","",IF(C{r}<>"","Feiertag",INDEX(MusterStart,WEEKDAY(A{r},2))))'
        d[f'E{r}'] = f'=IF(A{r}="","",IF(NOT(ISNUMBER(D{r})),"",INDEX(MusterPause,WEEKDAY(A{r},2))))'
        d[f'F{r}'] = f'=IF(A{r}="","",IF(NOT(ISNUMBER(D{r})),"",INDEX(MusterAnm,WEEKDAY(A{r},2))))'
        d[f'G{r}'] = f'=IF(A{r}="","",IF(D{r}<>"",D{r},IF(C{r}<>"","Feiertag",INDEX(MusterStart,WEEKDAY(A{r},2)))))'
        d[f'J{r}'] = f'=IF(A{r}="","",IF(ISNUMBER(G{r}),K{r}/86400,""))'
        d[f'K{r}'] = f'=IF(A{r}="","",IF(ISNUMBER(G{r}),MAX(0,68400-ROUND(G{r}*86400,0)-H{r}*60),0))'
        d[f'M{r}'] = None
    leeren(d, 'O5:P18')
    d['A3'] = 'Summe geplante Arbeitszeit (nur Arbeitstage)'
    d['J3'] = '=SUM(K6:K267)/86400'
    d.merge_cells('A4:J4')
    d['A4'] = ('=IF(COUNTIF(G6:G267,"Urlaub")>UrlaubAnspruchTage,"Achtung: "&COUNTIF(G6:G267,"Urlaub")&'
               '" Urlaubstage eingetragen – Anspruch: "&UrlaubAnspruchTage&" Tage!","")')
    d['A4'].font = Font(bold=True, color='FFC00000')
    ohne_regeln(d, lambda rng, f: 'kein Vertrag' in f)
    d.conditional_formatting.add('D6:D267', Rule(type='expression', dxf=ROT, stopIfTrue=True,
                                 formula=['AND($G6="Urlaub",COUNTIF($G$6:$G6,"Urlaub")>UrlaubAnspruchTage)']))
    d['A2'] = ('Die gelben Felder zeigen bereits die Werte aus der Musterwoche. Weicht ein Tag ab, im gelben Feld aus dem Klappmenü '
               'wählen – eigene Einträge erscheinen fett. Hessische Feiertage sind vorbelegt. Graue Spalten rechnet Excel. '
               'Hellorange = Tag liegt in der Zukunft. Rot = mehr Urlaub als Anspruch. Anmerkung: Homeoffice, Teammentor*in; '
               '„abwesend (Gutschrift)“ trägt nur der Steuerer ein. Eigenen Eintrag zurücknehmen: ein nicht fettes Feld derselben '
               'Spalte kopieren und einfügen – dann gilt wieder die Musterwoche.')

    ex = wb['Export']
    ex['H1'] = '=Start!$B$5'; ex['G1'] = 'Anspruch'
    for i, h in enumerate(('Gilt: Start/Status (roh)', 'Gilt: Pause (roh)', 'Gilt: Anmerkung (roh)')):
        ex.cell(3, 5 + i, h).font = Font(bold=True)
    for k in range(1, 263):
        r, q = 3 + k, 5 + k
        ex[f'E{r}'] = f'=Dateneingabe!G{q}'; ex[f'E{r}'].number_format = 'hh:mm'
        ex[f'F{r}'] = f'=Dateneingabe!H{q}'
        ex[f'G{r}'] = f'=Dateneingabe!I{q}'

    anleitung(wb['Anleitung'], [
        ('Schichtplanung – Eingabe-Datei – Anleitung', 'h'), ('', 't'),
        ('Wofür ist diese Datei?', 'u'),
        ('Hier planst du deine Arbeitszeiten. Daraus entstehen automatisch Schichtplan und Tagesplan. Die Datei liegt im offenen '
         'Ordner – Kollegen können sie ansehen, aber nur du (Kennwort zum Ändern) und der Steuerer können sie ändern. '
         'Vertrag und Stundenkonto stehen NICHT hier, sondern im Stundenzettel beim Steuerer.', 't'),
        ('Farben', 'u'), ('Gelb = hier eintragen · Grau = rechnet Excel · fett = eigener Eintrag · Rot = mehr Urlaub als Anspruch', 't'),
        ('Schritt für Schritt', 'u'),
        ('1. Blatt „Start“: Name kommt aus dem Dateinamen („Nachname, Vorname.xlsx“). Den Urlaubsanspruch trägt der Steuerer ein.', 't'),
        ('2. Blatt „Musterwoche“: üblichen Arbeitsbeginn (11:00–17:00) oder „Frei“, Pause und ggf. Anmerkung (Homeoffice, Teammentor*in) wählen.', 't'),
        ('3. Blatt „Dateneingabe“: Die gelben Felder zeigen schon die Musterwoche. Abweichende Tage (Urlaub, Frei, anderer Beginn, '
         'Homeoffice …) im Klappmenü wählen. Speichern nicht vergessen – der Schichtplan liest nur gespeicherte Werte.', 't'),
        ('4. Blatt „Urlaub“: Anspruch, eingetragene Urlaubstage und Rest in Tagen.', 't'),
        ('„abwesend (Gutschrift)“', 'u'),
        ('Trägt nur der Steuerer ein. Die geplante Zeit wird gutgeschrieben; im Schichtplan steht nur „abwesend“.', 't'),
        ('Schutz', 'u'),
        ('Datei → Speichern unter → Tools → Allgemeine Optionen → „Kennwort zum Ändern“ (NICHT „Kennwort zum Öffnen“, '
         'sonst funktionieren Schichtplan und Stundenzettel nicht mehr). Keine Makros.', 't'),
    ])
    wb.save(os.path.join(ZIEL, 'Eingabe.xlsx'))


# =============================================================== Stundenzettel
def stundenzettel():
    wb = openpyxl.load_workbook(HAUPT)
    del wb['Export']
    blatt_umbenennen(wb, 'Dateneingabe', 'Tagesliste')
    wb['Musterwoche'].sheet_state = 'hidden'

    da = wb.create_sheet('Daten')
    da['A1'] = 'Name'; da['A2'] = 'Urlaubsanspruch laut Eingabe-Datei'; da['A3'] = 'Werktag → Start / Pause / Anmerkung'
    da['B1'] = '=[1]Export!$B$1'; da['B2'] = '=[1]Export!$H$1'
    for r in range(4, 266):
        da[f'B{r}'] = f'=[1]Export!E{r}'; da[f'C{r}'] = f'=[1]Export!F{r}'; da[f'D{r}'] = f'=[1]Export!G{r}'
    da.sheet_state = 'hidden'

    t = wb['Tagesliste']
    for r in range(6, 268):
        q = r - 2
        t[f'D{r}'] = f'=IF(AND(ISNUMBER(Daten!B{q}),Daten!B{q}=0),"",Daten!B{q})'
        t[f'E{r}'] = f'=Daten!C{q}'
        t[f'F{r}'] = f'=IF(AND(ISNUMBER(Daten!D{q}),Daten!D{q}=0),"",Daten!D{q})'
        t[f'G{r}'] = f'=IF(A{r}="","",IF(M{r}=0,"kein Vertrag",IF(D{r}<>"",D{r},IF(C{r}<>"","Feiertag",""))))'
        t[f'I{r}'] = f'=IF(A{r}="","",IF(NOT(ISNUMBER(G{r})),"",IF(F{r}="Büro","",F{r})))'
        for col in 'DEF':
            c = t[f'{col}{r}']; c.fill = GRAU; c.border = RAND
            c.protection = openpyxl.styles.Protection(locked=True)
            f = copy(c.font); f.b = False; f.color = 'FF000000'; c.font = f
    t.data_validations.dataValidation = []
    ohne_regeln(t, lambda rng, f: 'ISFORMULA' in f)
    for col, txt in zip('DEF', ('Eingetragen:\nStart / Status', 'Eingetragen:\nPause', 'Eingetragen:\nAnmerkung')):
        t[f'{col}5'] = txt
    t['A1'] = '="4. Tagesliste "&Jahr&IF(Mitarbeiter<>""," – "&Mitarbeiter,"")'
    t['A2'] = ('Diese Liste kommt automatisch aus der Eingabe-Datei des Mitarbeiters (Ordner „Schichtplanung“). Hier nichts '
               'eintragen – Änderungen in der Eingabe-Datei vornehmen, dort speichern, dann hier Daten → Links bearbeiten → '
               '„Werte aktualisieren“. Grau „kein Vertrag“ = Monat mit Jahres-Basis 0.')

    st = wb['Stammdaten']
    st['B3'] = '=IF(ISNUMBER(Daten!B1),"",Daten!B1)'
    st['G3'] = None; st['H3'] = None
    st['D3'] = ('=IF(B3="","Noch keine Daten – Verknüpfung zur Eingabe-Datei aktualisieren (Daten → Links bearbeiten → '
                'Werte aktualisieren).","")')

    u = wb['Urlaub']
    u['A28'] = 'Urlaubsanspruch laut Eingabe-Datei'; u['A28'].font = Font(bold=True)
    u['B28'] = '=IF(ISNUMBER(Daten!B2),Daten!B2,"")'; u['B28'].number_format = 'General" Tage"'
    u['B28'].fill = GRAU; u['B28'].border = RAND
    u['C28'] = ('=IF(B28="","",IF(ROUND(B28,2)<>ROUND(UrlaubAnspruchTage,2),"Achtung: In der Eingabe-Datei steht "&B28&'
                '" Tage – bitte dort auf "&UrlaubAnspruchTage&" ändern.",""))')
    u['C28'].font = Font(bold=True, color='FFC00000')

    v = wb.create_sheet('Vertrag')
    for i in range(12):
        v.cell(1, 1 + i).value = f'=IF(Stammdaten!C{13 + i}>0,1,0)'
    v['A2'] = 'Merker je Monat (1 = Vertrag) – liest der Schichtplan, damit nach einem Austritt nichts mehr angezeigt wird.'
    v.sheet_state = 'hidden'

    for name in ('Stammdaten', 'Urlaub', 'Auswertung', 'Tagesliste'):
        ws = wb[name]
        ws.page_setup.orientation = 'portrait' if name != 'Tagesliste' else 'landscape'
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb['Tagesliste'].print_title_rows = '5:5'

    anleitung(wb['Anleitung'], [
        ('Stundenzettel – Anleitung (nur für den Steuerer)', 'h'), ('', 't'),
        ('Wofür ist diese Datei?', 'u'),
        ('Der Stundenzettel liegt im Ordner „Steuerung“, auf den nur der Steuerer Zugriff hat. Er enthält Vertrag (Stammdaten), '
         'Urlaub, Stundenkonto (Auswertung) und die Tagesliste. Die Tagesliste kommt automatisch aus der Eingabe-Datei des '
         'Mitarbeiters im Ordner „Schichtplanung“.', 't'),
        ('Aktualisieren', 'u'),
        ('Mitarbeiter speichert seine Eingabe-Datei → hier Datei öffnen → „Inhalt aktivieren“ bzw. Daten → Links bearbeiten → '
         '„Werte aktualisieren“ → speichern (wichtig: der Schichtplan liest von hier, ob ein Vertrag besteht).', 't'),
        ('Ausdruck / PDF für den Mitarbeiter', 'u'),
        ('Blätter Stammdaten, Urlaub, Auswertung und Tagesliste markieren (Strg + Klick auf die Reiter) → Datei → Drucken bzw. '
         'Exportieren → PDF. Die Blätter sind auf Seitenbreite eingestellt.', 't'),
        ('Abwesenheit (z. B. Krankheit)', 'u'),
        ('In der EINGABE-Datei des Mitarbeiters am Tag die Anmerkung „abwesend (Gutschrift)“ wählen. Die geplante Zeit wird '
         'gutgeschrieben; Schicht- und Tagesplan zeigen nur „abwesend“. Der Grund wird nirgends gespeichert.', 't'),
        ('Ein- und Austritt', 'u'),
        ('In den Stammdaten die Jahres-Basis für Monate ohne Vertrag auf 0:00 setzen → Urlaub wird anteilig, Schichtplan zeigt in '
         'diesen Monaten nichts mehr. Den anteiligen Urlaubsanspruch (siehe Stammdaten) zusätzlich in der Eingabe-Datei eintragen – '
         'das Blatt „Urlaub“ hier warnt, wenn beide nicht übereinstimmen.', 't'),
        ('Rechenregeln', 'u'),
        ('Regel-Tag = Jahresstunden ÷ 52 ÷ 5 · Arbeitstag: 19:00 − Beginn − Pause · Feiertag zählt einen Regel-Tag aufs Ist · '
         'Urlaub mindert das Soll · Frei = 0 · Soll = Vertrag − Urlaub ± Übertrag + vereinbarte Überstunden − unbezahlter Urlaub.', 't'),
    ])
    wb.save(os.path.join(ZIEL, 'Stundenzettel-Vorlage.xlsx'))


if __name__ == '__main__':
    os.makedirs(ZIEL, exist_ok=True)
    eingabe(); stundenzettel()
    print('Erzeugt:', os.listdir(ZIEL))
