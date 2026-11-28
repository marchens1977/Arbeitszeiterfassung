#!/usr/bin/env python3
"""Baut aus der Hauptvorlage excel/Arbeitszeiterfassung.xlsx die zwei Vorlagen für den Abteilungsbetrieb:

  excel/vorlagen/Eingabe.xlsx              – OFFENE Eingabe-Datei (wird „Platz 01.xlsx“ … „Platz 35.xlsx“, Ordner „Schichtplanung“):
                                             Musterwoche, Dateneingabe, Urlaub in Tagen (mit Warnung) – KEINE Vertrags-/Stundendaten.
                                             Platz-Nummer aus dem Dateinamen; Name und Vorgaben aus „000 Inhaltsverzeichnis.xlsx“.
  excel/vorlagen/Stundenzettel-Vorlage.xlsx – PRIVATER Stundenzettel (wird „Stundenzettel 01.xlsx“ …, Ordner „Steuerung“):
                                             Stammdaten, Urlaub, Auswertung, Tagesliste. Nummer aus dem Dateinamen; Name und
                                             Einträge aus „002 Schichtplan.xlsx“ (eigene Spalten per INDEX).
Die Verknüpfungen selbst (externalLink) setzt plaene-erzeugen.py. Alle Kopien einer Vorlage sind gleich – deshalb zeigen
sie nur auf die Sammel-Dateien, nie direkt aufeinander (siehe Kopf von plaene-erzeugen.py).

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
MAXP = 35                       # Plätze (Platz 01 … 35); Bereichsgrößen der Verknüpfungen hängen daran
IV_LETZTE = 'BA'                # letzte Spalte im Blatt Vorgaben des Inhaltsverzeichnisses (5 + 12 × 4 = 53)
SP_LETZTE = 'DB'                # letzte Spalte im Blatt Daten des Schichtplans (1 + 3 × 35 = 106)
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
HMS = lambda x: f'TEXT(INT(ABS({x})/3600),"00")&":"&TEXT(INT(MOD(ABS({x}),3600)/60),"00")'


def eingabe():
    """Offene Eingabe-Datei. Jahr, Urlaubsanspruch, Soll, Regel-Tag, Vertragsmonate und Werte der Stundenabrechnung 057
    kommen per Verknüpfung aus dem Stundenzettel (Blatt „Vorgaben“ → hier verstecktes Blatt „Vorgaben“, [1]-Bezüge;
    die Verknüpfung setzt plaene-erzeugen.py). Ohne Vorgaben: Jahr 2027, Vertrag in allen Monaten, keine Urlaubswarnung."""
    wb = openpyxl.load_workbook(HAUPT)
    ls, d, u, m, ex = wb['Listen'], wb['Dateneingabe'], wb['Urlaub'], wb['Musterwoche'], wb['Export']
    st = wb['Stammdaten']

    # --- Vorgaben aus dem Inhaltsverzeichnis (versteckt). Alle Eingabe-Dateien sind Kopien EINER Vorlage und zeigen auf
    # dasselbe Ziel („000 Inhaltsverzeichnis.xlsx“, Blatt Vorgaben: eine Zeile je Platz). Die eigene Zeile ergibt sich aus der
    # Platz-Nummer im Dateinamen („Platz 07.xlsx“ → 7, Listen!O10). Layout dort: A Nr, B Name, C Jahr, D Urlaubsanspruch,
    # E Soll (Sek.), ab F je Monat 4 Spalten (ohne Vertrag, Regel-Tag Sek., 057 Sek., 057 vorhanden).
    vg = wb.create_sheet('Vorgaben')
    hol = lambda sp: f'=INDEX([1]Vorgaben!$A$2:${IV_LETZTE}${1 + MAXP},Listen!$O$10,{sp})'
    vg['A1'] = hol(3); vg['B1'] = hol(4); vg['C1'] = hol(5); vg['D1'] = hol(2)
    for i in range(12):
        col = chr(65 + i)
        for r in (3, 4, 5, 6):
            vg[f'{col}{r}'] = hol(6 + 4 * i + (r - 3))
    vg['A7'] = ('Kommt aus dem Inhaltsverzeichnis des Steuerers (Zeile = Platz-Nummer aus dem Dateinamen): A1 Jahr, '
                'B1 Urlaubsanspruch (Tage), C1 Soll (Sek.), D1 Name, Zeile 3 ohne Vertrag (1), Zeile 4 Regel-Tag (Sek.), '
                'Zeile 5 Stundenabrechnung 057 (Sek.), Zeile 6 057 vorhanden (1)')
    vg.sheet_state = 'hidden'; vg.protection.sheet = True

    # --- Hilfswerte in „Listen“ (versteckt)
    ls['N1'] = 'Eingabe-Datei: Hilfswerte'; ls['N1'].font = Font(bold=True)
    ls['O2'] = '=CELL("filename",A1)'
    ls['O3'] = st['H3'].value.replace('G3', 'O2')
    ls['O4'] = '=IFERROR(IF(OR(ISNUMBER(Vorgaben!D1),Vorgaben!D1=""),"",TRIM(Vorgaben!D1)),"")'   # Name aus dem Inhaltsverzeichnis
    # Platz-Nummer aus dem Dateinamen „Platz 07.xlsx“ (0 = Dateiname falsch oder nicht lesbar)
    ls['O10'] = (f'=IFERROR(IF(LEFT(O3,6)="Platz ",IF(AND(VALUE(TRIM(MID(O3,7,FIND(".",O3&".",7)-7)))>=1,'
                 f'VALUE(TRIM(MID(O3,7,FIND(".",O3&".",7)-7)))<={MAXP}),VALUE(TRIM(MID(O3,7,FIND(".",O3&".",7)-7))),0),0),0)')
    # IFERROR: kommt über die Verknüpfung ein Fehler (Stundenzettel noch nie gespeichert/nicht erreichbar), gilt „keine Vorgabe“
    ls['O5'] = '=IFERROR(IF(N(Vorgaben!A1)>2000,Vorgaben!A1,2027),2027)'
    ls['O6'] = '=IFERROR(IF(ISNUMBER(Vorgaben!B1),Vorgaben!B1,0),0)'
    ls['O7'] = '=IFERROR(IF(ISNUMBER(Vorgaben!C1),Vorgaben!C1,0),0)'
    ls['O8'] = '=SUM(T2:T13)'
    ls['O9'] = '=O7-O8'
    for t, r in (('Pfad', 2), ('Dateiname', 3), ('Mitarbeiter', 4), ('Jahr', 5), ('Urlaubsanspruch (0 = unbekannt)', 6),
                 ('Soll Sek. (0 = unbekannt)', 7), ('Ist Sek. (Plan bzw. 057)', 8), ('Rest Sek.', 9),
                 ('Platz-Nr. (0 = Dateiname falsch)', 10)):
        ls[f'N{r}'] = t
    for i in range(12):
        r, col = 2 + i, chr(65 + i)
        ls[f'P{r}'] = i + 1
        ls[f'Q{r}'] = f'=IFERROR(IF(N(Vorgaben!{col}3)=1,1,0),0)'
        ls[f'R{r}'] = f'=IFERROR(N(Vorgaben!{col}4),0)'
        ls[f'S{r}'] = f'=IFERROR(IF(N(Vorgaben!{col}6)=1,N(Vorgaben!{col}5),""),"")'
        ls[f'T{r}'] = (f'=IF(Q{r}=1,0,IF(S{r}<>"",S{r},SUMIFS(Dateneingabe!$K$6:$K$267,Dateneingabe!$L$6:$L$267,P{r},'
                       f'Dateneingabe!$G$6:$G$267,"<>Urlaub")))')
    for t, c in (('Monat', 'P'), ('ohne Vertrag', 'Q'), ('Regel-Tag Sek.', 'R'), ('057 Sek.', 'S'), ('Ist Sek.', 'T')):
        ls[f'{c}1'] = t
    names = {'Jahr': 'Listen!$O$5', 'Mitarbeiter': 'Listen!$O$4', 'UrlaubAnspruchTage': 'Listen!$O$6',
             'SollSek': 'Listen!$O$7', 'RegelTagSek': 'Listen!$R$2:$R$13', 'OhneVertrag': 'Listen!$Q$2:$Q$13'}
    for n in ('UrlaubAnspruchSek',):
        del wb.defined_names[n]
    for n, ref in names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
    ls['G2'] = '=Jahr'

    # --- Dateneingabe
    for r in range(6, 268):
        d[f'N{r}'] = f'=IF(A{r}="","",INDEX(OhneVertrag,L{r}))'
        d[f'M{r}'] = f'=IF(A{r}="","",INDEX(RegelTagSek,L{r}))'
        d[f'D{r}'] = f'=IF(A{r}="","",IF(N{r}=1,"",IF(C{r}<>"","Feiertag",INDEX(MusterStart,WEEKDAY(A{r},2)))))'
        d[f'E{r}'] = f'=IF(A{r}="","",IF(OR(N{r}=1,NOT(ISNUMBER(D{r}))),"",INDEX(MusterPause,WEEKDAY(A{r},2))))'
        d[f'F{r}'] = f'=IF(A{r}="","",IF(OR(N{r}=1,NOT(ISNUMBER(D{r}))),"",INDEX(MusterAnm,WEEKDAY(A{r},2))))'
        d[f'G{r}'] = f'=IF(A{r}="","",IF(N{r}=1,"kein Vertrag",IF(D{r}<>"",D{r},IF(C{r}<>"","Feiertag",INDEX(MusterStart,WEEKDAY(A{r},2))))))'
        d[f'J{r}'] = f'=IF(A{r}="","",IF(ISNUMBER(G{r}),K{r}/86400,""))'
        d[f'N{r}'].font = Font(color='FF7F7F7F')
    d.column_dimensions['N'].hidden = True
    leeren(d, 'O5:P18')
    d['A1'] = '="Dateneingabe "&Jahr&IF(Mitarbeiter<>""," – "&Mitarbeiter,"")'
    d['A3'] = 'Reststunden zum Jahresende (nach deinen Einträgen)'
    d.merge_cells('H3:J3')
    d['H3'] = ('=IF(SollSek=0,"noch keine Vorgabe vom Steuerer",IF(Listen!$O$9>0,"noch zu leisten: "&' + HMS('Listen!$O$9') +
               '&" Std.",IF(Listen!$O$9<0,"zu viel geleistet: "&' + HMS('Listen!$O$9') + '&" Std.","Soll genau erreicht")))')
    for c in ('H3', 'I3', 'J3'):
        d[c].fill = GRAU; d[c].border = RAND
    d['H3'].font = Font(bold=True); d['H3'].alignment = Alignment(horizontal='center')
    gruen = DifferentialStyle(font=Font(color='FF00843D', bold=True))
    rot = DifferentialStyle(font=Font(color='FFC00000', bold=True))
    d.conditional_formatting.add('H3:J3', Rule(type='expression', dxf=gruen, formula=['AND(SollSek<>0,Listen!$O$9<0)']))
    d.conditional_formatting.add('H3:J3', Rule(type='expression', dxf=rot, formula=['AND(SollSek<>0,Listen!$O$9>0)']))
    d['A4'] = (f'=IF(Listen!$O$10=0,"Achtung: Dateiname falsch oder nicht lesbar – die Datei muss „Platz 01.xlsx“ … „Platz {MAXP}.xlsx“ '
               f'heißen (Excel am PC).",'
               'IF(AND(UrlaubAnspruchTage>0,COUNTIF(G6:G267,"Urlaub")>UrlaubAnspruchTage),"Achtung: "&COUNTIF(G6:G267,"Urlaub")&'
               '" Urlaubstage eingetragen – Anspruch: "&UrlaubAnspruchTage&" Tage!",""))')
    d.merge_cells('A4:J4'); d['A4'].font = Font(bold=True, color='FFC00000')
    d.conditional_formatting.add('D6:D267', Rule(type='expression', dxf=ROT, stopIfTrue=True,
                                 formula=['AND(UrlaubAnspruchTage>0,$G6="Urlaub",COUNTIF($G$6:$G6,"Urlaub")>UrlaubAnspruchTage)']))
    d['A2'] = ('Die gelben Felder zeigen bereits die Werte aus der Musterwoche. Weicht ein Tag ab, im gelben Feld aus dem Klappmenü '
               'wählen – eigene Einträge erscheinen fett. Hessische Feiertage sind vorbelegt. Graue Spalten rechnet Excel. '
               'Hellorange = Tag liegt in der Zukunft. Rot = mehr Urlaub als Anspruch. Anmerkung: Homeoffice, Teammentor*in; '
               '„abwesend (Gutschrift)“ trägt nur der Steuerer ein. Oben: deine Reststunden zum Jahresende, wenn alles so bleibt wie eingetragen. '
               'Eigenen Eintrag zurücknehmen: ein nicht fettes Feld derselben Spalte kopieren und einfügen.')

    # --- Musterwoche, Urlaub: Titel mit Namen, ohne Nummer
    m['A1'] = '="Musterwoche (Montag bis Freitag)"&IF(Mitarbeiter<>""," – "&Mitarbeiter,"")'
    u['A1'] = '="Urlaubsübersicht"&IF(Mitarbeiter<>""," – "&Mitarbeiter,"")'
    leeren(u, 'C8:D21'); leeren(u, 'A23:D26')
    ohne_regeln(u, lambda rng, f: 'D$26' in f)
    u['A2'] = ('Übersicht in Tagen: Urlaubsanspruch (legt der Steuerer fest), eingetragener Urlaub und Rest. Urlaub trägst du in der '
               'Dateneingabe ein (Klappmenü „Urlaub“). Rot = eingetragener Urlaub weicht vom Anspruch ab, Grün = genau gleich.')

    # --- Export
    ex['B1'] = '=Listen!$O$4'; ex['F1'] = '=Listen!$O$5'
    ex['G1'] = 'Anspruch'; ex['H1'] = '=Listen!$O$6'
    for i, h in enumerate(('Gilt: Start/Status (roh)', 'Gilt: Pause (roh)', 'Gilt: Anmerkung (roh)')):
        ex.cell(3, 5 + i, h).font = Font(bold=True)
    for k in range(1, 263):
        r, q = 3 + k, 5 + k
        ex[f'E{r}'] = f'=Dateneingabe!G{q}'; ex[f'E{r}'].number_format = 'hh:mm'
        ex[f'F{r}'] = f'=Dateneingabe!H{q}'
        ex[f'G{r}'] = f'=Dateneingabe!I{q}'

    # --- Blätter ohne Vertragsdaten entfernen
    del wb['Auswertung']; del wb['Stammdaten']
    for n in ('ListeUebertrag',):
        if n in wb.defined_names: del wb.defined_names[n]
    rest = [(ws.title, c.coordinate) for ws in wb for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and ('Stammdaten!' in c.value or 'Auswertung!' in c.value)]
    rest += [(n, dn.attr_text) for n, dn in wb.defined_names.items() if 'Stammdaten' in dn.attr_text or 'Auswertung' in dn.attr_text]
    assert not rest, f'Verweise auf entfernte Blätter: {rest[:5]}'

    del wb['Anleitung']                     # keine Anleitung – erklärt der Steuerer bei der Einweisung
    wb.active = wb.sheetnames.index('Dateneingabe')
    wb.save(os.path.join(ZIEL, 'Eingabe.xlsx'))


# =============================================================== Stundenzettel
def stundenzettel():
    wb = openpyxl.load_workbook(HAUPT)
    del wb['Export']
    blatt_umbenennen(wb, 'Dateneingabe', 'Tagesliste')
    wb['Musterwoche'].sheet_state = 'hidden'

    # Einträge kommen aus dem Schichtplan (der sammelt alle Eingabe-Dateien). Alle Stundenzettel sind Kopien EINER Vorlage;
    # die eigenen Spalten ergeben sich aus der Nummer im Dateinamen („Stundenzettel 07.xlsx“ → 7, Daten!B2).
    st = wb['Stammdaten']
    da = wb.create_sheet('Daten')
    da['A1'] = 'Name'; da['A2'] = 'Nr. (Dateiname)'; da['A3'] = 'Werktag → Start / Pause / Anmerkung (roh)'
    da['E1'] = '=CELL("filename",A1)'
    da['F1'] = st['H3'].value.replace('G3', 'E1')
    n_aus = 'VALUE(TRIM(MID(F1,15,FIND(".",F1&".",15)-15)))'
    da['B2'] = f'=IFERROR(IF(LEFT(F1,14)="Stundenzettel ",IF(AND({n_aus}>=1,{n_aus}<={MAXP}),{n_aus},0),0),0)'
    da['B1'] = f'=INDEX([1]Mitarbeiter!$C$4:$C${3 + MAXP},$B$2)'
    for r in range(4, 266):
        k = r - 3
        for j, col in enumerate('BCD'):
            da[f'{col}{r}'] = f'=INDEX([1]Daten!$B$4:${SP_LETZTE}$265,{k},$B$2*3-{2 - j})'
    da.sheet_state = 'hidden'

    t = wb['Tagesliste']
    for r in range(6, 268):
        q = r - 2
        t[f'D{r}'] = f'=IFERROR(IF(AND(ISNUMBER(Daten!B{q}),Daten!B{q}=0),"",Daten!B{q}),"")'   # IFERROR: Verknüpfung noch ohne Werte
        t[f'E{r}'] = f'=IFERROR(Daten!C{q},"")'
        t[f'F{r}'] = f'=IFERROR(IF(AND(ISNUMBER(Daten!D{q}),Daten!D{q}=0),"",Daten!D{q}),"")'
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
    # Überschriften ohne Nummern, überall mit Namen (Wunsch des Nutzers nach Test V5)
    NAME = '&IF(Mitarbeiter<>""," – "&Mitarbeiter,"")'
    t['A1'] = '="Tagesliste "&Jahr' + NAME
    wb['Stammdaten']['A1'] = '="Stammdaten"' + NAME
    wb['Urlaub']['A1'] = '="Urlaubsübersicht & Abgleich"' + NAME
    wb['Auswertung']['A1'] = '="Auswertung & Jahresbilanz "&Jahr' + NAME
    wb['Musterwoche']['A1'] = '="Musterwoche (Montag bis Freitag)"' + NAME
    t['A2'] = ('Diese Liste kommt automatisch aus der Eingabe-Datei des Mitarbeiters (gleiche Nummer, Ordner „Schichtplanung“) '
               'über den Schichtplan. Hier nichts eintragen – Eingabe-Datei speichern, Schichtplan aktualisieren und speichern, '
               'dann hier Daten → Links bearbeiten → „Werte aktualisieren“. Grau „kein Vertrag“ = Monat mit Jahres-Basis 0.')

    st['B3'] = '=IFERROR(IF(OR(ISNUMBER(Daten!B1),Daten!B1=""),"",Daten!B1),"")'
    st['G3'] = None; st['H3'] = None
    st['D3'] = (f'=IF(Daten!B2=0,"Dateiname falsch oder nicht lesbar – die Datei muss „Stundenzettel 01.xlsx“ … „Stundenzettel {MAXP}.xlsx“ heißen.",'
                'IF(B3="","Noch kein Name – im Inhaltsverzeichnis eintragen, dann Schichtplan und hier aktualisieren '
                '(Daten → Links bearbeiten → Werte aktualisieren).",""))')

    # Vorgaben für die Eingabe-Datei (die liest dieses Blatt per Verknüpfung; nur direkte Zellbezüge, keine Namen)
    v = wb.create_sheet('Vorgaben')
    v['A1'] = '=Stammdaten!$B$4'; v['B1'] = '=Stammdaten!$G$37'; v['C1'] = '=Stammdaten!$G$32'
    for i in range(12):
        col = chr(65 + i)
        v[f'{col}3'] = f'=IF(Stammdaten!C{13 + i}>0,0,1)'
        v[f'{col}4'] = f'=Stammdaten!H{13 + i}'
        v[f'{col}5'] = f'=IF(ISNUMBER(Auswertung!C{5 + i}),ROUND(Auswertung!C{5 + i}*86400,0),0)'
        v[f'{col}6'] = f'=IF(ISNUMBER(Auswertung!C{5 + i}),1,0)'
    v['A7'] = ('Für die Eingabe-Datei: A1 Jahr, B1 Urlaubsanspruch (anteilig), C1 Soll (Sek.), Zeile 3: 1 = ohne Vertrag, '
               'Zeile 4: Regel-Tag (Sek.), Zeile 5: Stundenabrechnung 057 (Sek.), Zeile 6: 1 = 057-Wert vorhanden '
               '(leere Zellen kommen über Verknüpfungen als 0 an – deshalb das Kennzeichen)')
    v.sheet_state = 'hidden'

    for name in ('Stammdaten', 'Urlaub', 'Auswertung', 'Tagesliste'):
        ws = wb[name]
        ws.page_setup.orientation = 'portrait' if name != 'Tagesliste' else 'landscape'
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb['Tagesliste'].print_title_rows = '5:5'

    del wb['Anleitung']                     # keine Anleitung – erklärt der Steuerer bei der Einweisung
    wb.active = wb.sheetnames.index('Stammdaten')
    for ws in wb: ws.sheet_view.tabSelected = (ws.title == 'Stammdaten')
    wb.save(os.path.join(ZIEL, 'Stundenzettel-Vorlage.xlsx'))


if __name__ == '__main__':
    os.makedirs(ZIEL, exist_ok=True)
    eingabe(); stundenzettel()
    print('Erzeugt:', os.listdir(ZIEL))
