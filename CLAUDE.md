# Projekt-Tagebuch: Arbeitszeiterfassung

## 1. Worum es geht

Eine kleine Web-App, mit der der Nutzer seine Arbeitszeit für ein Kalenderjahr (aktuell 2027) plant und abgleicht:
Stammdaten (vertragliche Jahresstunden je Monat, Urlaubsanspruch, Übertrag), Urlaubsübersicht, Musterwoche,
tägliche Dateneingabe (Start/Pause, Ende fix 19:00 Uhr) und eine Jahresbilanz (Soll gegen Ist, Vergleich mit
„Stundenabrechnung 057“). Hessische Feiertage werden automatisch berechnet.

**Online:** GitHub Pages des Repositories `marchens1977/Arbeitszeiterfassung` (Hauptzweig `main`).
Die App ist als installierbare App (PWA) eingerichtet und läuft auch offline.

**Technischer Aufbau** (alles ohne Framework, nur HTML/CSS/JavaScript):

| Datei | Zweck |
|---|---|
| `index.html` | Die komplette App: Oberfläche, Rechenlogik, Speicherung, Datensicherung |
| `manifest.json` | Beschreibung für die Installation als App (Name, Symbol, Farben) |
| `sw.js` | Offline-Helfer (Service Worker): speichert die App-Dateien; „Netz zuerst, sonst Kopie“ |
| `icons/` | App-Symbol (blaue Uhr) als SVG und PNG (192, 512, 512 maskable, 180 Apple) |
| `excel/Arbeitszeiterfassung.xlsx` | Excel-Version der App (eigenständig, siehe Abschnitt 9) – **Hauptvorlage**, aus der die beiden Vorlagen unten gebaut werden |
| `excel/vorlagen/Eingabe.xlsx`, `excel/vorlagen/Stundenzettel-Vorlage.xlsx` | Abgeleitete Vorlagen (NICHT von Hand ändern – `werkzeuge/vorlagen-bauen.py`); die Stundenzettel-Vorlage enthält Platzhalter-Verknüpfungen `[1]Export!…` und ist allein nicht benutzbar |
| `excel/Programmeinsatzverfahren-Aenderung.md` | Änderungsmeldung zum bereits genehmigten Programmeinsatzverfahren (Excel, Stand V9) – Vorlage für den Nutzer, Platzhalter in [ ]; bei Änderungen am Excel-Paket mitpflegen |
| `werkzeuge/` | Automatische Prüfungen (Tests) und Hilfsprogramme, siehe Abschnitt 8 |
| `CLAUDE.md` | Dieses Tagebuch |

Wichtige Stellen in `index.html`:
- `const Y = 2027` – das Jahr; für ein neues Jahr nur diese Zahl ändern.
- `holidaysForYear()` – berechnet die hessischen Feiertage (Ostern nach Gauß).
- `defaultVacations` – vorbelegter Urlaub (nur für 2027 hinterlegt).
- `computeStats()` – die **einzige** Stelle, an der Soll/Ist/Urlaub berechnet werden; alle Reiter nutzen sie.
- Daten liegen im Browser-Speicher (`localStorage`) unter dem Schlüssel `arbeitszeit_hessen_<Jahr>_v2`;
  der alte Schlüssel `arbeitszeit_hessen_persistence_v1` wird beim Start automatisch übernommen.

## 2. Regeln und Entscheidungen

- **Grundregel: Die Excel-Datei ist die Vorlage, die App zieht nach** (Festlegung des Nutzers, 27.11.2026).
  Hintergrund: Die Excel-Datei liegt zentral im Abteilungsordner und darf von allen Mitarbeitern genutzt werden.
  Die App („für die Hosentasche“, Stundenkonto auch von zu Hause) muss erst noch geprüft/freigegeben werden.
  Daraus folgt:
  1. Logik-Änderungen wünscht der Nutzer zuerst für die Excel-Datei. Was dort entschieden ist, gilt auch für die App
     und wird in der App **genau gleich** nachgebaut (gleiche Rechnung, gleiche Rundung, gleiche Ergebnisse).
  2. Auch das **Aussehen** soll möglichst gleich sein (Farben, Rot/Grün-Anzeigen, Bezeichnungen, Auswahllisten),
     damit Mitarbeiter mit beiden Varianten gleich umgehen können.
  3. Nach jeder Excel-Änderung prüfen, ob die App nachgezogen werden muss, und dem Nutzer sagen, was noch abweicht
     (Liste in Abschnitt 4 pflegen). Ohne Rückfrage nachziehen nur, wenn der Nutzer es so beauftragt.
  4. **Keine Synchronisation** zwischen Excel und App – technisch nicht vorgesehen und ausdrücklich nicht gewünscht.
     Beide werden inhaltlich getrennt geführt; der Mitarbeiter überträgt seine Daten selbst, je nachdem, was er hauptsächlich nutzt.

- **Rechenmodell nicht verändern**, es sei denn, der Nutzer wünscht es ausdrücklich: Regel-Tag = Jahresstunden / 52 / 5;
  Urlaub wird vom Soll abgezogen; Feiertage zählen als Regel-Tag auf das Ist; Urlaubstage zählen nicht auf das Ist.
- **Ausscheiden/Eintritt unter dem Jahr** (Wunsch des Nutzers, 27.11.2026, in App **und** Excel gleich):
  Monat mit Jahres-Basis 0 = kein Vertrag (zählt nichts, auch kein Urlaub/Musterwoche). Urlaub anteilig =
  Anspruch × Vertragsmonate ÷ 12, ab halbem Tag aufrunden (BUrlG § 5), kleinere Bruchteile bleiben (2 Stellen).
  Urlaubsstunden = Tage × Ø Regel-Tag **nur der Vertragsmonate**. In der App: `isContractMonth()`, `proratedVacDays()`,
  `fmtDays()` (Komma statt Punkt). Urlaubs-Karten: grün (`vac-ok`) bei genommen = zustehend, sonst rot (`vac-diff`).
- App und Excel müssen gleich rechnen: Bei Änderungen am Rechenmodell beide anpassen und `ausscheiden.test.js`
  (Sollwerte stammen aus `excel-pruefen.py`) laufen lassen.
- Bei Umbauten immer **alte und neue Version mit denselben Daten im Browser vergleichen** (Playwright);
  alle Zahlen müssen sekundengenau gleich bleiben.
- Zeiteingaben werden beim Speichern mit `normHMS()` auf `HH:MM:SS` gebracht („7,5“ → `07:30:00`).
- In der Musterwoche sind nur Uhrzeiten und „Frei“ erlaubt (kein Urlaub/Feiertag).
- Bei jeder Änderung an der App die `VERSION` in `sw.js` hochzählen. **Wichtig:** Nur daran erkennt die App ein Update
  (der Browser vergleicht `sw.js`). Die angezeigte „Installierte Version“ ist genau dieser Wert (z. B. `v4`).
- Updates: `sw.js` installiert sich sofort (`skipWaiting`), die Seite lädt bei `controllerchange` neu – aber nicht,
  solange ein „Ändern“-Modus offen ist (`editing0/2/3/4`); dann erst nach dem Speichern (`setEditButton`).
  Netzabfragen im Offline-Helfer laufen mit `cache: 'no-cache'`, damit keine veraltete Datei aus dem Browser-Zwischenspeicher kommt.
- Der Name unter dem App-Symbol muss der **volle** Name sein: `short_name` = „Arbeitszeiterfassung“.
- Vorgehen für Änderungen: eigener Branch → Pull Request → Merge in `main`. Der Nutzer hat ausdrücklich erlaubt,
  dass Claude PRs anlegt **und** zusammenführt, damit die neue Version sofort live ist.
- Keine persönlichen Daten, Passwörter oder Schlüssel in Code, Commits oder dieses Tagebuch schreiben.
- **Farben:** Mets-Blau `#002D72` (Hauptfarbe, Reiterleiste, Knopf „Ändern“, Überschriften, Monatszeilen) und
  Mets-Orange `#FF5910` (aktiver Reiter, Knopf „Speichern“ mit blauer Schrift, Akzentlinien, Uhrzeiger im Symbol).
  Zukünftige Tage sind hellorange (`#fff0e8`). Grün/Rot bleiben für Plus/Minus und Feiertage, weil sie eine Bedeutung haben.
  Nach Änderung am Symbol: `node werkzeuge/symbole-erzeugen.js` (Hintergrundfarbe steht dort fest drin).
- **Farben wie in der Excel-Datei** (seit #14): Eingabefeld im „Ändern“-Modus gelb `#FFF2CC` mit orangem Rand `#FF5910`;
  gesperrte/berechnete Felder und Kästen grau `#F2F2F2` (Rand `#BFBFBF`); Urlaubstag hellblau `#DDEBF7`; „kein Vertrag“ `#EDEDED`;
  Feiertag dunkelrot `#C00000`; Urlaub-Abgleich grün/rot. Mitarbeitername: `st.name`, Anzeige `showEmployee()`.
- Uhrzeit-Auswahl in Dateneingabe und Musterwoche: 11:00–17:00 im 15-Minuten-Takt (`timeList`), wie in Excel.

## 3. Verlauf

| Datum | Was | PR |
|---|---|---|
| 2026-11-16 | Code-Überarbeitung: Tippfehler beim „Ändern“-Knopf, Zeiteingaben normalisiert, Musterwoche ohne Urlaub/Feiertag, Feiertage berechnet, Jahr nur an einer Stelle, zentrale Berechnung `computeStats()`, schnellere Dateneingabe (nur geänderte Zeile), Datensicherung (Datei herunterladen/teilen/wiederherstellen/zurücksetzen) | #1 |
| 2026-11-16 | Sicherung als Text (kopieren/einfügen), Anzeige der letzten Sicherung mit Erinnerung nach 30 Tagen, Dateiauswahl ohne Typfilter, Hinweis zum Download-Ordner im Browser | #2 |
| 2026-11-16 | Installierbare App (PWA): `manifest.json`, Symbole, `sw.js`, Knopf „App installieren“ | #3 |
| 2026-11-16 | Name unter dem Symbol auf „Arbeitszeiterfassung“ geändert | #4 |
| 2026-11-16 | Projekt-Tagebuch `CLAUDE.md` angelegt | #5 |
| 2026-11-18 | Regel „Hilfsprogramme und Tests gehören ins Projekt“; vorhandene Prüfungen aus dem Arbeitsbereich nach `werkzeuge/` übernommen | #6 |
| 2026-11-19 | Farben im Stil der New York Mets: Blau `#002D72`, Orange `#FF5910` (Reiter, Knöpfe, Überschriften, Karten, Monatszeilen, App-Symbol, Statusleiste); Rechnungen unverändert | #7 |
| 2026-11-19 | Automatische App-Updates: Kasten „App-Updates“ im Reiter Stammdaten (Version, letzte Suche, Knopf „Jetzt nach Updates suchen“); sucht beim Start, beim Zurückkehren und stündlich, installiert selbst, wartet während einer Bearbeitung; Test `updates.test.js`; Sicherungs-Test wartet jetzt auf das Kopieren | #8 |
| 2026-11-27 | Excel-Version `excel/Arbeitszeiterfassung.xlsx` ins Projekt übernommen und im Tagebuch beschrieben (Abschnitt 9) | #9 |
| 2026-11-27 | Excel: Komma hinter ganzen Zahlen entfernt (Format `0.##` → `General`), Jahres-Basis jetzt als Zeit `[hh]:mm:ss` (Eingabe `2028:00`); Prüfprogramm `werkzeuge/excel-pruefen.py` | #10 |
| 2026-11-27 | Excel: Klappmenü „Start / Status“ in der Dateneingabe nur noch 11:00–17:00 (15-Minuten-Takt; Liste `ListeTag` = `Listen!A2:A29`); Musterwoche unverändert | direkt auf `main` |
| 2026-11-27 | Excel: Stammdaten Zeile 3 „Mitarbeiter“ – Name wird aus dem Dateinamen gelesen („Mustermann, Max.xlsx“ → „Mustermann, Max“), Hinweis bei falschem Dateinamen; Name `Mitarbeiter` | #11 |
| 2026-11-27 | Excel: Mitarbeiter auffälliger – Überschrift „1. Stammdaten – Nachname, Vorname“ (bleibt beim Scrollen stehen), Feld B3:C3 verbunden, blau mit großer weißer Schrift und orangem Rahmen; Hinweis jetzt in D3 | direkt auf `main` |
| 2026-11-27 | Excel: Klappmenü „Start / Status“ in der Musterwoche ebenfalls nur noch Frei und 11:00–17:00 (15-Minuten-Takt; Liste `ListeMuster` = `Listen!B2:B27`) | direkt auf `main` |
| 2026-11-27 | Excel (Rechenmodell auf Wunsch erweitert): Ausscheiden unter dem Jahr – Monate mit Jahres-Basis 0 zählen nicht („kein Vertrag“, grau), Urlaub anteilig (Anspruch × Vertragsmonate ÷ 12, gesetzlich gerundet), Urlaubsstunden mit Ø Regel-Tag nur der Vertragsmonate; Urlaubsblatt: Block rot/grün; `excel-pruefen.py` mit `--ab-monat`, `--anspruch`, `--urlaub` | direkt auf `main` |
| 2026-11-27 | App: dieselbe Logik wie Excel – Ausscheiden unter dem Jahr (Monate mit Jahres-Basis 0 = „kein Vertrag“, grau), anteiliger Urlaub gesetzlich gerundet, Hinweis unter „Urlaubsanspruch volles Jahr“, Urlaubs-Karten grün/rot; `sw.js` v5; neuer Test `ausscheiden.test.js` (vergleicht mit den Excel-Werten) | #12 |
| 2026-11-27 | Grundregel ins Tagebuch: Excel-Datei ist Vorlage, App zieht Logik und Aussehen exakt gleich nach; keine Synchronisation; Liste der noch offenen Unterschiede | #13 |
| 2026-11-27 | App an Excel angeglichen: Uhrzeit-Auswahl (Dateneingabe + Musterwoche) nur 11:00–17:00 (alte gespeicherte Zeiten bleiben erhalten); Mitarbeitername (Feld „Mitarbeiter (Nachname, Vorname)“, Überschrift „1. Stammdaten – Name“, blaues Feld mit orangem Rand, Hinweis ohne Komma, in der Sicherung); Excel-Farben: Eingabe gelb/orange, gesperrt/berechnet grau, Urlaub hellblau, Feiertag dunkelrot, Legende; Überschriften/Reiter wie in Excel („5. Auswertung“); `sw.js` v6; neuer Test `name-und-aussehen.test.js` | #14 |
| 2026-11-27 | Excel: Spalte „Anmerkung“ (Dateneingabe F = Eingabe, I = gilt; Musterwoche D) mit Büro/Homeoffice/Teammentor*in/krank; „krank“ = geplante Zeit wird gutgeschrieben (Rechnung unverändert); verstecktes Blatt „Export“ für den Schichtplan. Neu: Schichtplan + MA-Tagesplan ohne Makros (Abschnitt 10), Mini-Paket zum Test beim Arbeitgeber; Werkzeuge `xlsx_spalten.py`, `plaene-erzeugen.py`, `plaene.test.py` | #15 |
| 2026-11-27 | Nach Test des Nutzers (Excel): Eingabefelder der Dateneingabe mit Musterwoche vorbelegt (Formel, grau-kursiv; eigene Einträge fett, per Regel `ISFORMULA`), alle gesperrten/berechneten Felder grau (Gilt-Spalten mit grauer Schrift), Hilfsspalten ausgeblendet, Spalten für „Teammentor*in“ verbreitert (auch Schicht-/Tagesplan); Firmenname im Tagebuch durch „Arbeitgeber“ ersetzt | direkt auf `main` |
| 2026-11-27 | Nach 2. Test (Excel): Eingabefelder schwarz und normal (Vorbelegung nicht mehr grau-kursiv, eigene Einträge weiter fett); gesperrte/berechnete Felder kräftiger grau `#D9D9D9` (vorher `#F2F2F2`, wirkte wie Weiß), graue Schrift `#595959` | direkt auf `main` |
| 2026-11-27 | Neuer Aufbau (Abschnitt 10a): offene **Eingabe-Datei** + privater **Stundenzettel** je Mitarbeiter, Ordner „Schichtplanung“/„Steuerung“; „krank“ → „abwesend (Gutschrift)“ (auch Hauptvorlage); Schichtplan erkennt Vertragsende über den Stundenzettel (Blatt `Werte`); neues Werkzeug `vorlagen-bauen.py`; Kettentest mit 57 Prüfpunkten | direkt auf `main` |
| 2026-11-27 | Nach Test V4: Eingabe-Datei ohne Blatt „Start“ (Jahr/Urlaubsanspruch/Soll/Vertragsmonate/057 kommen per Verknüpfung aus dem Stundenzettel, Blatt `Vorgaben`), Name in den Titeln, keine Nummern, **Reststunden zum Jahresende** (rot „noch zu leisten“, grün „zu viel geleistet“, fett); Stundenzettel heißen „Stundenzettel 01–35.xlsx“; Schicht- UND Tagesplan lesen direkt die Eingabe-Dateien (kein Speichern des Stundenzettels/Schichtplans mehr nötig); ohne Vorgabe gilt „Vertrag besteht“; Kettentest 64 Prüfpunkte | direkt auf `main` |
| 2026-11-27 | Stundenzettel: Überschriften ohne Nummern und auf allen Blättern mit Namen (Stammdaten, Tagesliste, Urlaub, Auswertung); Kettentest 68 Prüfpunkte | direkt auf `main` |
| 2026-11-27 | Anleitungs-Blätter entfernt (Eingabe-Datei, Stundenzettel, Schichtplan) und Erklärtext in der Tagesplan-Übersicht gestrichen – der Steuerer erklärt das Tool bei der Einweisung; Kettentest 70 Prüfpunkte | direkt auf `main` |
| 2026-11-27 | Nach Test V6: `#BEZUG!` in der ganzen Dateneingabe, wenn die Verknüpfung zum Stundenzettel (noch) einen Fehler liefert → alle Formeln, die Verknüpfungen auswerten, mit `IFERROR` abgesichert (Eingabe-Datei: „keine Vorgabe“, Stundenzettel: leer, Schicht-/Tagesplan: leer). Neu: `Steuerung/Inhaltsverzeichnis.xlsx` (Stundenzettel-Nr. → Name, Klick öffnet die Datei); `plaene-erzeugen.py --testnamen` (35 erfundene „Mustermann, Vorname“ mit unterschiedlichen Musterwochen). Großes Test-Paket per E-Mail; Kettentest 84 Prüfpunkte | direkt auf `main` |
| 2026-11-27 | Test des großen Pakets V7 beim Arbeitgeber abgeschlossen (35 Testpersonen, schnell genug); Beobachtungen 2–6 im Tagebuch gesammelt (u. a. Verknüpfung geht bei falscher Reihenfolge verloren), Umsetzung in der nächsten Sitzung | – |
| 2026-11-28 | Excel-Paket V8 „Vorlagen-Prinzip“ (Wunsch des Nutzers): Namen NUR im `000 Inhaltsverzeichnis` (Steuerer trägt ein, niemand wird überschrieben, Neue unten dazu); Eingabe-Dateien heißen fest „Platz 01–35“, Stundenzettel „01–35“ (seit V9: „MA 01“ / „MA 01 (Admin)“) und erkennen ihre Nummer am Dateinamen; Nutzer bekommt nur 5 Dateien und kopiert selbst; Verknüpfungen nur noch über die Sammel-Dateien (Inhaltsverzeichnis, `002 Schichtplan`); `001 MA-Tagesplan` nach Platz-Nummer; löst Beobachtung 4/5 (Kopien einer in Excel gespeicherten Vorlage). Kettentest neu (Stand V8) | direkt auf `main` |
| 2026-11-28 | Paket V9: Dateinamen „MA 01–35.xlsx“ (Mitarbeitende) und „MA 01–35 (Admin).xlsx“ (Steuerer) statt „Platz“/„Stundenzettel“; Kettentest 112 Prüfpunkte; V8 ersetzt | direkt auf `main` |
| 2026-11-28 | Änderungsmeldung zum bestehenden (bereits genehmigten) Programmeinsatzverfahren: `excel/Programmeinsatzverfahren-Aenderung.md`, per E-Mail an den Nutzer. Kernpunkt: gleiche Daten und Rechenregeln, neu sind Ablauf und Sichtbarkeit (Schicht-/Tagesplan); Kennwort zum Öffnen als Test aufgenommen | direkt auf `main` |

## 4. Offene Ideen / nächste Schritte

- Vorbelegter Urlaub: Der 26.03.2027 (Karfreitag) war als Urlaub eingetragen und zählt nicht; dadurch sind nur
  29 statt 30 Tage vorbelegt. Nutzer kann einen Tag nachtragen – oder die Vorbelegung wird angepasst.
- Vorbereitung auf ein neues Jahr (2028): Jahr umstellen, vorbelegten Urlaub pflegen, ggf. Übernahme des Übertrags.
- **Noch nicht in der App nachgezogen** (Excel ist Vorlage, siehe Grundregel in Abschnitt 2):
  - Spalte „Anmerkung“ (Büro/Homeoffice/Teammentor*in/krank) in Dateneingabe und Musterwoche (seit #15 in Excel).
  - Gesperrte Felder kräftig grau `#D9D9D9` (App nutzt noch `#F2F2F2`), Gilt-Spalten unauffällig, eigene Einträge fett (seit 27.11.2026 in Excel). Die Vorbelegung mit der Musterwoche hat die App schon.
  Nutzer hat am 27.11.2026 entschieden: **App bleibt vorerst ganz außen vor** und wird erst am Ende nachgezogen – je nachdem,
  wohin sich das Projekt entwickelt (evtl. braucht die App ganz andere Funktionen als das Geflecht Schichtplan/Stundenzettel/Tagesplan).
- Schichtplan/Tagesplan: nach erfolgreichem Mini-Test beim Arbeitgeber das volle Paket erzeugen (35 Plätze, 52 Wochen, alle Tage) und schicken.
  Bewusste Unterschiede, weil technisch nicht anders möglich: In der App wird der Mitarbeitername eingetippt (kein Dateiname);
  Excel hat zusätzlich das Blatt „Anleitung“, die App erklärt in kurzen Hinweistexten.
- **Nächster Schritt:** Nutzer testet Paket V9 (ab Montag) (Vorlagen-Prinzip) beim Arbeitgeber – Abschnitt 10b, Test-Liste V8. Offen: Beobachtung 6b (Verknüpfungen in versteckten Blättern).
- Produktivstart Excel-Paket: echte Ordner-Anordnung (Schichtplanung offen, Steuerung privat, nicht nebeneinander) einbauen – Details in Abschnitt 10, „Für den Produktivstart merken“.
- Excel-Version und App können auseinanderlaufen: Wird das Rechenmodell der App geändert, muss die Excel-Datei
  von Hand nachgezogen werden (oder bewusst nicht). Ein automatischer Vergleich App ↔ Excel fehlt noch.
- Es gibt keinen eigenen Reiter „Einstellungen“; App-Einstellungen (Installieren, Updates, Sicherung) stehen im Reiter „1. Stammdaten“.
  Falls sie dort zu viel werden: eigenen Reiter „Einstellungen“ anlegen.

## 5. Zusammenarbeit mit dem Nutzer

- Der Nutzer ist **Programmier-Laie**: einfaches Deutsch, Fachbegriffe kurz erklären, kleine nummerierte Schritte.
- Immer klar sagen, **was Claude übernimmt** und **was der Nutzer selbst tun muss**.
- Claude übernimmt alles, was technisch möglich ist – inklusive Branch, Pull Request anlegen und Merge.
- Hardware des Nutzers: nur ein **Samsung-Smartphone (Chrome)** und ein **Raspberry Pi 400**. Anleitungen darauf abstimmen.
- Änderungen vor dem Live-Schalten automatisiert im Browser testen und das Ergebnis ehrlich berichten.

## 6. Modell-Hinweis

Wenn Claude nicht auf Opus läuft (z. B. Fable) und eine Aufgabe abgeschlossen ist, am Ende der Antwort erinnern:
„Tipp: Du kannst jetzt mit `/model claude-opus-5-5` zurückstellen.“

## 7. Pflege dieses Tagebuchs

Nach jeder größeren Änderung dieses Tagebuch aktualisieren (Verlauf, Entscheidungen, offene Ideen).

## 8. Hilfsprogramme und Tests gehören ins Projekt

Regel:
1. Automatische Prüfungen (Tests) und Hilfsprogramme, die über eine einzelne Aufgabe hinaus nützlich sind, werden im
   Projekt gespeichert (Ordner `werkzeuge/`) – nicht nur im Arbeitsbereich der Sitzung, der beim Archivieren verloren geht.
2. Hier im Tagebuch steht, wie man sie startet und was sie prüfen.
3. Nach jeder Code-Änderung die vorhandenen Prüfungen laufen lassen; bei neuen Funktionen passende Prüfpunkte ergänzen.
4. Gibt es noch keine automatische Prüfung, bei der nächsten Code-Änderung eine einfache anlegen.
5. Was Claude nicht selbst testen kann (z. B. das echte Handy, den Raspberry Pi), unten als Test für den Nutzer vermerken.

Vorhandene Werkzeuge (benötigen Node.js und Playwright mit Chromium; in der Claude-Cloud-Umgebung vorinstalliert):

| Befehl | Was es tut |
|---|---|
| `bash werkzeuge/alle-tests.sh` | Startet alle `*.test.js` nacheinander |
| `node werkzeuge/vergleich-und-bedienung.test.js [referenz.html]` | Vergleicht alle berechneten Werte (Stammdaten, Urlaub, Musterwoche, jeder einzelne Tag, Auswertung) zwischen einer Referenz (Standard: `index.html` aus `origin/main`) und der aktuellen `index.html` – mit Standarddaten und einem Beispieldatensatz. Danach Bedienung: Tag ändern, Speichern, Urlaubsgrenze, Zeitformat, Export/Import. |
| `node werkzeuge/sicherung.test.js` | Sicherung als Text: kopieren, Daten ändern, wiederherstellen, ungültigen Text abweisen; macht einen Handy-Screenshot |
| `node werkzeuge/updates.test.js` | Kopiert die App in einen Testordner, startet dort einen Webserver und „veröffentlicht“ neue Versionen: prüft Versionsanzeige, Knopf „Jetzt nach Updates suchen“, automatisches Neuladen, Warten während einer Bearbeitung, Daten bleiben erhalten, automatische Suche ohne Knopf |
| `node werkzeuge/ausscheiden.test.js` | Ausscheiden unter dem Jahr in der App: anteiliger Urlaub (6/7/5/1 Monate, 30 und 31 Tage), Urlaubsstunden und Soll **gegen die Werte der Excel-Datei**, „kein Vertrag“ in der Dateneingabe, Urlaubs-Karten rot/grün, Urlaubsgrenze. Endet mit „ERGEBNIS: alles in Ordnung“. |
| `node werkzeuge/name-und-aussehen.test.js` | Angleichung an Excel: Mitarbeitername (Hinweis ohne Komma, große Anzeige, bleibt nach Neuladen, steckt in der Sicherung), Farben gelb/grau/hellblau, Uhrzeiten 11:00–17:00, Reiter-/Überschriften; macht Handy-Screenshots. Endet mit „ERGEBNIS: alles in Ordnung“. |
| `node werkzeuge/offline-app.test.js` | Startet einen lokalen Webserver und prüft Manifest, Symbole, Offline-Helfer und den Start ohne Internet |
| `python3 werkzeuge/excel-pruefen.py [datei.xlsx] [--jahresbasis WERT] [--anspruch TAGE] [--ab-monat N] [--urlaub K]` | Rechnet die Excel-Datei mit LibreOffice neu und gibt Soll, Regel-Werte, Urlaub und Auswertung aus. Zum Vergleich alte und neue Datei prüfen und die Ausgaben mit `diff` vergleichen. `--ab-monat 7` = Ausscheiden ab Juli (Erwartung bei 30 Tagen: 15 Tage, 117:00:00), `--urlaub 15` trägt 15 Urlaubstage im Januar ein. |
| `python3 werkzeuge/plaene.test.py` | Ganze Kette (Abschnitt 10b, Stand V8 „Vorlagen-Prinzip“): 4 Plätze als Kopien je einer Vorlage; Verknüpfungen, die Fehler liefern (kein `#BEZUG!`); falscher Dateiname → Warnung; Nummer aus dem Dateinamen; Inhaltsverzeichnis sammelt Vorgaben; Schichtplan (Namen aus dem Inhaltsverzeichnis, nach Nachname) und Tagesplan (nach Platz-Nummer); Stundenzettel aus dem Schichtplan (muss **genau wie die Hauptvorlage** rechnen, 057-Wert); Eingabe holt Name + Vorgaben (Reststunden = Stundenzettel sekundengenau, Urlaubswarnung); Vertragsende. Aktualisieren wie Excel (gespeicherte Werte, auch Bereiche für INDEX). 112 Prüfpunkte, Ende „ERGEBNIS: alles in Ordnung“ (ca. 3 Min). |
| `python3 werkzeuge/vorlagen-bauen.py` | Kein Test: baut aus der Hauptvorlage `excel/vorlagen/Eingabe.xlsx` und `Stundenzettel-Vorlage.xlsx` neu – **nach jeder Änderung an der Hauptvorlage ausführen**, danach `plaene.test.py` |
| `python3 werkzeuge/plaene-erzeugen.py --ordner ORDNER [...]` | Kein Test: erzeugt das Paket V9 – `MA 01.xlsx`, `MA 01 (Admin).xlsx`, `000 Inhaltsverzeichnis.xlsx`, `001 MA-Tagesplan.xlsx`, `002 Schichtplan.xlsx` (Optionen siehe Kopf der Datei; Versand-Paket: `--plaetze 35 --testnamen` bzw. ohne `--testnamen` = leeres Inhaltsverzeichnis; `--kopien` legt zusätzlich alle Kopien 02 … an, nur für Tests) |
| `werkzeuge/xlsx_spalten.py` | Kein Test: Hilfsfunktion `spalten_einfuegen(wb, blatt, idx)` – fügt Spalten ein und zieht alle Bezüge (Formeln, Namen, Formatierung, Prüfungen, Verbindungen) mit; danach immer `excel-pruefen.py` alt/neu vergleichen |
| `node werkzeuge/symbole-erzeugen.js` | Kein Test: erzeugt aus `icons/icon.svg` alle PNG-Symbole neu (nur nach Änderung am Symbol) |

Ergebnis lesen: Die Prüfungen geben ihre Ergebnisse als Text aus und brechen nicht von selbst mit „Fehlgeschlagen“ ab.
In Ordnung ist es, wenn beim Vergleich überall `IDENTISCH` steht (Ausnahme: gewollte Textänderungen),
`errors: []` erscheint und die Werte mit „erwartet …“ übereinstimmen. Bei gewollten Änderungen am Rechenmodell
weichen die Zahlen natürlich ab – dann die Unterschiede einzeln prüfen und dem Nutzer berichten.

Tests für den Nutzer (kann Claude nicht selbst prüfen):
- Auf dem Samsung-Handy (Chrome): App öffnet sich vom Startbildschirm, Name „Arbeitszeiterfassung“ unter dem Symbol,
  App startet im Flugmodus, „Sicherung als Text kopieren“ und „Sicherung teilen“ funktionieren.
- Auf dem Raspberry Pi 400 (Chromium): App lässt sich installieren, Sicherung vom Handy per Text einspielen.
- Neue Farben (Blau/Orange) erscheinen nach dem Öffnen mit Internet; ggf. App einmal schließen und neu öffnen.
  Das neue Symbol auf dem Startbildschirm kann erst nach Tagen oder nach Neu-Installation erscheinen.
- Ausscheiden in der App: Jahres-Basis Juli–Dezember auf 0 → unter „Urlaubsanspruch volles Jahr“ steht „Anteilig 6 von 12 Monaten: 15 Tage“;
  Dateneingabe Juli–Dezember grau „kein Vertrag“; Reiter Urlaub rot, bei genau 15 eingetragenen Tagen grün.
- Mitarbeitername in der App: Stammdaten → „Ändern“ → „Mustermann, Max“ eintragen → „Speichern“: Überschrift und blaues Feld zeigen den Namen.
  Uhrzeit-Auswahl zeigt nur 11:00–17:00; Eingabefelder im „Ändern“-Modus gelb, sonst grau; Urlaubstage hellblau.
- Updates: Unter Stammdaten → „App-Updates“ steht eine Version (z. B. v4); „Jetzt nach Updates suchen“ meldet
  „neueste Version“. Nach dem nächsten Update öffnet sich die App kurz neu und zeigt „gerade aktualisiert“.

## 9. Excel-Version

**Datei:** `excel/Arbeitszeiterfassung.xlsx` (am 23.11.2026 von Claude erstellt – in welcher Sitzung, ist nicht mehr bekannt;
der Nutzer hat sie am 27.11.2026 vom Handy hochgeladen). Sie ist eine **leere Vorlage ohne persönliche Daten** und darf deshalb
im öffentlichen Repository liegen. Eigene ausgefüllte Kopien gehören **nicht** ins Repository.

Zweck: **Vorlage für die Logik** (siehe Grundregel in Abschnitt 2); die App rechnet genauso. Die Datei liegt beim Nutzer zentral
im Abteilungsordner für alle Mitarbeiter. Sie ist unabhängig von der App – Daten werden **nicht** zwischen App und Excel ausgetauscht
(keine Synchronisation, ausdrücklich nicht gewünscht).

Aufbau (Blätter):

| Blatt | Inhalt |
|---|---|
| Anleitung | Bedienung, Farben (Gelb = eintragen, Grau = rechnet Excel), Rechenregeln, Sichern, neues Jahr |
| Stammdaten | Mitarbeiter (aus dem Dateinamen), Jahr (2027), Urlaubsanspruch, Übertrag (Plus/Minus), vereinbarte Überstunden, unbezahlter Urlaub, Jahres-Basis je Monat |
| Musterwoche | Mo–Fr Start (Klappmenü Frei oder 11:00–17:00, 15-Min.-Takt) und Pause, Ende fix 19:00 |
| Dateneingabe | alle Werktage des Jahres, hessische Feiertage vorbelegt, Klappmenü für Urlaub/Feiertag/Frei und Arbeitsbeginn 11:00–17:00 (15-Min.-Takt) |
| Urlaub | Anspruch, genommen, Rest, je Monat; Warnung bei zu viel Urlaub |
| Auswertung | Monatswerte Plan, Eingabe „Stundenabrechnung 057“, Soll/Ist/Differenz |
| Listen | Hilfsblatt: Klappmenü-Werte, Osterberechnung (Gauß), Feiertage |

Wichtig:
- Gleiches Rechenmodell wie die App (Regel-Tag = Jahresstunden / 52 / 5; Urlaub vom Soll abgezogen; Feiertage zählen auf das Ist).
  **Ausscheiden/Eintritt unter dem Jahr (Wunsch des Nutzers, 27.11.2026; die App rechnet seit #12 genauso):**
  - Monat mit Jahres-Basis 0 = kein Vertrag: Dateneingabe zeigt „kein Vertrag“ (grau), Netto 0, Urlaub dort zählt nicht.
  - Urlaub anteilig: Stammdaten G35 = Anspruch volles Jahr (B5), G36 = Monate mit Jahres-Basis > 0, G39 = G35 × G36 ÷ 12,
    G37 (`UrlaubAnspruchTage`) = gesetzlich gerundet (BUrlG § 5: ab halbem Tag aufrunden, kleinere Bruchteile bleiben, auf 2 Stellen).
    Beispiele bei 30 Tagen: 6 Monate → 15, 7 → 18 (17,5), 5 → 13 (12,5); bei 31 Tagen und 7 Monaten → 18,08.
  - Urlaubsstunden = Tage × Ø Regel-Tag **nur der Vertragsmonate** (G38), sonst würde doppelt gekürzt.
  - Stammdaten C5 zeigt „Anteilig x von 12 Monaten: y Tage“.
  - Urlaubsblatt A4:B6: grün, wenn genommen = zustehend; sonst rot. C6 zeigt „Noch x Tage offen“ bzw. Warnung.
- Zeiten als Stunden:Minuten eingeben (`12:30`), nicht als Kommazahl – **auch die Jahres-Basis** (`2028:00`, Anzeige `2028:00:00`).
  Intern ist sie ein Excel-Zeitwert (Tage); die Prüfung lässt höchstens 5000 Std. zu, damit `2028` ohne Doppelpunkt abgewiesen wird.
- Zahlenformat `0.##` nicht verwenden: es zeigt bei ganzen Zahlen ein Komma („30,“). Stattdessen `General`
  (bzw. `General" Tage"`). In Formeln Tage mit `&""` statt `TEXT(…;"0.##")` einbauen (sonst „031“ in deutschem Excel).
- **Mitarbeiter** (Stammdaten B3, Name `Mitarbeiter`): kommt aus dem Dateinamen über `CELL("filename")` (Hilfszellen G3 = Pfad,
  H3 = Dateiname). Datei als „Nachname, Vorname.xlsx“ speichern. Ohne „, “ im Namen erscheint in D3 ein Hinweis.
  Anzeige: Überschrift A1 = „1. Stammdaten – <Name>“; B3:C3 verbunden, Mets-Blau, 14 pt fett weiß, oranger Rahmen.
  `CELL("filename")` funktioniert in Excel am PC und in LibreOffice, laut Microsoft **nicht** in Excel im Browser/auf dem Handy
  und nicht in Google Tabellen – dann bleibt B3 leer und D3 zeigt „Name nicht lesbar“. Eine neue, nie gespeicherte Datei hat noch keinen Namen.
- Blätter sind **ohne Passwort** geschützt (nur gegen versehentliches Löschen von Formeln), keine Makros.
- **Anmerkung** (seit #15): Dateneingabe F = Eingabe (Klappmenü `ListeAnmTag`: Büro, Homeoffice, Teammentor*in, krank),
  I = gilt (Eingabe, sonst Musterwoche D über `MusterAnm`; nur an Arbeitstagen). Musterwoche D (`ListeAnmMuster`: Büro, Homeoffice,
  Teammentor*in). „krank“: Beginn/Pause bleiben wie geplant und werden gutgeschrieben – die Rechnung ändert sich nicht.
  Spalten wurden mit `xlsx_spalten.py` eingefügt: Dateneingabe G = Gilt Start, H = Gilt Pause, J = Netto, K–M versteckte Helfer.
- **Vorbelegung** (seit 27.11.2026): Dateneingabe D/E/F enthalten Formeln mit den Werten der Musterwoche (D: Feiertag oder
  `MusterStart`, E: `MusterPause`, F: `MusterAnm`; leer bei „kein Vertrag“). Der Mitarbeiter überschreibt per Klappmenü. Bedingte Formatierung
  (Priorität 2, vor den anderen Regeln): eigener Eintrag (keine Formel) → fett (`_xlfn.ISFORMULA`, ab Excel 2013); vorbelegte Werte
  schwarz und normal (Wunsch des Nutzers: Eingabefelder sollen nicht „gesperrt“ aussehen). Wird ein Feld gelöscht, rechnet „Gilt“
  trotzdem mit der Musterwoche. Zurücknehmen: ein nicht fettes Feld derselben Spalte kopieren.
- **Farben**: Gelb = eintragen, Grau `#D9D9D9` = gesperrt/berechnet (alle Blätter; Dateneingabe A–C, G–J; Gilt-Spalten mit Schrift `#595959`).
  `#F2F2F2` ist auf Bildschirmen zu hell (wirkt weiß) – nicht mehr verwenden.
  Ausgeblendete Hilfsspalten: Stammdaten G/H, Musterwoche F–I, Dateneingabe K–M, Urlaub D, Auswertung F/G.
- „krank“ bleibt als Bezeichnung (Entscheidung des Nutzers, 27.11.2026) – Hinweis: Kollegen können die Mitarbeiter-Dateien lesen.
- **Blatt „Export“** (versteckt, geschützt): Quelle für den Schichtplan. B1 = Name (direkt `=Stammdaten!$B$3`, KEIN Namensbezug –
  LibreOffice verwechselt Namen zwischen verknüpften Dateien), D1 = aktiv (0 bei leer/„Mustermann, Max…“), Zeilen 4–265 = Werktag 1–262
  (A Datum, B Beginn oder Urlaub/Feiertag/Frei/„abwesend“, C Pause, D Anmerkung ohne „krank“). Name: `%2C` wird zu „,“ dekodiert.
- Für ein neues Jahr: Kopie speichern, Jahr in „Stammdaten“ ändern, alte Eingaben löschen.
- Bei Änderungen an der Datei: mit `openpyxl` (Python) bearbeiten, danach die Formeln mit LibreOffice neu berechnen lassen
  und die Ergebnisse mit der App vergleichen (gleiche Eingaben → gleiche Zahlen).
- Download-Link (nach dem Merge): `https://marchens1977.github.io/Arbeitszeiterfassung/excel/Arbeitszeiterfassung.xlsx`

Tests für den Nutzer (Excel):
- [ ] Datei über den Download-Link am Handy öffnen (z. B. Excel-App oder Google Tabellen).
- [ ] In der Musterwoche einen Wert ändern → Dateneingabe und Auswertung ändern sich mit.
- [ ] Stammdaten: Urlaubsanspruch ohne Komma, Jahres-Basis als `2028:00:00`; Eingabe `2028:00` wird angenommen, `2028` abgewiesen.
- [ ] Dateneingabe: Klappmenü „Start / Status“ zeigt Urlaub, Feiertag, Frei und 11:00 bis 17:00.
- [ ] Musterwoche: Klappmenü „Start / Status“ zeigt Frei und 11:00 bis 17:00.
- [ ] Excel am PC: Datei als „Mustermann, Max.xlsx“ speichern, neu öffnen → Stammdaten zeigt oben „1. Stammdaten – Mustermann, Max“ und das blaue Feld „Mitarbeiter“.
- [ ] Ausscheiden: Jahres-Basis Juli–Dezember auf `0:00` → Stammdaten zeigt „Anteilig 6 von 12 Monaten: 15 Tage“, Urlaub 15 Tage / 117:00:00, Juli–Dezember in der Dateneingabe grau „kein Vertrag“.
- [ ] Urlaubsblatt: Block oben rot, solange genommener ≠ zustehender Urlaub; grün, wenn genau gleich.
- [ ] Vorbelegung: gelbe Felder der Dateneingabe zeigen die Musterwoche (schwarz, normal); eigener Eintrag (z. B. Urlaub) erscheint fett; nicht änderbare Felder deutlich grau.
- [ ] Anmerkung: Musterwoche Freitag „Homeoffice“ → in der Dateneingabe steht freitags „Homeoffice“ (gelb); einzelner Tag „Büro“ hebt es auf; „krank“ ändert die Stunden nicht.
- [ ] Gleiche Eingaben wie in der App → gleiche Jahresbilanz (stichprobenartig vergleichen).

## 10. Schichtplan und MA-Tagesplan (Excel, ohne Makros)

Wunsch des Nutzers (27.11.2026): Steuerer sieht automatisch, wer wann da ist. Vorlagen kamen per Mail
(„Schichtplan.xlsx“ mit Wochenblatt „04.01.2027“, „MA-Tagesplan.xlsx“ mit Kampagnen-Spalten).
**Makros sind beim Arbeitgeber nicht erlaubt** → nur Excel-Verknüpfungen (Daten → Links bearbeiten → Werte aktualisieren).

Entscheidungen des Nutzers:
- Alle Dateien in EINEM Ordner (Abteilungsordner); Mitarbeiter dürfen alles ansehen.
- Mitarbeiter-Dateien heißen „Nachname, Vorname.xlsx“ (Name kommt aus dem Dateinamen). Bei Wechsel: Datei umbenennen
  und NUR im Schichtplan die Verknüpfung umstellen (Links bearbeiten → Quelle ändern). Der Tagesplan holt alles aus dem Schichtplan.
- 35 Plätze; freie Plätze sind Platzhalter-Dateien „Mustermann, Max NN.xlsx“ (werden nicht angezeigt).
- Tagesplan: ALLE Mitarbeiter fest nach Nachname sortiert; wer nicht da ist, ist ausgegraut (Grund steht in „Anm.“).
  So verrutschen die handschriftlichen Kampagnen-Einträge nicht. (Achtung: Ändert sich die Mitarbeiterliste, ändert sich die Sortierung
  auf allen Tagesblättern – dann Kampagnen-Einträge prüfen.)
- Krank: im Schicht-/Tagesplan nur „abwesend“ (Grund darf nicht kommuniziert werden). Urlaub/Feiertag werden angezeigt (grün).
- Anmerkungen: Homeoffice (gelb, im Tagesplan „HO“), Teammentor*in (hellblau).
- Schutz: „Kennwort zum Ändern“ (nicht „Kennwort zum Öffnen“ – das verschlüsselt und bricht die Verknüpfung). Bester Schutz: Ordnerrechte der IT.

Aufbau:
- `Schichtplan.xlsx`: Anleitung · Mitarbeiter (Plätze, Name, aktiv, Sortierung nach Nachname) · ein Blatt je Woche („04.01.2027“ …, 52 Wochen) ·
  versteckt „Daten“ (je Platz 3 Spalten Beginn/Pause/Anm, Zeilen 4–265 = Werktag; Zeile 1 Name, Zeile 2 aktiv; je Platz eine Verknüpfung
  auf `Export` der Mitarbeiter-Datei).
- `MA-Tagesplan.xlsx`: Übersicht (Links zu jedem Tag) · ein Blatt je Arbeitstag ohne Feiertage · versteckt „Daten“ und „Mitarbeiter“
  (eine einzige Verknüpfung auf `Schichtplan.xlsx`). Kampagnen-Felder gelb und entsperrt, Rest gesperrt (ohne Passwort).
- Verknüpfungen werden von `plaene-erzeugen.py` als echte Excel-Verknüpfungen (externalLink, relativer Pfad) geschrieben und enthalten
  NUR einfache Bezüge `=[n]Blatt!A1` (keine WENN drumherum) – Bereinigung passiert lokal.
- Excel liest bei geschlossenen Dateien die **gespeicherten** Werte: Mitarbeiter müssen ihre Datei speichern; der Steuerer muss den
  Schichtplan nach dem Aktualisieren speichern, bevor der Tagesplan aktualisiert wird. Neue (von Claude erzeugte) Dateien haben noch keine
  gespeicherten Werte → einmal in Excel öffnen und speichern.

Test beim Arbeitgeber (Mini-Paket: „Beispiel, Anna.xlsx“, „Test, Ben.xlsx“, „Mustermann, Max 03.xlsx“, Schichtplan 1 Woche, Tagesplan 5 Tage):
- [ ] Alle 5 Dateien in einen neuen Ordner auf dem Firmenlaufwerk legen.
- [ ] Anna und Ben je einmal öffnen und speichern (Stammdaten zeigt den Namen).
- [ ] Schichtplan öffnen → „Inhalt aktivieren“ bzw. Daten → Links bearbeiten → Werte aktualisieren → Werte erscheinen
  (Anna Mo 11:00 Homeoffice gelb, Di Urlaub grün; Ben Mi „abwesend“ grau, Do Teammentor*in, Fr Frei) → speichern.
- [ ] MA-Tagesplan öffnen → aktualisieren → 05.01.: Anna grau „Urlaub“, Ben 13:00; Kampagnen-Feld eintragbar.
- [ ] Ben in der Datei auf „Urlaub“ am 04.01. setzen, speichern → Schichtplan aktualisieren, speichern → Tagesplan aktualisieren → Ben grau.
- [ ] „Kennwort zum Ändern“ an einer Datei setzen → Schichtplan aktualisiert trotzdem.

Test Mini-Paket V5 (zwei Ordner nebeneinander: „Schichtplanung“ und „Steuerung“):
- [ ] Eingabe-Dateien öffnen, speichern → Schichtplan und Tagesplan direkt aktualisieren → gefüllt (ohne Stundenzettel anzufassen).
- [ ] Titel mit Namen, keine Nummern, kein Blatt „Start“; oben „noch keine Vorgabe vom Steuerer“.
- [ ] Stundenzettel 02 (Ben) öffnen → aktualisieren → speichern; dann „Test, Ben“ öffnen → aktualisieren → speichern → oben Reststunden (rot).
- [ ] Ausscheiden: Stundenzettel 02 Januar 0:00 → speichern → „Test, Ben“ aktualisieren, speichern → Schichtplan aktualisieren → Ben im Januar leer.

Test Mini-Paket V4 (erledigt; Ergebnis führte zu V5):
- [ ] Dateien richtig verteilen (alle „Stundenzettel …“ nach „Steuerung“, Rest nach „Schichtplanung“).
- [ ] Eingabe-Dateien (Anna, Ben, Mustermann 03) öffnen, speichern; Blatt „Start“ zeigt Name, keine Vertragsdaten.
- [ ] Stundenzettel öffnen → aktualisieren → Tagesliste zeigt Bens Einträge, Stammdaten zeigen Name → speichern (alle drei).
- [ ] Schichtplan öffnen → aktualisieren (Eingabe-Dateien UND Stundenzettel) → Werte wie bisher → speichern.
- [ ] Tagesplan öffnen → aktualisieren → wie bisher.
- [ ] In Bens Stundenzettel Jahres-Basis Januar auf 0:00 (Februar 2028:00) → speichern → Schichtplan aktualisieren → Ben im Januar ohne Zeiten.
- [ ] In Annas Eingabe-Datei Urlaubsanspruch auf 0 → oben in der Dateneingabe rote Warnung.

### Übergabe für die nächste Sitzung (Stand 27.11.2026)

**>>> HIER WEITERMACHEN (Stand 28.11.2026): Beobachtungen 2, 4, 5, 6a sind mit Paket V8 („Vorlagen-Prinzip“, Abschnitt 10b) umgesetzt. Paket V9 (5 Dateien, Namen „MA 01“ / „MA 01 (Admin)“; V8 ist überholt) ist per E-Mail beim Nutzer; er testet es ab Montag beim Arbeitgeber. Ablauf wie immer: Beobachtungen nur sammeln, bei „Test fertig“ bewerten. Offen: 6b (Verknüpfungen in versteckten Blättern), Ordner-Anordnung für den Produktivstart.**

Ergebnis Test V7 (35 Testpersonen, Excel beim Arbeitgeber):
- Funktioniert: Eingabe-Datei → Stundenzettel → Inhaltsverzeichnis (bei richtiger Reihenfolge); Schichtplan und Tagesplan zeigen alle 35
  korrekt; Öffnen/Aktualisieren **unter einer Minute** – die Größe ist kein Problem.
- Richtige Reihenfolge beim ersten Mal (so dem Nutzer erklärt): 1. alle Eingabe-Dateien öffnen, NUR „Bearbeitung aktivieren“ (NICHT
  „Inhalte aktivieren“), speichern, schließen → 2. Stundenzettel öffnen, Bearbeitung + Inhalte aktivieren/„Aktualisieren“, Name prüfen,
  speichern → 3. Inhaltsverzeichnis, Schichtplan, Tagesplan aktualisieren → 4. (optional) Eingabe-Dateien nochmals mit Inhalten
  aktivieren → Reststunden. Nicht getestet: Reststunden in der Eingabe-Datei, Änderung durchreichen (Urlaub → Schichtplan).

Gesammelte Beobachtungen (offen, zur Umsetzung):
1. `#BEZUG!` beim ersten Öffnen – **erledigt** (IFERROR, V7).
2. Schichtplan, Tagesplan, Inhaltsverzeichnis dürfen umbenannt werden („001 MA-Tagesplan“, „002 Schichtplan“, ggf. „000 Inhaltsverzeichnis“)
   → beim Produktivpaket gleich so erzeugen (vorher kurz nachfragen).
3. Ordner liegen im Produktivbetrieb nicht nebeneinander → siehe „Für den Produktivstart merken“ unten.
4. Erstes Füllen ist mühsam (70 Dateien der Reihe nach öffnen/speichern).
5. **Wichtigster Punkt:** Excel ersetzt eine Verknüpfung DAUERHAFT durch `=#BEZUG!`, wenn aktualisiert wird, solange die Quelldatei noch nie in
   Excel gespeichert wurde (von openpyxl erzeugt, ohne gespeicherte Werte). Passiert: Stundenzettel 01 (vor der Eingabe-Datei aktualisiert)
   und danach Zeile 01 im Inhaltsverzeichnis. Folge: „Workbook Links“ grau (Excel kennt keine Verknüpfung mehr), Datei unbrauchbar → frisch
   aus der E-Mail nehmen. Beim Öffnen ist die Verknüpfung noch da (Frage „enthält Verknüpfungen…“), sie geht erst beim Aktualisieren verloren.
   Lösungsidee (4+5): Paket **vorgespeichert** ausliefern (z. B. per LibreOffice alle Dateien in richtiger Reihenfolge rechnen und speichern,
   sodass jede Datei gespeicherte Werte und gefüllte Verknüpfungs-Zwischenspeicher hat). Achtung: `CELL("filename")` (Name aus Dateiname)
   darf dabei NICHT als fester Text landen – der Test setzt Listen!O2 per Hand, das wäre für die Auslieferung falsch. Danach unbedingt
   vom Nutzer in Excel prüfen lassen.
6. a) Stundenzettel über den Klick-Link im Inhaltsverzeichnis geöffnet → gelbe Leiste „Aktualisierung nicht möglich … Links der Arbeitsmappe
   verwalten“ (Werte bleiben, nichts kaputt); per Doppelklick im Explorer klappt alles. → Link anders bauen oder Hinweis „im Explorer öffnen“.
   b) Excels Bereich „Links zu Arbeitsmappen“ zeigt ein gelbes Dreieck „Link befindet sich auf einem ausgeblendeten Blatt“ (Verknüpfungen
   liegen absichtlich in versteckten Blättern `Daten`/`Vorgaben`) → prüfen, ob Blätter sichtbar+gesperrt sein sollen (Nutzer entscheiden lassen).
- Bedienhinweise aus dem Test: In diesem Excel heißt der Knopf **Daten → „Workbook Links“** („Links zu Arbeitsmappen“); „Alle aktualisieren“
  reicht für Verknüpfungen evtl. nicht. Arbeitsmappenschutz (Struktur) ist an; er war NICHT der Grund für den grauen Knopf.

**Stand vor Test V7 (27.11.2026, nach Test V6):** Mini-Test bestanden. Großes Test-Paket (35 erfundene „Mustermann, Vorname“, 52 Wochen,
alle Arbeitstage, `--plaetze 35 --testnamen`) per E-Mail beim Nutzer – er testet damit die große Anzahl. Stundenzettel bleiben nummeriert
(Entscheidung des Nutzers: Variante „Inhaltsverzeichnis“ statt Namen im Dateinamen, weil sonst beim Mitarbeiterwechsel mehr umzustellen wäre).
Regel: Mitarbeiternamen in Testpaketen immer „Mustermann, <Fantasievorname>“ (nie mit „Max“ beginnen – „Mustermann, Max…“ = freier Platz).
Wichtig für Excel: Von Claude erzeugte Dateien haben keine gespeicherten Werte. Reihenfolge beim ersten Mal: Eingabe-Dateien öffnen+speichern →
Stundenzettel aktualisieren+speichern → Inhaltsverzeichnis aktualisieren. Liefert eine Verknüpfung trotzdem Fehler, zeigen die Dateien jetzt
„keine Vorgabe“/leer statt `#BEZUG!`.

**Versand per E-Mail:** immer an die Arbeits-Adresse des Nutzers (wie bisher, steht im Gmail-Postausgang – NICHT hier eintragen), als
einzelne `.xlsx`-Dateien, **keine ZIP-Dateien** (Wunsch des Nutzers). Grenzen des Gmail-Werkzeugs: höchstens **10 Anhänge** und ca. **3 MB**
je E-Mail → großes Paket = 9 E-Mails (1: Schicht-/Tagesplan, 2–5: Eingabe-Dateien, 6–9: Inhaltsverzeichnis + Stundenzettel).

**Für den Produktivstart merken (Fragen des Nutzers, 27.11.2026):**
- **Ordner liegen später NICHT nebeneinander:** „Schichtplanung“ kommt in einen Ordner für alle Mitarbeiter, „Steuerung“ in einen Unterordner,
  auf den nur der Steuerer zugreifen kann (gleiches Laufwerk). Die Verknüpfungen sind relativ und setzen bisher fest „nebeneinander“ voraus
  (`../Steuerung/…`, `../Schichtplanung/…` in `plaene-erzeugen.py`, Konstanten `OFFEN`/`PRIVAT`). → Vor dem Produktivpaket den echten Weg
  beider Ordner erfragen (Laufwerk, Ordner ab der Abzweigung; nur Struktur, keine persönlichen Daten ins Repository) und `plaene-erzeugen.py`
  um eine Option für den relativen Weg zwischen beiden Ordnern erweitern; Kettentest anpassen. Betroffen: Eingabe-Datei ↔ Stundenzettel und
  Inhaltsverzeichnis → Stundenzettel (liegen beide in „Steuerung“, unkritisch). Schicht-/Tagesplan liegen bei den Eingabe-Dateien (unkritisch).
  Danach darf die Anordnung nicht mehr geändert werden.
- **Umbenennen erlaubt:** Schichtplan, MA-Tagesplan und Inhaltsverzeichnis (auf sie zeigt keine Verknüpfung). Wunsch: „001 MA-Tagesplan“,
  „002 Schichtplan“ (ggf. „000 Inhaltsverzeichnis“), damit sie im Explorer oben stehen – beim Produktivpaket gleich so erzeugen (nachfragen).
  NICHT umbenennen: Eingabe-Dateien, Stundenzettel, die Ordner.
- **Echte Namen (Frage des Nutzers, 27.11.2026):** Der Nutzer will KEINE echten Namen an Claude/Anthropic geben. Plan: Produktivpaket mit
  Platzhaltern „Mustermann, Max 01…35“ (gelten als freier Platz, erscheinen nicht im Plan); der Nutzer benennt selbst um. Empfohlener Weg
  (noch in Excel zu testen): Schichtplan, MA-Tagesplan und den passenden Stundenzettel öffnen, dann die Platzhalter-Eingabe-Datei öffnen und
  per **„Speichern unter“** als „Nachname, Vorname.xlsx“ speichern → Excel soll die Verknüpfungen in den geöffneten Dateien automatisch
  mitziehen → diese speichern, alte Platzhalter-Datei löschen. Alternative: im Explorer umbenennen und in Stundenzettel, Schichtplan und
  Tagesplan je „Workbook Links → … → Quelle ändern“. Beides beim nächsten Test mit einer Testperson ausprobieren und eine Kurzanleitung schreiben.
- **Ordner verschieben:** nur beide zusammen und mit gleicher Anordnung zueinander; danach in einer Eingabe-Datei „Links bearbeiten → Status“ prüfen.

Ältere Übergabe:

- **Mini-Paket V6** (alle Änderungen bis hier, 3 Plätze, 1 Woche, 5 Tage) ist per E-Mail beim Nutzer; er testet es in der NÄCHSTEN Sitzung
  in seinem Arbeitsumfeld (zwei Ordner „Schichtplanung“/„Steuerung“). Erzeugen: `python3 werkzeuge/plaene-erzeugen.py --ordner X
  --plaetze 3 --wochen 1 --tage 5 --mitarbeiter "Beispiel, Anna" "Test, Ben" --beispieldaten` (vorher `vorlagen-bauen.py`).
- Ablauf in der nächsten Sitzung: Nutzer meldet Beobachtungen nacheinander → Claude sammelt nur (nummeriert) → erst bei
  „Test fertig“ bewerten und umsetzen → `plaene.test.py` → in `main` → Mini-Paket per E-Mail (Adresse kennt der Nutzer / steht in seinen
  Einstellungen, NICHT hier eintragen).
- Danach: **volles Paket** erzeugen (35 Plätze, 52 Wochen, alle Arbeitstage: `--plaetze 35` ohne `--wochen/--tage`, Namen der echten
  Mitarbeiter NICHT ins Repository) und schicken.
- Offen: Kurze Hinweiszeilen unter den Blatt-Überschriften (z. B. Dateneingabe A2) – Nutzer wurde gefragt, ob sie auch weg sollen; noch
  keine Antwort (bisher behalten). App bleibt außen vor (siehe Abschnitt 4).
- Für LibreOffice-Tests: eigenes Profil mit `BlockUntrustedRefererLinks=false` und `MacroSecurityLevel=0` (macht `plaene.test.py` selbst);
  `pkill -x soffice.bin` statt `pkill -f` (sonst beendet man die eigene Shell).

### 10b. Stand V8: Vorlagen-Prinzip (Entscheidungen vom 28.11.2026) – geht vor, wo 10/10a anders lauten

Wunsch des Nutzers: **eine Quelle für die Namen** und **keine großen Pakete** mehr (nur je eine Vorlage, selbst kopieren).
- **Namen nur im `000 Inhaltsverzeichnis.xlsx`** (Ordner „Steuerung“, Spalte B gelb). Im laufenden Jahr wird **niemand gelöscht oder
  überschrieben**, auch nach dem Ausscheiden nicht (Ausscheiden = Jahres-Basis 0 im Stundenzettel); Neue kommen in die nächste freie Zeile.
  Aktuell 22 Mitarbeitende, 35 Plätze reichen. Echte Namen trägt der Nutzer selbst ein – sie kommen nie zu Claude.
- **Feste Dateinamen** (nie umbenennen, sonst brechen Verknüpfungen): `Schichtplanung/MA 01.xlsx … MA 35.xlsx`, `001 MA-Tagesplan.xlsx`,
  `002 Schichtplan.xlsx`; `Steuerung/MA 01 (Admin).xlsx … MA 35 (Admin).xlsx`, `000 Inhaltsverzeichnis.xlsx` (Nummern vorne: Wunsch des Nutzers). **Dateinamen „MA 01“ / „MA 01 (Admin)“ seit V9 (Wunsch des Nutzers, 28.11.2026):** nicht zweimal derselbe Name, weil Excel zwei gleichnamige Dateien nicht gleichzeitig öffnen kann; „(Admin)“ ohne Rücksicht auf Groß-/Kleinschreibung. Eine Admin-Datei in „Schichtplanung“ bzw. eine Datei ohne „(Admin)“ in „Steuerung“ zeigt die Warnung „Dateiname falsch“.
- **Vorlagen-Prinzip:** Der Nutzer bekommt nur 5 Dateien und kopiert MA 01 / MA 01 (Admin) selbst. Jede Datei liest ihre **Nummer aus
  dem eigenen Dateinamen** (`CELL("filename")`; Eingabe `Listen!O10`, Stundenzettel `Daten!B2`; falscher Name → Warnung in Dateneingabe A4
  bzw. Stammdaten D3). Weil eine Verknüpfung fest auf EINE Datei zeigt, verbinden sich die Kopien nur mit den Sammel-Dateien:
  - MA NN (Export E–G, roh) → `002 Schichtplan` (Blatt `Daten`, daraus `Werte` = Anzeige) und → `001 MA-Tagesplan` (gleich aufgebaut)
  - `000 Inhaltsverzeichnis` (Namen) → Schicht- und Tagesplan (Blatt `Mitarbeiter`)
  - `002 Schichtplan` (`Daten`, `Mitarbeiter`) → MA NN (Admin) (eigene Spalten per `INDEX` über den ganzen Bereich)
  - MA NN (Admin) (`Vorgaben`) → `000 Inhaltsverzeichnis` (verstecktes Blatt `Vorgaben`, Zeile NN)
  - `000 Inhaltsverzeichnis` (`Vorgaben`: Name, Jahr, Urlaub, Soll, je Monat 4 Werte) → MA NN (eigene Zeile per `INDEX`)
  Bereichsgrößen sind fest auf 35 Plätze (`MAXP` in `vorlagen-bauen.py` und `plaene-erzeugen.py`).
- **Tagesplan fest nach Platz-Nummer** (Neue erscheinen unten, Kampagnen-Einträge verrutschen nicht); Schichtplan weiter nach Nachname.
- Löst Beobachtung 4/5: alle Kopien stammen von einer Vorlage, die einmal in Excel gespeichert wurde → jede Quelle hat gespeicherte Werte.
  6a: Hinweis im Inhaltsverzeichnis „Stundenzettel im Explorer per Doppelklick öffnen“. 6b offen.
- **Nachteil (dem Nutzer gesagt):** Der Stundenzettel bekommt die Einträge jetzt über den Schichtplan → Schichtplan nach dem Aktualisieren
  **speichern**. Mitarbeitende müssen ihre Platz-Nummer kennen. Ändert Claude später (im Produktivbetrieb) eine Vorlage, bringen neue Dateien
  die Einträge der Mitarbeitenden nicht mit → Änderungen dann nur mit Übertragen der Daten.
- **Reihenfolge beim ersten Mal (Steuerer, Excel am PC):**
  1. Dateien verteilen: `MA 01`, `001`, `002` nach „Schichtplanung“; `MA 01 (Admin)`, `000` nach „Steuerung“.
  2. `MA 01` öffnen → nur „Bearbeitung aktivieren“ (NICHT „Inhalte aktivieren“) → speichern → schließen. Ebenso `MA 01 (Admin)`.
  3. Im Explorer kopieren und umbenennen: `MA 02` … `MA 35`, `MA 02 (Admin)` … `MA 35 (Admin)` (genau so: Leerzeichen, führende Null).
  4. `000 Inhaltsverzeichnis` öffnen → Bearbeitung + Inhalte aktivieren → Workbook Links → aktualisieren → Namen prüfen/eintragen → speichern.
  5. `002 Schichtplan` öffnen → aktualisieren → **speichern**. 6. `001 MA-Tagesplan` öffnen → aktualisieren → speichern.
  7. Stundenzettel öffnen (Explorer) → aktualisieren → Name prüfen → Stammdaten pflegen → speichern.
  8. `000 Inhaltsverzeichnis` aktualisieren → speichern. 9. MA-Dateien öffnen → aktualisieren → Name im Titel, Reststunden → speichern.
- **Neue Mitarbeiterin / neuer Mitarbeiter:** Name in die nächste freie Zeile → speichern → Schichtplan + Tagesplan aktualisieren, speichern →
  MA NN (Admin) aktualisieren, Stammdaten pflegen, speichern → Inhaltsverzeichnis aktualisieren, speichern → MA NN aktualisieren, speichern.
- **Alltag:** Mitarbeitende speichern ihre Platz-Datei → Steuerer aktualisiert Schichtplan (+ speichern) und Tagesplan; für Stundenkonten zusätzlich
  Stundenzettel aktualisieren.

Test Paket V9 beim Arbeitgeber (V8 mit neuen Dateinamen) (35 Plätze, Inhaltsverzeichnis mit 35 Testnamen „Mustermann, …“ vorbelegt):
- [ ] Erstes Mal genau nach der Reihenfolge oben (Schritte 1–9); alle Kopien angelegt.
- [ ] Schichtplan: 35 Testnamen nach Nachname, Zeiten 11:00 (Vorlage); Tagesplan: Namen in der Reihenfolge des Inhaltsverzeichnisses.
- [ ] MA 05 (Admin): Name „Mustermann, …“ aus Zeile 5, Tagesliste gefüllt.
- [ ] MA 05: Titel mit Namen, oben Reststunden.
- [ ] Eine Kopie falsch benennen (z. B. „MA 5.xlsx“) → oben rote Warnung „Dateiname falsch“; danach wieder richtig benennen.
- [ ] Neuer Name: Zeile 5 im Inhaltsverzeichnis ändern (nur im Test!) → speichern → Schichtplan aktualisieren → neuer Name erscheint.
- [ ] Änderung durchreichen: in MA 05 Urlaub am 04.01. → speichern → Schichtplan aktualisieren, speichern → MA 05 (Admin) aktualisieren → Urlaub.
- [ ] Prüfen, ob irgendwo `#BEZUG!` erscheint oder „Workbook Links“ grau wird.
- [ ] Kennwort zum Öffnen (Wunsch von 2 Mitarbeitenden, ungetestet): bei „MA 05“ „Kennwort zum Öffnen“ setzen → Schichtplan und Tagesplan
  aktualisieren → fragt Excel nach dem Kennwort (erwartet: je Datei einmal, also bei 2 Personen 4 Abfragen je Durchgang) und stimmen danach
  die Werte? Abbrechen → alte Werte bleiben? Admin-Dateien lesen über den Schichtplan und sollten nicht fragen.
- [ ] Vergleich: bei „MA 06“ nur „Kennwort zum Ändern“ setzen → Aktualisieren ohne Kennwort-Abfrage? (Empfehlung, wenn es nur ums
  Nicht-Verändern geht; Schichtzeiten sind im Schicht-/Tagesplan ohnehin für alle sichtbar.)

### 10a. Geplanter Umbau: Eingabe-Datei (offen) + Stundenzettel (privat) – Entscheidungen vom 27.11.2026

Ziel: Kollegen sehen nur Schichten, keine Vertrags-/Kontodaten. **Gebaut am 27.11.2026**; nach Test V4 überarbeitet (Stand V5, siehe
„Stand V5“ unten – geht vor, wo es den Punkten darunter widerspricht). Abschnitt 10 gilt weiter für Schicht-/Tagesplan, außer wo hier anders.

**Stand V5 (Wünsche nach Test V4):**
- **Keine Anleitungs-Blätter** in Eingabe-Datei, Stundenzettel und Schichtplan (Wunsch des Nutzers: erklärt er bei der Einweisung).
  Die kurzen Hinweiszeilen unter den Überschriften der Blätter bleiben.
- Eingabe-Datei: kein Blatt „Start“ mehr; Blätter Anleitung, Musterwoche, Dateneingabe, Urlaub (+ versteckt Listen, Export, Vorgaben).
  Titel „Musterwoche … – Name“, „Dateneingabe 2027 – Name“, „Urlaubsübersicht – Name“ (keine Nummern). Ebenso im Stundenzettel:
  „Stammdaten – Name“, „Tagesliste 2027 – Name“, „Urlaubsübersicht & Abgleich – Name“, „Auswertung & Jahresbilanz 2027 – Name“. Name weiter aus dem Dateinamen
  (Hilfszellen Listen!O2:O4).
- **Vorgaben** kommen per Verknüpfung aus dem Stundenzettel (dort Blatt `Vorgaben`: A1 Jahr, B1 Urlaubsanspruch anteilig, C1 Soll Sek.,
  Zeile 3 ohne Vertrag (1), Zeile 4 Regel-Tag Sek., Zeile 5 Wert 057 Sek. (0 wenn leer), Zeile 6 = 1 wenn 057-Wert vorhanden – leere
  Zellen kommen über Verknüpfungen als 0 an, darum das Kennzeichen). Ohne Vorgaben: Jahr 2027, Vertrag in allen Monaten,
  keine Urlaubswarnung, Hinweis „noch keine Vorgabe vom Steuerer“. Mitarbeiter können die Verknüpfung nicht aktualisieren (kein Zugriff) –
  es bleiben die zuletzt vom Steuerer gespeicherten Vorgaben. Nach Vertragsänderung: Stundenzettel speichern → Eingabe-Datei öffnen →
  aktualisieren → speichern.
- **Reststunden** (Dateneingabe H3:J3): Soll − Ist, Ist je Monat = Plan aus der Dateneingabe (wie Auswertung, inkl. Feiertags-Regel-Tag,
  ohne Urlaub, nur Vertragsmonate) bzw. 057-Wert. Text „noch zu leisten: hh:mm Std.“ (rot) / „zu viel geleistet: …“ (grün) / „Soll genau
  erreicht“, fett. Rechnet identisch zum Stundenzettel (Test vergleicht die Sekunden). Nutzer nimmt in Kauf, dass man daraus auf den
  Vertrag schließen kann.
- Stundenzettel heißen **„Stundenzettel 01.xlsx“ … „35.xlsx“** (fest je Platz). Beim Mitarbeiterwechsel umstellen: Stundenzettel (Quelle =
  neue Eingabe-Datei), Schichtplan und Tagesplan (je Quelle ändern). Die Eingabe-Datei zeigt weiter auf denselben Stundenzettel.
- **Schicht- und Tagesplan lesen beide direkt die Eingabe-Dateien** (je Platz eine Verknüpfung, Blatt `Daten`; keine Blätter `Vertrag`/`Werte`
  mehr). Monate ohne Vertrag erkennt die Eingabe-Datei selbst („kein Vertrag“) → im Plan leer. Kein Speichern von Stundenzettel oder
  Schichtplan nötig, damit die Pläne aktuell sind – nur die Mitarbeiter müssen ihre Eingabe-Datei speichern.

Umsetzung:
- `vorlagen-bauen.py` leitet aus der Hauptvorlage ab:
  - **Eingabe.xlsx**: Blätter Anleitung, Start (Name aus Dateiname, Jahr, Urlaubsanspruch in Tagen `UrlaubAnspruchTage` = Start!B5),
    Musterwoche, Dateneingabe (ohne „kein Vertrag“/Regel-Tag; Netto nur an Arbeitstagen; A4 = Urlaubswarnung, Urlaub-Zellen rot
    bei Überschreitung), Urlaub (nur Tage), Listen, Export (wie bisher + E/F/G = Gilt Start/Pause/Anmerkung roh, H1 = Anspruch).
  - **Stundenzettel-Vorlage.xlsx**: Hauptvorlage ohne Export; Dateneingabe heißt **Tagesliste** (D/E/F grau, aus Blatt `Daten` =
    Verknüpfung `[1]Export!` der Eingabe-Datei, keine Musterwoche-Ersatzwerte); Musterwoche versteckt; Stammdaten B3 = Name aus der
    Eingabe-Datei; Urlaub A28:C28 = Abgleich Anspruch Eingabe ↔ Stammdaten; Blatt `Vertrag` A1:L1 = 1/0 je Monat (Jahres-Basis > 0).
    Druck: Stammdaten, Urlaub, Auswertung, Tagesliste auf Seitenbreite.
- `plaene-erzeugen.py` legt `Schichtplanung/` (Eingabe-Dateien, Schichtplan, Tagesplan) und `Steuerung/` („Stundenzettel Name.xlsx“
  mit Verknüpfung `../Schichtplanung/Name.xlsx`) an. Schichtplan: je Platz Verknüpfung auf Eingabe-Datei (Export) UND Stundenzettel
  (`../Steuerung/…`, Blatt Vertrag); verstecktes Blatt `Werte` = `Daten`, aber leer in Monaten ohne Vertrag; Wochenblätter und
  Tagesplan lesen `Werte`. Mitarbeiter sehen im Schichtplan den zuletzt vom Steuerer gespeicherten Stand (Aktualisieren scheitert bei
  ihnen am fehlenden Zugriff auf „Steuerung“ – Meldung einfach schließen).
- Reihenfolge Steuerer: Eingabe-Dateien gespeichert → Stundenzettel aktualisieren + speichern → Schichtplan aktualisieren + speichern
  → Tagesplan aktualisieren.
- Ein neuer Stundenzettel, der noch nie gespeichert wurde, liefert „kein Vertrag“ (0) → der Mitarbeiter erscheint im Schichtplan ohne Zeiten,
  bis der Steuerer den Stundenzettel einmal geöffnet und gespeichert hat.

Entscheidungen:

- **Abteilungsordner (alle):** je Mitarbeiter eine **Eingabe-Datei** (Musterwoche + Dateneingabe: Beginn, Pause, Anmerkung) sowie
  Schichtplan und Tagesplan. **Ordner des Steuerers (nur er):** je Mitarbeiter ein **Stundenzettel** (Stammdaten, Urlaub, Auswertung,
  Tagesliste), der die Einträge per Verknüpfung aus der Eingabe-Datei holt. Verknüpfungen nur Richtung privat → offen.
- Beide Ordner liegen auf **demselben Laufwerk** (relative Pfade möglich; Ordner nicht umbenennen/verschieben).
- **Abwesenheit:** trägt nur der **Steuerer** in die Eingabe-Datei ein, neutral als **„abwesend (Gutschrift)“** (ersetzt „krank“;
  geplante Zeit wird gutgeschrieben). Mitarbeiter tragen Krankheit nicht selbst ein. **Keine** extra Spalte „Grund“.
- **Eingabe-Datei zeigt nur Resturlaub in Tagen** (Anspruch / genommen / Rest) und **warnt bei zu viel Urlaub** (Anspruch in Tagen
  steht dafür in der Eingabe-Datei). Stunden, Soll/Ist, Regel-Tag, Jahres-Basis nur im Stundenzettel. Netto an Urlaubs-/Feiertagen
  in der Eingabe-Datei nicht anzeigen (verriete den Regel-Tag).
- **Ausscheiden automatisch:** Schichtplan soll das Vertragsende selbst erkennen (z. B. Vertragsmonate aus dem Stundenzettel – nur
  der Steuerer aktualisiert den Schichtplan; dort würden dann nur Ja/Nein-Merker je Monat gespeichert).
- **Schutz:** „Kennwort zum Ändern“ je Eingabe-Datei (Mitarbeiter + Steuerer kennen es).
- **Ausdruck/PDF** für den Mitarbeiter: Stammdaten, Urlaub, Auswertung und Tagesliste.
- Bei Mitarbeiterwechsel: Verknüpfung im Schichtplan UND im Stundenzettel umstellen (Anleitung schreiben).
