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
| `excel/Arbeitszeiterfassung.xlsx` | Excel-Version der App (eigenständig, siehe Abschnitt 9) – zugleich Vorlage für jede Mitarbeiter-Datei |
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
| 2026-11-27 | Excel: Spalte „Anmerkung“ (Dateneingabe F = Eingabe, I = gilt; Musterwoche D) mit Büro/Homeoffice/Teammentor*in/krank; „krank“ = geplante Zeit wird gutgeschrieben (Rechnung unverändert); verstecktes Blatt „Export“ für den Schichtplan. Neu: Schichtplan + MA-Tagesplan ohne Makros (Abschnitt 10), Mini-Paket zum Test bei der Naspa; Werkzeuge `xlsx_spalten.py`, `plaene-erzeugen.py`, `plaene.test.py` | #15 |

## 4. Offene Ideen / nächste Schritte

- Vorbelegter Urlaub: Der 26.03.2027 (Karfreitag) war als Urlaub eingetragen und zählt nicht; dadurch sind nur
  29 statt 30 Tage vorbelegt. Nutzer kann einen Tag nachtragen – oder die Vorbelegung wird angepasst.
- Vorbereitung auf ein neues Jahr (2028): Jahr umstellen, vorbelegten Urlaub pflegen, ggf. Übernahme des Übertrags.
- **Noch nicht in der App nachgezogen** (Excel ist Vorlage, siehe Grundregel in Abschnitt 2):
  - Spalte „Anmerkung“ (Büro/Homeoffice/Teammentor*in/krank) in Dateneingabe und Musterwoche (seit #15 in Excel). Nutzer fragen.
- Schichtplan/Tagesplan: nach erfolgreichem Mini-Test bei der Naspa das volle Paket erzeugen (35 Plätze, 52 Wochen, alle Tage) und schicken.
  Bewusste Unterschiede, weil technisch nicht anders möglich: In der App wird der Mitarbeitername eingetippt (kein Dateiname);
  Excel hat zusätzlich das Blatt „Anleitung“, die App erklärt in kurzen Hinweistexten.
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
| `python3 werkzeuge/plaene.test.py` | Schicht- und Tagesplan: erzeugt ein Mini-Paket, speichert die Mitarbeiter-Dateien mit LibreOffice, setzt die gespeicherten Werte wie Excels „Werte aktualisieren“ ein und prüft Namen, Sortierung, Zeiten, Urlaub, „abwesend“ (krank), Teammentor*in, frei, HO, Platzhalter. Endet mit „ERGEBNIS: alles in Ordnung“ (dauert ca. 1–2 Min). |
| `python3 werkzeuge/plaene-erzeugen.py --ordner ORDNER [...]` | Kein Test: erzeugt Mitarbeiter-Dateien, `Schichtplan.xlsx`, `MA-Tagesplan.xlsx` (Optionen siehe Kopf der Datei; Mini-Paket: `--plaetze 3 --wochen 1 --tage 5 --mitarbeiter "Beispiel, Anna" "Test, Ben" --beispieldaten`) |
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
- [ ] Anmerkung: Musterwoche Freitag „Homeoffice“ → in der Dateneingabe steht freitags „Homeoffice“ (gelb); einzelner Tag „Büro“ hebt es auf; „krank“ ändert die Stunden nicht.
- [ ] Gleiche Eingaben wie in der App → gleiche Jahresbilanz (stichprobenartig vergleichen).

## 10. Schichtplan und MA-Tagesplan (Excel, ohne Makros)

Wunsch des Nutzers (27.11.2026): Steuerer sieht automatisch, wer wann da ist. Vorlagen kamen per Mail
(„Schichtplan.xlsx“ mit Wochenblatt „04.01.2027“, „MA-Tagesplan.xlsx“ mit Kampagnen-Spalten).
**Makros sind bei der Naspa nicht erlaubt** → nur Excel-Verknüpfungen (Daten → Links bearbeiten → Werte aktualisieren).

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

Test bei der Naspa (Mini-Paket: „Beispiel, Anna.xlsx“, „Test, Ben.xlsx“, „Mustermann, Max 03.xlsx“, Schichtplan 1 Woche, Tagesplan 5 Tage):
- [ ] Alle 5 Dateien in einen neuen Ordner auf dem Firmenlaufwerk legen.
- [ ] Anna und Ben je einmal öffnen und speichern (Stammdaten zeigt den Namen).
- [ ] Schichtplan öffnen → „Inhalt aktivieren“ bzw. Daten → Links bearbeiten → Werte aktualisieren → Werte erscheinen
  (Anna Mo 11:00 Homeoffice gelb, Di Urlaub grün; Ben Mi „abwesend“ grau, Do Teammentor*in, Fr Frei) → speichern.
- [ ] MA-Tagesplan öffnen → aktualisieren → 05.01.: Anna grau „Urlaub“, Ben 13:00; Kampagnen-Feld eintragbar.
- [ ] Ben in der Datei auf „Urlaub“ am 04.01. setzen, speichern → Schichtplan aktualisieren, speichern → Tagesplan aktualisieren → Ben grau.
- [ ] „Kennwort zum Ändern“ an einer Datei setzen → Schichtplan aktualisiert trotzdem.
