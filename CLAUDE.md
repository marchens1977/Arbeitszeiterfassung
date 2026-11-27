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
| `excel/Arbeitszeiterfassung.xlsx` | Excel-Version der App (eigenständig, siehe Abschnitt 9) |
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

- **Rechenmodell nicht verändern**, es sei denn, der Nutzer wünscht es ausdrücklich: Regel-Tag = Jahresstunden / 52 / 5;
  Urlaub wird vom Soll abgezogen; Feiertage zählen als Regel-Tag auf das Ist; Urlaubstage zählen nicht auf das Ist.
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

## 4. Offene Ideen / nächste Schritte

- Vorbelegter Urlaub: Der 26.03.2027 (Karfreitag) war als Urlaub eingetragen und zählt nicht; dadurch sind nur
  29 statt 30 Tage vorbelegt. Nutzer kann einen Tag nachtragen – oder die Vorbelegung wird angepasst.
- Vorbereitung auf ein neues Jahr (2028): Jahr umstellen, vorbelegten Urlaub pflegen, ggf. Übernahme des Übertrags.
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
| `node werkzeuge/offline-app.test.js` | Startet einen lokalen Webserver und prüft Manifest, Symbole, Offline-Helfer und den Start ohne Internet |
| `python3 werkzeuge/excel-pruefen.py [datei.xlsx] [--jahresbasis WERT]` | Rechnet die Excel-Datei mit LibreOffice neu und gibt Soll, Regel-Werte, Urlaub und Auswertung aus. Zum Vergleich alte und neue Datei prüfen und die Ausgaben mit `diff` vergleichen. |
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
- Updates: Unter Stammdaten → „App-Updates“ steht eine Version (z. B. v4); „Jetzt nach Updates suchen“ meldet
  „neueste Version“. Nach dem nächsten Update öffnet sich die App kurz neu und zeigt „gerade aktualisiert“.

## 9. Excel-Version

**Datei:** `excel/Arbeitszeiterfassung.xlsx` (am 23.11.2026 von Claude erstellt – in welcher Sitzung, ist nicht mehr bekannt;
der Nutzer hat sie am 27.11.2026 vom Handy hochgeladen). Sie ist eine **leere Vorlage ohne persönliche Daten** und darf deshalb
im öffentlichen Repository liegen. Eigene ausgefüllte Kopien gehören **nicht** ins Repository.

Zweck: dieselbe Rechnung wie die App, aber als Excel-Tabelle (ohne Internet, ohne Browser). Sie ist unabhängig von der App –
Daten werden **nicht** zwischen App und Excel ausgetauscht.

Aufbau (Blätter):

| Blatt | Inhalt |
|---|---|
| Anleitung | Bedienung, Farben (Gelb = eintragen, Grau = rechnet Excel), Rechenregeln, Sichern, neues Jahr |
| Stammdaten | Jahr (2027), Urlaubsanspruch, Übertrag (Plus/Minus), vereinbarte Überstunden, unbezahlter Urlaub, Jahres-Basis je Monat |
| Musterwoche | Mo–Fr Start/„Frei“ und Pause, Ende fix 19:00 |
| Dateneingabe | alle Werktage des Jahres, hessische Feiertage vorbelegt, Klappmenü für Start/Urlaub/Feiertag/Frei |
| Urlaub | Anspruch, genommen, Rest, je Monat; Warnung bei zu viel Urlaub |
| Auswertung | Monatswerte Plan, Eingabe „Stundenabrechnung 057“, Soll/Ist/Differenz |
| Listen | Hilfsblatt: Klappmenü-Werte, Osterberechnung (Gauß), Feiertage |

Wichtig:
- Gleiches Rechenmodell wie die App (Regel-Tag = Jahresstunden / 52 / 5; Urlaub vom Soll abgezogen; Feiertage zählen auf das Ist).
- Zeiten als Stunden:Minuten eingeben (`12:30`), nicht als Kommazahl – **auch die Jahres-Basis** (`2028:00`, Anzeige `2028:00:00`).
  Intern ist sie ein Excel-Zeitwert (Tage); die Prüfung lässt höchstens 5000 Std. zu, damit `2028` ohne Doppelpunkt abgewiesen wird.
- Zahlenformat `0.##` nicht verwenden: es zeigt bei ganzen Zahlen ein Komma („30,“). Stattdessen `General`
  (bzw. `General" Tage"`). In Formeln Tage mit `&""` statt `TEXT(…;"0.##")` einbauen (sonst „031“ in deutschem Excel).
- Blätter sind **ohne Passwort** geschützt (nur gegen versehentliches Löschen von Formeln), keine Makros.
- Für ein neues Jahr: Kopie speichern, Jahr in „Stammdaten“ ändern, alte Eingaben löschen.
- Bei Änderungen an der Datei: mit `openpyxl` (Python) bearbeiten, danach die Formeln mit LibreOffice neu berechnen lassen
  und die Ergebnisse mit der App vergleichen (gleiche Eingaben → gleiche Zahlen).
- Download-Link (nach dem Merge): `https://marchens1977.github.io/Arbeitszeiterfassung/excel/Arbeitszeiterfassung.xlsx`

Tests für den Nutzer (Excel):
- [ ] Datei über den Download-Link am Handy öffnen (z. B. Excel-App oder Google Tabellen).
- [ ] In der Musterwoche einen Wert ändern → Dateneingabe und Auswertung ändern sich mit.
- [ ] Stammdaten: Urlaubsanspruch ohne Komma, Jahres-Basis als `2028:00:00`; Eingabe `2028:00` wird angenommen, `2028` abgewiesen.
- [ ] Gleiche Eingaben wie in der App → gleiche Jahresbilanz (stichprobenartig vergleichen).
