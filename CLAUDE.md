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
- Bei jeder Änderung an der App die `VERSION` in `sw.js` hochzählen.
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

## 4. Offene Ideen / nächste Schritte

- Vorbelegter Urlaub: Der 26.03.2027 (Karfreitag) war als Urlaub eingetragen und zählt nicht; dadurch sind nur
  29 statt 30 Tage vorbelegt. Nutzer kann einen Tag nachtragen – oder die Vorbelegung wird angepasst.
- Vorbereitung auf ein neues Jahr (2028): Jahr umstellen, vorbelegten Urlaub pflegen, ggf. Übernahme des Übertrags.
- Optional: Hinweis in der App, wenn eine neue Version geladen wurde.

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
| `node werkzeuge/offline-app.test.js` | Startet einen lokalen Webserver und prüft Manifest, Symbole, Offline-Helfer und den Start ohne Internet |
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
