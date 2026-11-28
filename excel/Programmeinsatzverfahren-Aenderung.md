# Änderungsmeldung zum bestehenden Programmeinsatzverfahren „Arbeitszeiterfassung“ (Excel)

Stand: 28.11.2026 · Version der Lösung: V9
Felder in [eckigen Klammern] sind vom Fachbereich auszufüllen.

## 1. Bezug

- Bestehendes, genehmigtes Programmeinsatzverfahren: [Bezeichnung / Nummer / Datum der Genehmigung]
- Fachlich verantwortlich: [Name, Funktion]
- Technisch verantwortlich / Pflege: [Name]
- Nutzerkreis: Mitarbeitende der Abteilung [Bezeichnung] (bis zu 35 Personen) und Steuerer/Teamleitung

## 2. Art der Änderung

Technische Umgestaltung des Ablaufs. Inhaltlich unverändert bleiben:

- die verarbeiteten Daten: Arbeitszeiten (Arbeitsbeginn, Pause), Urlaub, Abwesenheiten, Feiertage, vertragliche Jahresstunden, Urlaubsanspruch, Stundenkonto
- die Rechenregeln: Regelarbeitstag = Jahresstunden des Monats / 52 / 5; Urlaub wird vom Soll abgezogen; Feiertage zählen mit dem Regelarbeitstag auf das Ist; Urlaubstage zählen nicht auf das Ist; festes Arbeitsende 19:00 Uhr; Ist je Tag = 19:00 Uhr − Arbeitsbeginn − Pause
- der Zweck: Einsatzplanung und Abgleich mit der offiziellen Stundenabrechnung (Abrechnung 057). Die Lösung ist kein führendes Zeiterfassungssystem und ersetzt die offizielle Abrechnung nicht.

Bereits enthalten und nur der Vollständigkeit halber genannt: unterjähriger Eintritt/Austritt (Monat mit Jahresstunden 0 = „kein Vertrag“; Urlaub anteilig = Anspruch × Vertragsmonate ÷ 12, Bruchteile ab einem halben Tag aufgerundet, angelehnt an § 5 BUrlG). [Streichen, falls schon Teil des genehmigten Verfahrens.]

## 3. Was sich geändert hat

1. **Aufteilung in mehrere verbundene Dateien statt einer Datei je Mitarbeitendem**
   - Ordner „Schichtplanung“ (Lesezugriff Abteilung): „MA 01.xlsx“ bis „MA 35.xlsx“ (Eingabe-Datei je Mitarbeitendem: Musterwoche, tägliche Eingabe, Urlaub in Tagen – ohne Vertrags- und Kontodaten), „002 Schichtplan.xlsx“ (Wochenansicht), „001 MA-Tagesplan.xlsx“ (Tagesbesetzung mit Feldern für die Kampagnenzuordnung)
   - Ordner „Steuerung“ (Zugriff nur Steuerer): „MA 01 (Admin).xlsx“ bis „MA 35 (Admin).xlsx“ (Stundenzettel je Mitarbeitendem: Stammdaten, Tagesliste, Urlaubsabgleich, Auswertung), „000 Inhaltsverzeichnis.xlsx“ (einzige Stelle für die Namenspflege, Sammelstelle der Vorgaben)
2. **Neu: Schichtplan und Tagesplan** füllen sich automatisch aus den Eingabe-Dateien.
3. **Trennung offen/privat:** Vertrags- und Kontodaten liegen nur noch im Ordner des Steuerers.
4. **Abwesenheiten** trägt ausschließlich der Steuerer ein, neutral als „abwesend (Gutschrift)“; die geplante Zeit wird gutgeschrieben. Ein Grund wird nicht gespeichert, in den Plänen steht nur „abwesend“.
5. **Technik:** Die Dateien sind über Excel-Verknüpfungen zwischen Arbeitsmappen verbunden (relative Pfade, gleiches Laufwerk, manuelle Aktualisierung über „Daten → Arbeitsmappenlinks“). Die Dateien „MA NN“ und „MA NN (Admin)“ sind unveränderte Kopien je einer Vorlage und erkennen ihre Nummer am Dateinamen; ein falscher Dateiname erzeugt eine sichtbare Warnung. Weiterhin keine Makros, kein VBA, keine Verbindung zu anderen Systemen oder ins Internet.
6. **Erstellung:** Die neue Fassung wurde vom Fachbereich mit Unterstützung eines KI-Assistenten (Claude, Anthropic) erstellt. Erzeugungs- und Prüfprogramme sind versioniert dokumentiert. Personenbezogene Daten wurden dabei nicht weitergegeben; echte Namen trägt der Steuerer erst im Betrieb ein.

## 4. Auswirkungen auf Zugriff und Sichtbarkeit (Kernpunkt der Meldung)

- Für die Mitarbeitenden der Abteilung sichtbar (Schicht- und Tagesplan, Eingabe-Dateien): Name, geplante Schichtzeiten, Urlaub/Feiertag/Frei, „abwesend“ (ohne Grund), Anmerkung (Büro/Homeoffice/Teammentor*in).
  [Angeben, ob diese gegenseitige Sichtbarkeit bereits im genehmigten Verfahren bestand. Falls nein, ist dies die wesentliche inhaltliche Änderung.]
- Nur für den Steuerer sichtbar: vertragliche Jahresstunden, Urlaubsanspruch, Übertrag, Stundenkonto, Soll-/Ist-Abgleich.
- Hinweis: Aus den in der Eingabe-Datei angezeigten Reststunden lassen sich grobe Rückschlüsse auf den Vertragsumfang ziehen; vom Fachbereich bewusst in Kauf genommen.
- Gesundheitsdaten werden nicht gespeichert.

## 5. Berechtigungen und Schutz

- Der Zugriffsschutz erfolgt über die Ordnerrechte des Netzlaufwerks: „Schichtplanung“ für die Abteilung, „Steuerung“ nur für den Steuerer.
- Blattschutz ohne Kennwort: verhindert nur versehentliches Überschreiben von Formeln, keine Manipulationssicherung.
- Optional „Kennwort zum Ändern“ je Eingabe-Datei (andere können nur lesen; die Aktualisierung funktioniert weiter).
- Auf Wunsch einzelner Mitarbeitender „Kennwort zum Öffnen“ der eigenen Eingabe-Datei; das Kennwort wird dem Steuerer mitgeteilt. Erwartung: Excel fragt beim Aktualisieren von Schicht- und Tagesplan nach dem Kennwort. [Testergebnis eintragen.] Die Schichtzeiten bleiben im Schicht- und Tagesplan für die Abteilung sichtbar.

## 6. Risiken und Kontrollen (neu durch den geänderten Ablauf)

- Veraltete Werte, weil nicht gespeichert oder nicht aktualisiert → feste Reihenfolge der Aktualisierung (Abschnitt 8), Hinweise in den Dateien.
- Falsch benannte oder kopierte Datei → automatische Erkennung, rote Warnung „Dateiname falsch“; Admin-Datei im offenen Ordner wird ebenfalls erkannt.
- Verknüpfung nicht erreichbar oder fehlerhaft → alle Auswertungen sind abgesichert und zeigen „keine Vorgabe“ bzw. leer statt Fehlerwert.
- Excel ersetzt eine Verknüpfung dauerhaft durch `#BEZUG!`, wenn eine noch nie in Excel gespeicherte Quelle aktualisiert wird → Vorlagen zuerst in Excel speichern, erst dann kopieren; Reihenfolge beim ersten Einrichten einhalten.
- Umbenennen oder Verschieben von Dateien/Ordnern → Dateinamen und Ordneranordnung sind fest; Umbenennen ist nicht zulässig.
- Mehr Urlaub eingetragen als Anspruch → Warnung in Eingabe-Datei und Stundenzettel.
- Abweichung zur offiziellen Abrechnung → Eintrag des Werts aus Abrechnung 057 mit sichtbarem Abgleich.

## 7. Prüfung der Änderung

- Automatisierter Kettentest (`werkzeuge/plaene.test.py`, 112 Prüfpunkte, mit LibreOffice): Erzeugung, Kopieren, Nummernerkennung, falsche Dateinamen, Fehlerfälle bei Verknüpfungen, Schicht- und Tagesplan, Stundenzettel. Der Stundenzettel liefert sekundengenau dieselben Ergebnisse wie die bisherige Rechenlogik (Hauptvorlage); Reststunden in Eingabe-Datei und Stundenzettel stimmen überein; Vertragsende und anteiliger Urlaub werden geprüft.
- Bisherige Tests beim Anwender in Excel: Mini-Pakete V4–V6, Volltest V7 mit 35 Testpersonen (Öffnen/Aktualisieren unter einer Minute).
- Fachlicher Test der Fassung V9 in Excel: [Datum, Tester, Ergebnis]
- Abnahme: [Name, Datum]

## 8. Betriebsablauf

- Erstes Einrichten: Vorlagen einmal in Excel speichern → im Explorer kopieren (MA 02 … MA 35, MA 02 (Admin) … MA 35 (Admin)) → Inhaltsverzeichnis aktualisieren, speichern → Schichtplan und Tagesplan aktualisieren, speichern → Stundenzettel aktualisieren, Stammdaten pflegen, speichern → Inhaltsverzeichnis erneut aktualisieren, speichern → Eingabe-Dateien aktualisieren, speichern.
- Laufender Betrieb: Mitarbeitende speichern ihre Eingabe-Datei; der Steuerer aktualisiert Schichtplan (und speichert) und Tagesplan, bei Bedarf die Stundenzettel.
- Neue Mitarbeitende: nächste freie Zeile im Inhaltsverzeichnis. Im laufenden Jahr wird nichts gelöscht oder überschrieben; Ausscheiden über Jahresstunden 0.
- Datensicherung über die reguläre Sicherung des Netzlaufwerks. [bestätigen]
- Aufbewahrung und Löschung: [wie im genehmigten Verfahren / Regel eintragen]
- Jahreswechsel: neues Paket für das Folgejahr; die alten Dateien bleiben unverändert erhalten.

## 9. Änderungsverfahren und Grenzen

- Änderungen erfolgen an der Hauptvorlage; abgeleitete Vorlagen werden per Programm erzeugt. Nach jeder Änderung: automatisierter Kettentest, fachlicher Test in Excel, neue Versionsnummer, Eintrag im Änderungsprotokoll.
- Neue Vorlagen übernehmen keine bereits erfassten Daten; Änderungen im laufenden Jahr erfordern eine Datenübernahme.
- Keine Mehrbenutzer-Datenbank; Aktualität hängt vom Speichern und Aktualisieren ab.
- Nur Excel am PC (Nummern-/Namenserkennung über den Dateinamen funktioniert nicht in Excel im Browser oder auf dem Handy).
- Höchstens 35 Plätze je Jahr; kein revisionssicheres Protokoll der Änderungen an den Einträgen.

## 10. Bitte um Prüfung

Wir bitten um Prüfung, ob die Änderung im Rahmen des bestehenden Verfahrens fortgeführt werden kann und ob wegen der geänderten Sichtbarkeit (Schicht- und Tagesplan, Abschnitt 4) eine erneute Beteiligung von Datenschutz und Personalrat erforderlich ist.
