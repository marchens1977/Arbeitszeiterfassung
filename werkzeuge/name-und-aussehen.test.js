/* Prüft die Angleichung an die Excel-Datei: Mitarbeitername (Eingabe, große Anzeige, Hinweis, Speichern,
   Sicherung), Farben (Gelb = eintragen, Grau = gesperrt, Urlaub hellblau), Uhrzeit-Listen 11:00–17:00,
   Reiter-/Überschriften wie in Excel. Macht Handy-Screenshots in einen Temp-Ordner.
   Start: node werkzeuge/name-und-aussehen.test.js */
const { chromium } = require('playwright'); const path = require('path'); const fs = require('fs'); const os = require('os');
const APP = 'file://' + path.join(__dirname, '..', 'index.html');
const OUT = fs.mkdtempSync(path.join(os.tmpdir(), 'az-name-'));

(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 390, height: 760 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message));
  await p.goto(APP);
  let fehler = 0;
  const pruefe = (name, ist, soll) => { const ok = String(ist) === String(soll); if (!ok) fehler++;
    console.log(`${ok ? 'OK  ' : 'FEHLER'} ${name}: ${ist}${ok ? '' : ' (erwartet ' + soll + ')'}`); };
  const txt = sel => p.$eval(sel, e => e.innerText.replace(/\s+/g, ' ').trim());
  const bg = sel => p.$eval(sel, e => getComputedStyle(e).backgroundColor);

  // Ohne Namen
  pruefe('Überschrift ohne Namen', await txt('#stammTitle'), '1. Stammdaten');
  pruefe('Feld gesperrt = grau', await bg('#vac'), 'rgb(242, 242, 242)');
  await p.screenshot({ path: OUT + '/1-ohne-name.png' });

  // Name falsch (ohne Komma) → Hinweis
  await p.click('#mainBtn0');
  pruefe('Feld beim Ändern = gelb', await bg('#vac'), 'rgb(255, 242, 204)');
  await p.fill('#empName', 'Max Mustermann'); await p.click('#mainBtn0');
  pruefe('Hinweis ohne Komma', await txt('#empHint'), 'Bitte als „Nachname, Vorname“ eintragen, z. B. „Mustermann, Max“.');

  // Name richtig (auch mit schiefem Leerzeichen) → wird zu „Mustermann, Max“
  await p.click('#mainBtn0'); await p.fill('#empName', '  Mustermann ,Max '); await p.click('#mainBtn0');
  pruefe('Überschrift mit Namen', await txt('#stammTitle'), '1. Stammdaten – Mustermann, Max');
  pruefe('Blaues Feld', await txt('#empBanner'), 'Mitarbeiter Mustermann, Max');
  pruefe('Blaues Feld Farbe', await bg('#empBanner'), 'rgb(0, 45, 114)');
  pruefe('Kein Hinweis', await txt('#empHint'), '');
  await p.screenshot({ path: OUT + '/2-mit-name.png' });

  // Bleibt nach Neuladen und steckt in der Sicherung
  await p.reload();
  pruefe('Nach Neuladen', await txt('#stammTitle'), '1. Stammdaten – Mustermann, Max');
  pruefe('In der Sicherung', await p.evaluate(() => JSON.parse(backupText()).data.name), 'Mustermann, Max');

  // Reiter und Überschriften wie in Excel
  pruefe('Reiter 5', await txt('.tab:nth-child(5)'), '5. Auswertung');
  pruefe('Überschrift Urlaub', await txt('#pg1 h2'), '2. Urlaubsübersicht & Abgleich');
  pruefe('Überschrift Auswertung', await txt('#pg4 h2'), '5. Auswertung & Jahresbilanz 2027');

  // Uhrzeiten und Urlaub-Farbe in der Dateneingabe
  await p.evaluate(() => { st.days = {}; save(); showTab(3); }); await p.click('#mainBtn3');
  const zeiten = await p.$$eval('#row_2027-01-05 td:nth-child(2) select option', os => os.map(o => o.value));
  pruefe('Auswahl Dateneingabe', zeiten[0] + ',' + zeiten[1] + ',' + zeiten[2] + ',' + zeiten[3] + ' … ' + zeiten[zeiten.length - 1], 'U,F,X,11:00 … 17:00');
  await p.selectOption('#row_2027-01-05 select >> nth=0', 'U');
  pruefe('Urlaubszeile hellblau', await bg('#row_2027-01-05 td:nth-child(2)'), 'rgb(221, 235, 247)');
  await p.evaluate(() => document.querySelector('#row_2027-01-05').scrollIntoView({ block: 'center' }));
  await p.screenshot({ path: OUT + '/3-dateneingabe.png' });
  await p.click('#mainBtn3');
  await p.evaluate(() => showTab(2)); await p.click('#mainBtn2');
  const muster = await p.$$eval('#patternBody tr:first-child td:nth-child(2) select option', os => os.map(o => o.value));
  pruefe('Auswahl Musterwoche', muster.filter(v => v !== 'X' && (v < '11:00' || v > '17:00')).length === 0 && muster.includes('17:00'), true);
  await p.click('#mainBtn2');

  console.log('errors:', errors);
  console.log('Screenshots:', OUT);
  console.log(fehler === 0 && !errors.length ? 'ERGEBNIS: alles in Ordnung' : `ERGEBNIS: ${fehler} Fehler`);
  await b.close();
  process.exit(fehler || errors.length ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
