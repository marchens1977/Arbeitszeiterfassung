/* Prüft die automatischen App-Updates (Kasten "App-Updates" im Reiter Stammdaten).
   Kopiert die App in einen Testordner, startet dort einen Webserver (Port 8766) und "veröffentlicht"
   eine neue Version, indem die VERSION in der kopierten sw.js hochgezählt wird.
   Geprüft wird: Versionsanzeige, Knopf "Jetzt nach Updates suchen", automatisches Neuladen,
   Warten während einer Bearbeitung, Daten bleiben erhalten.  Start: node werkzeuge/updates.test.js */
const { chromium } = require('playwright'); const path = require('path'); const fs = require('fs'); const os = require('os');
const { spawn } = require('child_process');
const SRC = path.join(__dirname, '..');
const DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'az-update-'));
for (const f of ['index.html', 'manifest.json', 'sw.js']) fs.copyFileSync(path.join(SRC, f), path.join(DIR, f));
fs.cpSync(path.join(SRC, 'icons'), path.join(DIR, 'icons'), { recursive: true });
const swFile = path.join(DIR, 'sw.js');
const setVersion = v => fs.writeFileSync(swFile, fs.readFileSync(swFile, 'utf8').replace(/const VERSION = '[^']+'/, `const VERSION = 'arbeitszeit-${v}'`));
const server = spawn('python3', ['-m', 'http.server', '8766'], { cwd: DIR, stdio: 'ignore' });
process.on('exit', () => server.kill());
const URL = 'http://localhost:8766/index.html';

(async () => {
  await new Promise(r => setTimeout(r, 1000));
  setVersion('test1');
  const b = await chromium.launch(); const ctx = await b.newContext(); const p = await ctx.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message));
  let loads = 0; p.on('load', () => loads++);
  await p.goto(URL);
  await p.waitForFunction(() => navigator.serviceWorker.controller, null, { timeout: 10000 });
  await p.waitForTimeout(500);
  console.log('Version erste Installation (erwartet test1):', await p.textContent('#appVersion'));
  console.log('Neu geladen bei Erstinstallation (erwartet 1 Ladevorgang):', loads);

  await p.click('#updateBtn'); await p.waitForTimeout(2000);
  console.log('Ohne neue Version:', await p.textContent('#updateStatus'));
  console.log('Zuletzt gesucht gesetzt:', (await p.textContent('#lastUpdateCheck')) !== 'noch nie');

  // Neue Version veröffentlichen -> Knopf drücken -> App lädt neu
  setVersion('test2');
  const before = loads;
  await p.click('#updateBtn');
  await p.waitForTimeout(3000);
  console.log('Neu geladen nach Update:', loads > before);
  console.log('Version (erwartet test2):', await p.textContent('#appVersion'));
  console.log('Hinweis:', await p.textContent('#updateStatus'));

  // Während einer Bearbeitung darf nicht neu geladen werden
  await p.click('#mainBtn0'); // "Ändern"
  await p.fill('#vac', '27');
  setVersion('test3');
  const before2 = loads;
  await p.click('#updateBtn'); await p.waitForTimeout(3000);
  console.log('Während Bearbeitung neu geladen (erwartet false):', loads > before2);
  console.log('Hinweis:', await p.textContent('#updateStatus'));
  await p.click('#mainBtn0'); // "Speichern" -> danach Update
  await p.waitForTimeout(3000);
  console.log('Nach Speichern neu geladen:', loads > before2);
  console.log('Version (erwartet test3):', await p.textContent('#appVersion'));
  console.log('Urlaubsanspruch erhalten (erwartet 27):', await p.inputValue('#vac'));

  // Automatische Suche beim Zurückkehren in die App (ohne Knopf)
  setVersion('test4');
  const before3 = loads;
  await p.evaluate(() => { lastAutoCheck = 0; autoCheck(); });
  await p.waitForTimeout(3000);
  console.log('Automatisch aktualisiert:', loads > before3, '| Version (erwartet test4):', await p.textContent('#appVersion'));

  console.log('errors:', errors);
  await b.close(); server.kill();
  fs.rmSync(DIR, { recursive: true, force: true });
})().catch(e => { console.error(e); process.exit(1); });
