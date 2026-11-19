/* Prüft die Datensicherung als Text: kopieren, Daten ändern, wiederherstellen, ungültigen Text abweisen.
   Start: node werkzeuge/sicherung.test.js   (macht zusätzlich einen Handy-Screenshot in einen Temp-Ordner) */
const { chromium } = require('playwright'); const path = require('path'); const fs = require('fs'); const os = require('os');
const APP = 'file://' + path.join(__dirname, '..', 'index.html');
const OUT = fs.mkdtempSync(path.join(os.tmpdir(), 'az-shot-'));
(async () => {
  const b = await chromium.launch(); const ctx = await b.newContext({ permissions: ['clipboard-read','clipboard-write'] }); const p = await ctx.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => m.type()==='error' && errors.push(m.text()));
  await p.goto(APP);
  const t = sel => p.$eval(sel, e => e.innerText.trim());
  console.log('last backup initial:', await t('#lastBackupInfo'));
  await p.evaluate(() => { st.vac = '27'; save(); });
  await p.click('button:has-text("Sicherung als Text kopieren")');
  await p.waitForFunction(() => document.querySelector('#backupInfo').innerText.trim() !== '', null, { timeout: 5000 }).catch(() => {});
  console.log('info:', await t('#backupInfo'));
  console.log('last backup now:', await t('#lastBackupInfo'));
  const clip = await p.evaluate(() => navigator.clipboard.readText());
  console.log('clipboard has vac 27:', clip.includes('"vac": "27"'));
  await p.evaluate(() => { st.vac = '5'; save(); });
  await p.click('button:has-text("Text anzeigen")');
  await p.fill('#backupText', clip);
  p.on('dialog', d => d.accept());
  await p.click('button:has-text("Aus Text wiederherstellen")'); await p.waitForLoadState('load'); await p.waitForTimeout(300);
  console.log('after text restore vac (erwartet 27):', await p.$eval('#vac', e => e.value));
  await p.click('button:has-text("Text anzeigen")'); await p.fill('#backupText', 'kein json');
  
  await p.click('button:has-text("Aus Text wiederherstellen")'); await p.waitForTimeout(200);
  console.log('file input accept:', await p.$eval('#importFile', e => e.getAttribute('accept')));
  console.log('errors:', errors);
  await p.setViewportSize({ width: 390, height: 900 });
  await p.evaluate(() => document.querySelector('#pg0 .card:nth-child(2)').scrollIntoView());
  await p.screenshot({ path: OUT + '/backup.png' });
  console.log('Screenshot:', OUT + '/backup.png');
  await b.close();
})().catch(e => { console.error(e); process.exit(1); });
