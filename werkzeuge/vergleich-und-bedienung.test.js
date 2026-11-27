const { chromium } = require('playwright');
/* Vergleicht eine Referenzversion (Standard: index.html aus origin/main) mit der aktuellen index.html:
   alle berechneten Werte müssen gleich sein. Danach Bedienungstests an der aktuellen Version.
   Start: node werkzeuge/vergleich-und-bedienung.test.js [referenz.html] */
const path = require('path'); const fs = require('fs'); const os = require('os'); const { execSync } = require('child_process');
const ROOT = path.join(__dirname, '..');
const S = fs.mkdtempSync(path.join(os.tmpdir(), 'az-test-'));
const refArg = process.argv[2];
fs.writeFileSync(path.join(S, 'old.html'), refArg ? fs.readFileSync(refArg) : execSync('git show origin/main:index.html', { cwd: ROOT }));
fs.copyFileSync(path.join(ROOT, 'index.html'), path.join(S, 'new.html'));
const sample = {
  vac:'30', ct:'minus', cv:'12:30:00', agreedOt:'10', unpaidVac:'-3:15',
  monthlyAnnual:['2028','','','1900','','','','','','2100','',''],
  pattern:{1:{s:'10:45',p:'30'},2:{s:'11:00',p:'45'},3:{s:'X',p:'0'},4:{s:'09:30',p:'30'},5:{s:'12:00',p:'15'}},
  days:{'2027-01-04':{s:'08:00',p:'30'},'2027-01-05':{s:'X',p:'0'},'2027-02-01':{s:'U',p:'0'},'2027-03-15':{s:'10:00',p:'30'}},
  monthlyActual:{0:'150:00:00',3:'160,5'}
};
const text = (p, sel) => p.$eval(sel, e => e.innerText.replace(/\s+/g,' ').trim());
async function run(file, preload) {
  const b = await chromium.launch(); const ctx = await b.newContext(); const p = await ctx.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => m.type()==='error' && errors.push(m.text()));
  await p.goto('file://' + path.join(S, 'blank.html'));
  await p.evaluate(([k,v]) => { localStorage.clear(); if (v) localStorage.setItem(k, JSON.stringify(v)); }, ['arbeitszeit_hessen_persistence_v1', preload]);
  await p.goto('file://' + path.join(S, file));
  const out = { errors };
  out.sm0 = await text(p, '#sm0');
  await p.evaluate(() => showTab(1)); out.vac = await text(p, '#vacSummaryBox') + ' | ' + await text(p, '#vacMonthBody');
  await p.evaluate(() => showTab(2)); out.pat = await text(p, '#patternBody') + ' | ' + await text(p, '#patternSumBox');
  await p.evaluate(() => showTab(3)); out.daysSum = await text(p, '#monthSummaryBox');
  out.dayRows = await p.$$eval('#dayBody tr:not(.month-header-row)', rs => rs.map(r => r.innerText.replace(/\s+/g,' ').trim()).join('\n'));
  await p.evaluate(() => showTab(4)); out.out = await text(p, '#outBody') + ' | ' + await text(p, '#sm3');
  return { page: p, browser: b, out };
}
(async () => {
  require('fs').writeFileSync(path.join(S,'blank.html'), '<html></html>');
  for (const [label, pre] of [['default', null], ['sample', sample]]) {
    const o = await run('old.html', pre), n = await run('new.html', pre);
    const norm = s => s.replace(/\b(\d):(\d\d:\d\d)/g, '0$1:$2').replace(/ \(Plan\)/g,'').replace(/Vertraglich zu erfüllende Nettoarbeitszeit \(Soll\)/g,'SOLL').replace(/Tatsächlich geleistete Nettoarbeitszeit( \(Dateneingabe\))?/g,'IST').replace(/\(positiver Wert\)|\(negativer Wert\)|\(Übertrag[^)]*\)/g,'').replace(/\s+/g,' ');
    console.log(`\n=== ${label} ===  errors old:${JSON.stringify(o.out.errors)} new:${JSON.stringify(n.out.errors)}`);
    for (const k of ['sm0','vac','pat','daysSum','dayRows','out']) {
      const a = norm(o.out[k]), c = norm(n.out[k]);
      console.log(k, a === c ? 'IDENTISCH' : 'UNTERSCHIED');
      if (a !== c && k !== 'dayRows') { console.log(' ALT:', a); console.log(' NEU:', c); }
      if (a !== c && k === 'dayRows') { const al=a.split('\n'), cl=c.split('\n'); al.forEach((l,i)=>{ if(l!==cl[i]) console.log(' ALT:',l,'\n NEU:',cl[i]); }); }
    }
    await o.browser.close(); await n.browser.close();
  }
  // Interaktionstests neue Version
  const n = await run('new.html', sample); const p = n.page;
  await p.evaluate(() => showTab(3)); await p.click('#mainBtn3');
  console.log('\nbtn class editing:', await p.$eval('#mainBtn3', e => e.className));
  await p.selectOption('#row_2027-01-06 select >> nth=0', '12:00');
  await p.selectOption('#row_2027-01-06 select >> nth=1', '60');
  console.log('row after change:', await text(p, '#row_2027-01-06'));
  console.log('jan header:', await text(p, '#mh_0'));
  await p.click('#mainBtn3'); console.log('btn class saved:', await p.$eval('#mainBtn3', e => e.className));
  console.log('persisted (erwartet 12:00 / 60):', await p.evaluate(() => JSON.parse(localStorage.getItem('arbeitszeit_hessen_2027_v2')).days['2027-01-06']));
  // Urlaubsgrenze
  await p.evaluate(() => { st.vac = '1'; save(); }); await p.click('#mainBtn3');
  p.once('dialog', d => { console.log('dialog:', d.message()); d.accept(); });
  await p.selectOption('#row_2027-01-07 select >> nth=0', 'U');
  console.log('row 07 after blocked U:', await text(p, '#row_2027-01-07'));
  // Stammdaten normalisieren
  await p.evaluate(() => showTab(0)); await p.click('#mainBtn0');
  await p.fill('#cv', '7,5'); await p.fill('#agreedOt', '1:5'); await p.fill('#m_annual_1', '1950');
  console.log('live info Feb/Mar:', await text(p, '#info_1'), '//', await text(p, '#info_2'));
  await p.click('#mainBtn0');
  console.log('cv/agreedOt normalized:', await p.$eval('#cv', e=>e.value), await p.$eval('#agreedOt', e=>e.value));
  // Export
  const [dl] = await Promise.all([p.waitForEvent('download'), p.click('button:has-text("Sicherung herunterladen")')]);
  console.log('download:', dl.suggestedFilename());
  const fp = await dl.path(); const j = JSON.parse(require('fs').readFileSync(fp,'utf8'));
  console.log('export ok:', j.app, j.year, Object.keys(j.data).join(','));
  // Import
  await p.evaluate(() => { st.vac='5'; save(); });
  p.on('dialog', d => d.accept());
  await p.setInputFiles('#importFile', fp); await p.waitForLoadState('load'); await p.waitForTimeout(300);
  console.log('after import vac (erwartet 30):', await p.$eval('#vac', e=>e.value));
  // Musterwoche: kein Urlaub wählbar
  await p.evaluate(() => showTab(2)); await p.click('#mainBtn2');
  console.log('Tages-Uhrzeiten (erwartet 11:00 … 17:00):', await p.evaluate(() => timeList[0][0] + ' … ' + timeList[timeList.length - 1][0]));
  console.log('pattern options (erwartet X,11:00,11:15):', await p.$$eval('#patternBody tr:first-child select:first-child option', os => os.slice(0,3).map(o=>o.value).join(',')));
  console.log('errors:', n.out.errors);
  await n.browser.close();
})().catch(e => { console.error(e); process.exit(1); });
/* Ergebnis prüfen: In der Ausgabe darf bei 'default' und 'sample' kein UNTERSCHIED stehen (außer reinen Textänderungen) und 'errors' muss [] sein. */
