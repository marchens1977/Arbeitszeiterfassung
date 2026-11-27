/* Prüft das Ausscheiden unter dem Jahr (Monate mit Jahres-Basis 0):
   anteiliger Urlaub (gesetzlich gerundet), Urlaubsstunden, Soll, Monate „kein Vertrag“,
   Urlaubs-Karten rot/grün. Die erwarteten Soll-Werte stammen aus der Excel-Datei
   (python3 werkzeuge/excel-pruefen.py --ab-monat N [--anspruch T]) – App und Excel müssen gleich rechnen.
   Start: node werkzeuge/ausscheiden.test.js */
const { chromium } = require('playwright'); const path = require('path');
const APP = 'file://' + path.join(__dirname, '..', 'index.html');

const faelle = [
  // abMonat = erster Monat mit 0 (1-12), Anspruch, erwartete Tage, erwartetes Soll (aus Excel)
  { abMonat: 7, vac: '30', tage: 15,    urlaub: '117:00:00', soll: '897:00:00' },
  { abMonat: 8, vac: '30', tage: 18,    urlaub: '140:24:00', soll: '1042:36:00' },
  { abMonat: 6, vac: '30', tage: 13,    urlaub: '101:24:00', soll: '743:36:00' },
  { abMonat: 2, vac: '30', tage: 3,     urlaub: '23:24:00',  soll: '145:36:00' },
  { abMonat: 8, vac: '31', tage: 18.08, urlaub: '141:01:26', soll: '1041:58:34' },
  { abMonat: 13, vac: '30', tage: 30,   urlaub: '234:00:00', soll: '1794:00:00' },
];

(async () => {
  const b = await chromium.launch(); const p = await b.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message));
  p.on('dialog', d => { console.log('Dialog:', d.message()); d.accept(); });
  await p.goto(APP);
  let fehler = 0;
  const pruefe = (name, ist, soll) => { const ok = String(ist) === String(soll); if (!ok) fehler++;
    console.log(`${ok ? 'OK  ' : 'FEHLER'} ${name}: ${ist}${ok ? '' : ' (erwartet ' + soll + ')'}`); };

  for (const f of faelle) {
    const r = await p.evaluate(f => {
      st.vac = f.vac;
      st.monthlyAnnual = Array.from({ length: 12 }, (_, i) => i === 0 ? '2028' : (i + 1 >= f.abMonat ? '0' : ''));
      st.days = {}; save();
      const x = computeStats();
      showTab(0);
      return { tage: x.totalClaimDays, urlaub: fmtT(x.totalClaimHoursSec), soll: fmtT(x.nettoTargetSec),
               juli: fmtT(x.planNet[6]), info: $('vacInfo').innerText };
    }, f);
    console.log(`--- Ausscheiden ab Monat ${f.abMonat}, Anspruch ${f.vac}: ${r.info || '(volles Jahr, kein Hinweis)'}`);
    pruefe('Urlaubstage', r.tage, f.tage);
    pruefe('Urlaubsstunden', r.urlaub, f.urlaub);
    pruefe('Soll', r.soll, f.soll);
    if (f.abMonat <= 7) pruefe('Juli Plan', r.juli, '00:00:00');
  }

  // Dateneingabe: Juli zeigt „kein Vertrag“; Urlaubs-Karten rot/grün
  await p.evaluate(() => { st.vac = '30'; st.monthlyAnnual = ['2028','','','','','','0','','','','','']; st.days = {}; save(); });
  await p.evaluate(() => showTab(3));
  pruefe('Dateneingabe 1.7.', (await p.$eval('#row_2027-07-01', e => e.innerText)).replace(/\s+/g, ' ').trim(), 'Do 01.07. kein Vertrag – 00:00:00');
  await p.evaluate(() => showTab(1));
  const vorgabe = await p.evaluate(() => computeStats().takenDays);
  console.log('Vorbelegter Urlaub Jan–Jun:', vorgabe, 'Tage');
  pruefe('Karten-Farbe bei Abweichung', await p.$eval('#vacCards', e => e.className), 'grid-cards vac-diff');
  // genau 15 Urlaubstage eintragen: vorbelegten Urlaub entfernen, dann 15 Arbeitstage im Januar/Februar
  await p.evaluate(() => {
    days.forEach(d => { if (d.m < 6 && getDayData(d).s === 'U') st.days[d.k] = { s: 'X', p: '0' }; });
    let n = 0; days.forEach(d => { if (n < 15 && d.m < 2 && !holidays[d.k]) { st.days[d.k] = { s: 'U', p: '0' }; n++; } });
    save(); showTab(1);
  });
  pruefe('Genommen', await p.evaluate(() => computeStats().takenDays), 15);
  pruefe('Karten-Farbe bei genau 15', await p.$eval('#vacCards', e => e.className), 'grid-cards vac-ok');
  // 16. Urlaubstag muss abgelehnt werden
  await p.evaluate(() => showTab(3));
  await p.evaluate(() => setDay('2027-03-01', 's', 'U'));
  pruefe('16. Tag abgelehnt', await p.evaluate(() => computeStats().takenDays), 15);

  console.log('errors:', errors);
  console.log(fehler === 0 ? 'ERGEBNIS: alles in Ordnung' : `ERGEBNIS: ${fehler} Fehler`);
  await b.close();
  process.exit(fehler || errors.length ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
