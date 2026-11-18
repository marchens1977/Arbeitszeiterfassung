/* Prüft die installierbare App: Manifest, Symbole, Offline-Helfer (Service Worker) und Start ohne Internet.
   Startet dafür selbst einen kleinen Webserver auf Port 8765.  Start: node werkzeuge/offline-app.test.js */
const { chromium } = require('playwright'); const path = require('path'); const { spawn } = require('child_process');
const server = spawn('python3', ['-m', 'http.server', '8765'], { cwd: path.join(__dirname, '..'), stdio: 'ignore' });
process.on('exit', () => server.kill());
(async () => {
  await new Promise(r => setTimeout(r, 1000));
  const b = await chromium.launch(); const ctx = await b.newContext(); const p = await ctx.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => m.type()==='error' && errors.push(m.text()));
  await p.goto('http://localhost:8765/index.html');
  const m = await (await p.goto('http://localhost:8765/manifest.json')).json(); console.log('manifest:', m.name, m.icons.length, 'icons');
  for (const i of m.icons) console.log(' icon', i.src, (await p.request.get('http://localhost:8765/' + i.src)).status());
  await p.goto('http://localhost:8765/index.html');
  await p.waitForFunction(() => navigator.serviceWorker.controller || navigator.serviceWorker.getRegistration().then(r => r && r.active), null, { timeout: 10000 }).catch(()=>{});
  await p.reload(); await p.waitForTimeout(800);
  console.log('SW aktiv:', await p.evaluate(() => !!navigator.serviceWorker.controller));
  console.log('cache keys:', await p.evaluate(async () => { const c = await caches.keys().then(k => caches.open(k[0])); return (await c.keys()).length; }));
  await ctx.setOffline(true);
  await p.reload(); await p.waitForTimeout(500);
  console.log('offline Titel:', await p.title(), '| Tabs:', await p.$$eval('.tab', t => t.length));
  await ctx.setOffline(false);
  console.log('errors:', errors);
  await b.close(); server.kill();
})().catch(e => { console.error(e); process.exit(1); });
