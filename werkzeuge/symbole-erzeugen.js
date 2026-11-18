/* Erzeugt aus icons/icon.svg alle PNG-Symbole (192, 512, 512 maskable, 180 Apple).
   Nur nötig, wenn das Symbol geändert wurde.  Start: node werkzeuge/symbole-erzeugen.js */
const { chromium } = require('playwright'); const fs = require('fs'); const path = require('path');
const ICONS = path.join(__dirname, '..', 'icons');
(async () => {
  const svg = fs.readFileSync(path.join(ICONS, 'icon.svg'), 'utf8');
  const b = await chromium.launch(); const p = await b.newPage();
  const make = async (size, file, maskable) => {
    await p.setViewportSize({ width: size, height: size });
    // maskable: Symbol kleiner auf vollflächigem Hintergrund (Android schneidet Ränder ab)
    const inner = maskable ? Math.round(size * 0.8) : size;
    const off = Math.round((size - inner) / 2);
    const body = maskable
      ? `<div style="width:${size}px;height:${size}px;background:#2563eb"><div style="position:absolute;left:${off}px;top:${off}px;width:${inner}px;height:${inner}px">${svg.replace('rx="110" fill="#2563eb"', 'fill="none"')}</div></div>`
      : svg;
    await p.setContent(`<html><body style="margin:0;background:transparent">${body}</body></html>`);
    await p.screenshot({ path: path.join(ICONS, file), omitBackground: true, clip: { x: 0, y: 0, width: size, height: size } });
  };
  await make(192, 'icon-192.png'); await make(512, 'icon-512.png'); await make(512, 'icon-512-maskable.png', true); await make(180, 'apple-touch-icon.png');
  await b.close();
})();
