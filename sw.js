/* Offline-Helfer (Service Worker):
   Speichert die App-Dateien, damit die App auch ohne Internet startet.
   Bei jeder Änderung an der App die VERSION hochzählen, damit alle Geräte die neue Fassung bekommen. */
const VERSION = 'arbeitszeit-v4';
const FILES = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-512-maskable.png',
  './icons/apple-touch-icon.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

/* Die App fragt nach der installierten Version (Anzeige unter "App-Updates"). */
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'GET_VERSION' && event.ports[0]) event.ports[0].postMessage({ version: VERSION });
});

/* Erst das Netz fragen (damit Updates sofort ankommen), bei Fehler aus dem Speicher laden. */
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request, { cache: 'no-cache' })
      .then(resp => {
        const copy = resp.clone();
        caches.open(VERSION).then(c => c.put(event.request, copy)).catch(() => {});
        return resp;
      })
      .catch(() => caches.match(event.request, { ignoreSearch: true }).then(r => r || caches.match('./index.html')))
  );
});
