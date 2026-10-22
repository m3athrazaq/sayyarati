/* سيّارتي — service worker: keeps the app working offline */
const VERSION = 'sayyarati-v1.0.0-1';
const ASSETS = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './icon-maskable-512.png', './apple-touch-icon.png', './favicon.png',
  './vendor/exceljs.min.js',
  './fonts/readex-pro-arabic-300-normal.woff2', './fonts/readex-pro-arabic-400-normal.woff2', './fonts/readex-pro-arabic-500-normal.woff2', './fonts/readex-pro-arabic-700-normal.woff2',
  './fonts/readex-pro-latin-300-normal.woff2', './fonts/readex-pro-latin-400-normal.woff2', './fonts/readex-pro-latin-500-normal.woff2', './fonts/readex-pro-latin-700-normal.woff2'];
self.addEventListener('install', e => e.waitUntil(caches.open(VERSION).then(c => c.addAll(ASSETS.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting())));
self.addEventListener('activate', e => e.waitUntil((async () => {
  for (const k of await caches.keys()) if (k.startsWith('sayyarati-') && k !== VERSION) await caches.delete(k);
  await self.clients.claim();
})()));
self.addEventListener('fetch', e => {
  const req = e.request; if (req.method !== 'GET') return;
  const url = new URL(req.url); if (url.origin !== location.origin) return;
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(r => { const cp = r.clone(); caches.open(VERSION).then(c => c.put('./index.html', cp)); return r; }).catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => { if (r.ok) { const cp = r.clone(); caches.open(VERSION).then(c => c.put(req, cp)); } return r; })));
});
