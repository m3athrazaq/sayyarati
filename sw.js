/* سيّارتي — service worker: keeps the web version working offline.
   Only the app's own files are cached; nothing is ever sent anywhere. */
const APPV = '1.1.0';
const VERSION = 'sayyarati-v' + APPV;
const ASSETS = ['./', './index.html', './app.js?v=' + APPV, './manifest.webmanifest',
  './icon-192.png', './icon-512.png', './icon-maskable-512.png', './apple-touch-icon.png', './favicon.png',
  './vendor/exceljs.min.js',
  './fonts/readex-pro-arabic-300-normal.woff2', './fonts/readex-pro-arabic-400-normal.woff2', './fonts/readex-pro-arabic-500-normal.woff2', './fonts/readex-pro-arabic-700-normal.woff2',
  './fonts/readex-pro-latin-300-normal.woff2', './fonts/readex-pro-latin-400-normal.woff2', './fonts/readex-pro-latin-500-normal.woff2', './fonts/readex-pro-latin-700-normal.woff2'];
const SCOPE = self.registration.scope;
const KNOWN = new Set(ASSETS.map(u => new URL(u, SCOPE).href));

self.addEventListener('install', e => e.waitUntil(
  caches.open(VERSION).then(c => c.addAll(ASSETS.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting())
));
self.addEventListener('activate', e => e.waitUntil((async () => {
  for (const k of await caches.keys()) if (k.startsWith('sayyarati-') && k !== VERSION) await caches.delete(k);
  await self.clients.claim();
})()));

const cacheable = r => r && r.ok && r.type === 'basic';
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin || !url.href.startsWith(SCOPE)) return;
  if (req.mode === 'navigate') {
    /* network first, so updates arrive; only a real copy of the app page is ever stored */
    e.respondWith(fetch(req).then(r => {
      const isApp = url.pathname === new URL(SCOPE).pathname || url.pathname.endsWith('/index.html');
      if (cacheable(r) && isApp) { const cp = r.clone(); caches.open(VERSION).then(c => c.put('./index.html', cp)); }
      return r;
    }).catch(() => caches.match('./index.html')));
    return;
  }
  if (!KNOWN.has(url.href)) return;
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => {
    if (cacheable(r)) { const cp = r.clone(); caches.open(VERSION).then(c => c.put(req, cp)); }
    return r;
  })));
});
