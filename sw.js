/* سيّارتي — service worker: keeps the web version working offline.
   Only the app's own files are cached; nothing is ever sent anywhere. */
const APPV = '44479f6081da';
const VERSION = 'sayyarati-v' + APPV;
/* files the app loads on demand (Excel and PDF), with the version in their address — written by tools/stamp.py */
const LAZY = ["./vendor/exceljs.min.js?v=Pqp51FUN2qz", "./vendor/pdf-lib.min.js?v=weMABwrltA6j", "./vendor/fontkit.umd.min.js?v=2p6U1mmqF10", "./vendor/bidi.min.js?v=wQtTdqzwTFhz", "./report.js?v=4BZRXPnXfxil", "./fonts/readex-report-400.ttf?v=ZM9Zf46J0OTw", "./fonts/readex-report-700.ttf?v=3xKwULn0RBRa"];
const ASSETS = ['./', './index.html', './app.js?v=' + APPV, './manifest.webmanifest',
  './icon-192.png', './icon-512.png', './icon-maskable-512.png', './apple-touch-icon.png', './favicon.png',
  ...LAZY,
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
/* only ever read from this app's own cache (other pages on the same site have their own) */
const fromCache = req => caches.open(VERSION).then(c => c.match(req));
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
    }).catch(() => fromCache('./index.html')));
    return;
  }
  if (!KNOWN.has(url.href)) return;
  e.respondWith(fromCache(req).then(hit => hit || fetch(req).then(r => {
    if (cacheable(r)) { const cp = r.clone(); caches.open(VERSION).then(c => c.put(req, cp)); }
    return r;
  })));
});
