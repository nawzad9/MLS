/* MLS Course Registration — service worker
   The app shell works offline; data always comes live from the Apps Script API (never cached). */
const VERSION = 'mls-v3.0.0';
const SHELL = [
  './', 'index.html', 'manifest.webmanifest',
  'css/app.css', 'js/i18n.js', 'js/exporters.js', 'js/app.js',
  'icons/icon.svg', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/maskable-512.png',
  'icons/apple-touch-icon.png', 'icons/favicon-32.png'
];
const FONT_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(VERSION).then(cache => cache.addAll(SHELL.map(u => new Request(u, { cache: 'reload' })))));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// The page asks the waiting worker to take over when the user presses "Reload".
self.addEventListener('message', event => {
  if (event.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;                       // API calls are POST — untouched
  const url = new URL(req.url);
  const sameOrigin = url.origin === self.location.origin;
  if (!sameOrigin && !FONT_HOSTS.includes(url.hostname)) return;

  if (req.mode === 'navigate') {                          // pages: network first, cached shell offline
    event.respondWith(
      fetch(req).then(res => { putInCache(req, res.clone()); return res; })
        .catch(() => caches.match('index.html'))
    );
    return;
  }
  // assets and fonts: answer from cache, refresh in the background
  event.respondWith(
    caches.match(req).then(cached => {
      const fresh = fetch(req).then(res => { if (res.ok || res.type === 'opaque') putInCache(req, res.clone()); return res; })
        .catch(() => cached);
      return cached || fresh;
    })
  );
});

function putInCache(req, res) {
  caches.open(VERSION).then(cache => cache.put(req, res)).catch(() => {});
}
