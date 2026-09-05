// Service worker — app-shell caching, with a hard guarantee that stock
// data itself is never served stale.
//
// Strategy:
//  - Anything hitting the Apps Script backend (script.google.com /
//    script.googleusercontent.com), and any non-GET request, is ALWAYS
//    fetched from the network and never cached — stock numbers must never
//    come from a cache. This preserves the original "no caching of data"
//    guarantee exactly.
//  - Everything else — the app shell (index.html, manifest.json) and the
//    static libraries it loads — is cached so a repeat visit is instant,
//    and an install with no signal shows the last-seen shell instead of a
//    blank white screen. Cache-first, with a background refetch to keep
//    the cache warm for next time.
//
// Bump CACHE_VERSION whenever index.html / manifest.json / this file
// changes, so an already-installed client picks up the new shell instead
// of being stuck on an old cached one. Keep this in step with APP_VERSION
// in index.html.
const CACHE_VERSION = 'calgas-shell-v17';
const APP_SHELL = [
  './index.html',
  './manifest.json',
  './logo/icon-192.png',
  './logo/icon-512.png',
  './logo/icon-512-maskable.png',
  './logo/CALGAS%20CAPACITORS-logo-768x240.jpg',
  './vendor/chart.umd.min.js',
  './vendor/xlsx.full.min.js'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_VERSION)
      .then((cache) => cache.addAll(APP_SHELL))
      .catch(() => {}) // don't block install if a shell asset is briefly unreachable
  );
  self.skipWaiting(); // don't wait for every open tab to close before a
                       // newly deployed version of this file takes over
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((names) => Promise.all(names.filter((n) => n !== CACHE_VERSION).map((n) => caches.delete(n))))
      .then(() => self.clients.claim()) // take control of already-open tabs
  );
});

function isBackendRequest(url) {
  return url.hostname === 'script.google.com' || url.hostname === 'script.googleusercontent.com';
}

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);

  // Live stock data — always network, never cached or served from cache.
  if (isBackendRequest(url) || e.request.method !== 'GET') {
    e.respondWith(fetch(e.request));
    return;
  }

  // App shell & static assets — cache-first, refresh in the background so
  // the next visit has the latest version too.
  e.respondWith(
    caches.open(CACHE_VERSION).then((cache) =>
      cache.match(e.request).then((cached) => {
        const network = fetch(e.request)
          .then((resp) => {
            if (resp && resp.ok) cache.put(e.request, resp.clone());
            return resp;
          })
          .catch(() => cached); // offline with nothing fresh — fall back to cache if we have it
        return cached || network;
      })
    )
  );
});
