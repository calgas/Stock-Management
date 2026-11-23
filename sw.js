// Service worker — app-shell caching, with a hard guarantee that stock
// data itself is never served stale.
//
// Strategy:
//  - Anything hitting the Apps Script backend (script.google.com /
//    script.googleusercontent.com), and any non-GET request, is ALWAYS
//    fetched from the network and never cached — stock numbers must never
//    come from a cache.
//  - Everything else — the app shell (index.html, manifest.json) and the
//    static libraries it loads — is cached so a repeat visit is instant,
//    and an install with no signal shows the last-seen shell instead of a
//    blank white screen. Cache-first, with a background refetch to keep
//    the cache warm for next time.
//
// The sibling apps (ELE Tracker and the rest) are served from the same
// origin, calgas.github.io, and Cache Storage is shared across an origin,
// not per app. Every cache this worker owns carries CACHE_PREFIX, and
// cleanup only touches caches with that prefix — deleting "everything that
// isn't mine" would wipe each sibling's offline copy whenever this worker
// updated. A new sibling app must use a prefix of its own.
//
// Bump CACHE_VERSION whenever index.html / manifest.json / this file
// changes, so an already-installed client picks up the new shell instead
// of being stuck on an old cached one. Keep this in step with APP_VERSION
// in index.html.
const CACHE_PREFIX = 'calgas-shell-';
const CACHE_VERSION = CACHE_PREFIX + 'v58';
const APP_SHELL = [
  './',
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
    // One file at a time: addAll() caches nothing at all if a single file
    // fails, which left an install with no offline shell over one missing icon.
    caches.open(CACHE_VERSION)
      .then((cache) => Promise.all(APP_SHELL.map((url) => cache.add(url).catch(() => {}))))
  );
  // NOT skipWaiting() here. Taking over immediately would swap the shell
  // under someone mid-entry, and on a shop floor that means a half-typed
  // transaction disappearing. The new worker waits; the page notices it and
  // offers a Reload, and only then does the message below let it through.
});

self.addEventListener('message', (e) => {
  if (e.data && e.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((names) => Promise.all(
        names
          .filter((n) => n.startsWith(CACHE_PREFIX) && n !== CACHE_VERSION)
          .map((n) => caches.delete(n))
      ))
      .then(() => self.clients.claim()) // take control of already-open tabs
  );
});

function isBackendRequest(url) {
  return url.hostname === 'script.google.com' || url.hostname === 'script.googleusercontent.com';
}

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);

  // API calls and anything that isn't a GET are left entirely to the
  // browser: no respondWith, so this worker isn't in their path at all.
  // Re-sending them from here cached nothing and added a layer that can
  // fail on its own — mid-update, or when the browser stops the worker
  // between a POST and the redirect Apps Script answers it with.
  if (isBackendRequest(url) || e.request.method !== 'GET') return;

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

// ---- Notifications (3.30.0) ----
// The server sends an empty push (nothing to decrypt); this worker then asks
// the server (whose address the page keeps in PUSH_CACHE) for the newest
// unread notification, naming this device by its push address, and shows it.
// If that fails, a plain line shows instead. Open pages are told, so their
// bell catches up.
const PUSH_CACHE = 'calgas-stock-push';
const ICON = './logo/icon-192.png';

async function latestNotification() {
  const cache = await caches.open(PUSH_CACHE);
  const saved = await cache.match(new URL('__push_session', self.registration.scope).href);
  if (!saved) return null;
  const cfg = await saved.json();
  const sub = await self.registration.pushManager.getSubscription();
  if (!cfg || !cfg.url || !sub) return null;
  const res = await fetch(cfg.url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action: 'pushPeek', endpoint: sub.endpoint })
  });
  const json = await res.json();
  return json && json.success ? json.data : null;
}

self.addEventListener('push', (e) => {
  e.waitUntil((async () => {
    let title = 'Stock Management', options = { body: 'Something needs your attention. Open the app to see it.', icon: ICON, badge: ICON, tag: 'stock-notification', data: {} };
    try {
      const d = await latestNotification();
      const n = d && d.items && d.items[0];
      if (n) {
        title = n.Title;
        options = {
          body: n.Body + (d.unread > 1 ? '\n(' + (d.unread - 1) + ' more unread)' : ''),
          icon: ICON, badge: ICON, tag: n.ID, timestamp: Number(n.Created_At) || Date.now(),
          data: { link: n.Link || '', id: n.ID }
        };
        if (self.navigator && self.navigator.setAppBadge) self.navigator.setAppBadge(d.unread).catch(() => {});
      }
    } catch (err) { /* show the plain line */ }
    const open = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    open.forEach((c) => c.postMessage({ type: 'notif' }));
    await self.registration.showNotification(title, options);
  })());
});

self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  const data = e.notification.data || {};
  e.waitUntil((async () => {
    const open = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    const here = open.filter((c) => c.url.indexOf(self.registration.scope) === 0)[0];
    if (here) {
      await here.focus();
      if (data.link) here.postMessage({ type: 'open-notif', link: data.link, id: data.id });
      return;
    }
    const url = new URL('./', self.registration.scope);
    if (data.link) url.hash = 'notif=' + encodeURIComponent(data.link) + '&id=' + encodeURIComponent(data.id || '');
    await self.clients.openWindow(url.href);
  })());
});
