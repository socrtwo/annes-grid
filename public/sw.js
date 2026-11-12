// Anne's Grid service worker: makes the web app (and the ChromeOS / desktop-browser install) work
// offline. Pages are network-first so an update shows on the next visit; other same-origin files
// (hashed scripts, fonts, dictionaries) are served from the cache and refreshed in the background.
const CACHE = 'quillgrid-v1';
const SHELL = './index.html';
// A stored file answers every request for its URL: a Vary header (servers send Vary: Origin) must not hide it when offline.
const ANY = { ignoreVary: true };

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(['./', SHELL, './manifest.webmanifest'])).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

// Only complete, successful, same-origin responses are kept (a partial 206 cannot be stored).
const storable = (res) => res && res.status === 200 && res.type === 'basic';

/** True for the app page itself (the scope root or index.html), as opposed to another document such as the notices file. */
function isShell(url) {
  const scope = new URL(self.registration.scope).pathname;
  return url.pathname === scope || url.pathname === scope + 'index.html';
}

// The page sends the files it loaded before this worker controlled it (the scripts and styles of the very first
// visit), so the app works offline from then on rather than only after a second visit.
self.addEventListener('message', (e) => {
  const d = e.data;
  if (!d || d.type !== 'precache' || !Array.isArray(d.urls)) return;
  const urls = d.urls.filter((u) => typeof u === 'string' && new URL(u, self.location.href).origin === self.location.origin).slice(0, 400);
  e.waitUntil(caches.open(CACHE).then((c) => Promise.all(urls.map((u) => c.match(u, ANY).then((hit) => hit || fetch(u).then((res) => { if (storable(res)) return c.put(u, res); }).catch(() => undefined))))));
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then((res) => {
      // Only the app page refreshes the cached shell: any other document must not replace it.
      if (storable(res)) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(isShell(url) ? SHELL : req, copy)); }
      return res;
    }).catch(() => (isShell(url) ? caches.match(SHELL, ANY) : caches.match(req, ANY).then((hit) => hit || caches.match(SHELL, ANY)))));
    return;
  }
  e.respondWith(caches.open(CACHE).then(async (c) => {
    const hit = await c.match(req, ANY);
    const fresh = fetch(req).then((res) => { if (storable(res)) c.put(req, res.clone()); return res; }).catch(() => hit);
    return hit || fresh;
  }));
});
