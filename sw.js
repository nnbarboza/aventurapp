/* AVENTURApp — Service Worker
 * HTML: network-first (los despliegues se ven al recargar).
 * Resto (imágenes, manifest): stale-while-revalidate.
 * Nunca cachea el backend de Apps Script.
 * Al subir una versión: cambiar CACHE aquí y APP_VERSION en index.html a la vez.
 */
const CACHE = 'aventurapp-v0.6.0';

self.addEventListener('install', function () { self.skipWaiting(); });

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys()
      .then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); })); })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.hostname.indexOf('script.google') >= 0) return;
  if (url.hostname.indexOf('googleusercontent') >= 0) return;
  if (url.origin !== location.origin) return;

  const esHTML = req.mode === 'navigate' || (req.headers.get('accept') || '').indexOf('text/html') >= 0;

  /* Solo se guarda lo que vino bien: un 404 cacheado taparía la imagen cuando la subas. */
  function guardar(r) {
    if (r && r.ok && r.type !== 'opaque') { const c = r.clone(); caches.open(CACHE).then(function (ch) { ch.put(req, c); }); }
    return r;
  }

  if (esHTML) {
    e.respondWith(fetch(req).then(guardar).catch(function () {
      return caches.match(req).then(function (r) { return r || caches.match('./index.html'); });
    }));
    return;
  }
  e.respondWith(caches.match(req).then(function (cached) {
    const net = fetch(req).then(guardar).catch(function () { return cached; });
    return cached || net;
  }));
});
