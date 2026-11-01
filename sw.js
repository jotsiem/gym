// Caché de la app para que abra al instante y funcione sin cobertura. Los datos van siempre a la red.
var CACHE = 'gym-v4.0';
var SHELL = ['./', 'index.html', 'config.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png',
             'https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js'];
self.addEventListener('install', function (e) { e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); })); });
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var r = e.request;
  if (r.method !== 'GET' || /script\.google(usercontent)?\.com/.test(r.url)) return;   // la API nunca se cachea
  e.respondWith(fetch(r).then(function (resp) {   // red primero (para recibir actualizaciones), caché si no hay red
    var copia = resp.clone(); caches.open(CACHE).then(function (c) { c.put(r, copia); }); return resp;
  }).catch(function () { return caches.match(r, { ignoreSearch: true }).then(function (m) { return m || caches.match('index.html'); }); }));
});
