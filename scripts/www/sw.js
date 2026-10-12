/**
 * Service Worker：缓存 app shell，保证安卓上离线可用（添加到主屏幕后像原生 App）
 */
var CACHE = 'shuatikaoshi-v11';
var ASSETS = [
  './',
  './index.html',
  './styles.css',
  './manifest.json',
  './icon.svg',
  './lib/xlsx.full.min.js',
  './lib/mammoth.browser.js',
  './js/parser.js',
  './js/parser.worker.js',
  './js/judge.js',
  './js/db.js',
  './js/app.js'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) {
      return c.addAll(ASSETS);
    }).then(function () {
      return self.skipWaiting();
    })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) {
        return k === CACHE ? null : caches.delete(k);
      }));
    }).then(function () {
      return self.clients.claim();
    })
  );
});

self.addEventListener('fetch', function (e) {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request).then(function (hit) {
      if (hit) return hit;
      return fetch(e.request).then(function (res) {
        if (res && res.status === 200 && res.type === 'basic') {
          var copy = res.clone();
          caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
        }
        return res;
      }).catch(function () {
        return caches.match('./index.html');
      });
    })
  );
});
