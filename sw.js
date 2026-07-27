/* Service Worker
   一度開いたあとは通信なしで動くようにする。
   アプリを更新したら CACHE の版数を上げること。 */

const CACHE = 'slide-generator-v1';

const ASSETS = [
  './',
  './index.html',
  './style.css',
  './app.js',
  './vendor/pptxgen.bundle.js',
  './manifest.webmanifest',
  './icon.svg',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE)
      .then((cache) => cache.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // 参照のみ。POST等や外部ドメインには手を出さない
  if (req.method !== 'GET') return;
  if (new URL(req.url).origin !== self.location.origin) return;

  // キャッシュを即返しつつ裏で更新する（オフラインでも動き、次回起動時に最新になる）
  event.respondWith(
    caches.match(req).then((cached) => {
      const fresh = fetch(req)
        .then((res) => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then((cache) => cache.put(req, copy));
          }
          return res;
        })
        .catch(() => cached || caches.match('./index.html'));

      return cached || fresh;
    })
  );
});
