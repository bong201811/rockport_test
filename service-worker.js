// 간이실사 스타일 전체 사이즈 비교 사이즈별 부족 내역 v1.3 - 모바일 PWA 캐시 갱신
const CACHE_NAME = 'rbo-inventory-v4-style-size-details';
const ASSETS = [
  './',
  './index.html',
  './app.js',
  './partial-mode.js',
  './xlsx.full.min.js',
  './rock.png',
  './favicon.png',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)));
  self.skipWaiting();
});
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(key => key.startsWith('rbo-inventory-') && key !== CACHE_NAME).map(key => caches.delete(key))))
  );
  self.clients.claim();
});
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        return response;
      }).catch(() => caches.match('./index.html'));
    })
  );
});
