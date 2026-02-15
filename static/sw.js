const CACHE_NAME = 'lt-claim-v1';
const PRECACHE = ['/', '/static/manifest.json', '/static/icon-192.svg', '/static/icon-512.svg'];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(PRECACHE)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  // Network-first for all requests (app needs live server for calculations/docx)
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
