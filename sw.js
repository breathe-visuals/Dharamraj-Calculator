const CACHE_NAME = 'jewellery-calc-v1';
const urlsToCache = [
  './',
  './index.html',
  './Media/site.webmanifest',
  './Media/android-chrome-192x192.png',
  './Media/android-chrome-512x512.png',
  './Media/apple-touch-icon.png',
  './Media/favicon-16x16.png',
  './Media/favicon-32x32.png',
  './Media/favicon.ico'
];

self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache) {
        return cache.addAll(urlsToCache);
      })
  );
});

self.addEventListener('fetch', function(event) {
  event.respondWith(
    caches.match(event.request)
      .then(function(response) {
        if (response) {
          return response;
        }
        return fetch(event.request);
      }
    )
  );
});
