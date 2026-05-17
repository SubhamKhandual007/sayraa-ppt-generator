const CACHE_NAME = 'suusri-ai-v1';
const ASSETS_TO_CACHE = [
  '/',
  '/static/styles.css',
  '/static/script.js',
  '/static/logo.png',
  '/static/manifest.json'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        return cache.addAll(ASSETS_TO_CACHE);
      })
  );
});

self.addEventListener('fetch', (event) => {
  // Simple fetch handler to satisfy PWA installability requirements
  event.respondWith(
    caches.match(event.request)
      .then((response) => {
        return response || fetch(event.request);
      })
  );
});
