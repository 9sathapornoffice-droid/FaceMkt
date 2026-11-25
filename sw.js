const CACHE_NAME = 'isuzu-facescan-v1';
const STATIC_ASSETS = [
  'index.html',
  'scan.html',
  'register.html',
  'config.html',
  'manifest.json',
  'js/api-config.js',
  'img/mascot.png',
  'img/biturbo-logo.png',
  'img/icon-192.png',
  'img/icon-512.png',
  'img/apple-touch-icon.png'
];

// Install: Cache core assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(STATIC_ASSETS).catch(err => {
        console.warn('Cache addAll warning:', err);
      });
    }).then(() => self.skipWaiting())
  );
});

// Activate: Clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch: Network First for API and dynamic pages, fallback to cache
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Bypass Google Apps Script API calls and Face-API external models
  if (url.origin.includes('script.google.com') || url.origin.includes('github.io')) {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Cache successful GET responses
        if (event.request.method === 'GET' && response.status === 200) {
          const respClone = response.clone();
          caches.open(CACHE_NAME).then(cache => {
            cache.put(event.request, respClone);
          });
        }
        return response;
      })
      .catch(() => {
        // Fallback to cache when offline
        return caches.match(event.request);
      })
  );
});
