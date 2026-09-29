const CACHE = 'at-farm-pwa-v10-42';
const STATIC = [
  './config.js?v=10.42',
  './manifest.webmanifest?v=10.42',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(STATIC)));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(key => key !== CACHE).map(key => caches.delete(key)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return;

  // Navigasi/index selalu network-first agar wrapper lama tidak menetap.
  if (event.request.mode === 'navigate' || url.pathname.endsWith('/index.html') || url.pathname.endsWith('/at-farm-pwa/')) {
    event.respondWith(
      fetch(event.request, {cache:'no-store'})
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // Asset statis: network-first, cache sebagai cadangan.
  event.respondWith(
    fetch(event.request, {cache:'no-store'}).then(response => {
      if (response && response.ok) {
        const copy = response.clone();
        caches.open(CACHE).then(cache => cache.put(event.request, copy));
      }
      return response;
    }).catch(() => caches.match(event.request))
  );
});
