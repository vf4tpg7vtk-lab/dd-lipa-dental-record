const CACHE = 'dentarecord-v1';
const ASSETS = [
  './',
  './index.html',
  './css/app.css',
  './js/app.js',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;500;600;700&family=DM+Mono:wght@400;500;600&family=Crimson+Pro:ital,wght@0,300;0,400;0,600;1,300;1,400&display=swap'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  if (e.request.url.includes('graph.microsoft.com')) return; // never cache API calls
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).then(res => {
      if (res && res.status === 200 && res.type === 'basic') {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
      }
      return res;
    }).catch(() => caches.match('./index.html')))
  );
});
