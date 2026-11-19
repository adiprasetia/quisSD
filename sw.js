/* PintarSD service worker. Naikkan nomor V jika index.html berubah dan ingin dipaksa segar. */
const V = 'pintarsd-v2', IMG = 'pintarsd-img';
const CORE = ['./', 'index.html', 'manifest.json', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V && k !== IMG).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.hostname === 'script.google.com') return;
  const isImg = r.destination === 'image';
  const isFont = u.hostname === 'fonts.googleapis.com' || u.hostname === 'fonts.gstatic.com';
  if (isImg || isFont) {
    e.respondWith(caches.open(isImg ? IMG : V).then(c => c.match(r).then(hit => hit || fetch(r).then(n => {
      if (n.ok || n.type === 'opaque') c.put(r, n.clone());
      return n;
    }))));
    return;
  }
  if (u.origin === location.origin) {
    e.respondWith(fetch(r).then(n => {
      const cp = n.clone(); caches.open(V).then(c => c.put(r, cp)); return n;
    }).catch(() => caches.match(r).then(hit => hit || caches.match('index.html'))));
  }
});
