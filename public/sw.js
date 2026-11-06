/* Service worker: hanya menyimpan aset statis milik sendiri (/assets, /icons).
   Data Firestore, autentikasi, dan permintaan lintas-domain TIDAK disimpan di cache. */
const CACHE = 'perpus-v2';
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== self.location.origin) return;
  if (u.pathname.startsWith('/assets/') || u.pathname.startsWith('/icons/')) {
    e.respondWith(caches.open(CACHE).then(c => c.match(r).then(m => m || fetch(r).then(res => { if (res.ok) c.put(r, res.clone()); return res; }))));
  }
});
