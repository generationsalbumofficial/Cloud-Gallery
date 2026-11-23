const C = 'pg-v3';
self.addEventListener('install', e => { e.waitUntil(caches.open(C).then(c => c.addAll(['./', 'index.html'])).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET') return;
  if (u.hostname === 'drive.google.com' && u.pathname.startsWith('/thumbnail')) return;   // video thumbnails can be 'not ready yet': never cache them
  if (u.hostname === 'drive.google.com' && u.pathname.includes('/preview')) return;      // video player
  if (u.hostname.endsWith('googleusercontent.com') || u.hostname === 'drive.google.com') {   // photos: cache first
    e.respondWith(caches.open(C).then(async c => {
      const hit = await c.match(r); if (hit) return hit;
      try { const res = await fetch(r); if (res.ok || res.type === 'opaque') { await c.put(r, res.clone()); c.keys().then(k => { if (k.length > 150) c.delete(k[0]); }); } return res; }
      catch (err) { return Response.error(); }
    }));
  } else if (u.origin === location.origin) {   // app files: network first, cached copy when offline
    e.respondWith(fetch(r).then(res => { const cp = res.clone(); caches.open(C).then(c => c.put(r, cp)); return res; }).catch(() => caches.match(r).then(x => x || caches.match('index.html'))));
  }
});
