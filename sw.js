// ── AUTO-VERSIONED — do not edit this line manually ──
// Version is injected by build.js at deploy time
const VERSION = 'iv-1776345153235';
const ASSETS  = ['/', '/index.html', '/manifest.json', '/icon-192.png', '/icon-512.png', '/avatar.png'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(VERSION)
      .then(c => c.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== VERSION).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );

  // Tell all open tabs a new version is live
  self.clients.matchAll({ type: 'window' }).then(clients =>
    clients.forEach(c => c.postMessage({ type: 'SW_UPDATED', version: VERSION }))
  );
});

// Network first — always fetch fresh, cache as fallback for offline
self.addEventListener('fetch', e => {
  const url = e.request.url;
  if (
    url.includes('api.anthropic.com')    ||
    url.includes('script.google.com')    ||
    url.includes('fonts.googleapis.com') ||
    url.includes('fonts.gstatic.com')    ||
    url.includes('cdnjs.cloudflare.com') ||
    e.request.method !== 'GET'
  ) return;

  e.respondWith(
    fetch(e.request)
      .then(res => {
        if (res && res.ok) {
          const clone = res.clone();
          caches.open(VERSION).then(c => c.put(e.request, clone));
        }
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
