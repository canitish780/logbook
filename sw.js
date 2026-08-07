/**
 * Logbook service worker.
 *
 * Scope: caches the app shell (HTML/manifest/icons/fonts) so the app itself
 * opens instantly with no network. It deliberately does NOT touch requests
 * to the Google Apps Script backend (script.google.com) — those are handled
 * by the page's own online/offline + localStorage-cache logic in index.html,
 * which already knows how to queue changes and retry. The service worker's
 * job is just "can the app load at all", not "is my data fresh".
 *
 * Bump CACHE_VERSION whenever you ship a change to index.html/manifest/icons
 * so clients pick up the new shell instead of a stale cached copy.
 */
const CACHE_VERSION = 'logbook-shell-v6';

const CORE_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon-72.png',
  './icons/icon-96.png',
  './icons/icon-128.png',
  './icons/icon-144.png',
  './icons/icon-152.png',
  './icons/icon-192.png',
  './icons/icon-384.png',
  './icons/icon-512.png',
  './icons/icon-maskable-192.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-16.png',
  './icons/favicon-32.png'
];

const BACKEND_HOSTS = ['script.google.com', 'script.googleusercontent.com'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then((cache) => cache.addAll(CORE_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((names) => Promise.all(
        names.filter((n) => n !== CACHE_VERSION).map((n) => caches.delete(n))
      ))
      .then(() => self.clients.claim())
  );
});

// Let the page force-activate an updated worker immediately (used by the
// "Update available" toast in index.html).
self.addEventListener('message', (event) => {
  if (event.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return; // never intercept POSTs to the Sheets backend

  const url = new URL(req.url);

  // Never cache/intercept the Apps Script API — always hit the real network
  // so the page's own try/catch + offline-queue logic sees real failures.
  if (BACKEND_HOSTS.includes(url.hostname)) return;

  // App shell + same-origin assets: cache-first, refresh cache in background.
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(req).then((cached) => {
        const network = fetch(req).then((res) => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE_VERSION).then((cache) => cache.put(req, copy));
          }
          return res;
        }).catch(() => cached);
        return cached || network;
      })
    );
    return;
  }

  // Cross-origin static assets (Google Fonts CSS + font files): cache-first,
  // so once loaded once, they render offline too.
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    event.respondWith(
      caches.match(req).then((cached) => {
        if (cached) return cached;
        return fetch(req).then((res) => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE_VERSION).then((cache) => cache.put(req, copy));
          }
          return res;
        }).catch(() => cached);
      })
    );
  }
});
