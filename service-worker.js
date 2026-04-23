/**
 * Pelichet NLC — Service Worker
 * Stratégie : network-first pour les ressources applicatives, fallback cache.
 * Les appels Google Apps Script (SCRIPT_URL) passent toujours réseau (jamais cachés).
 */

const CACHE_VERSION = 'pelichet-v3.0.0';
const APP_SHELL = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './manifest.json',
  './icon.svg',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&display=swap'
];

// Install : pré-cache de l'app shell
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(APP_SHELL.map(u => new Request(u, { credentials: 'same-origin' }))))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting())
  );
});

// Activate : nettoyage anciens caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(
      keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

// Fetch : bypass Google Apps Script ; network-first pour le reste
self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = req.url;

  // Ne jamais cacher les appels API Google Apps Script
  if (url.includes('script.google.com') || url.includes('googleapis.com/auth') || req.method !== 'GET') {
    return; // laisse passer en réseau
  }

  // Network-first pour HTML / JS / CSS
  if (url.endsWith('.html') || url.endsWith('.js') || url.endsWith('.css') || url.endsWith('/')) {
    event.respondWith(
      fetch(req).then((resp) => {
        const copy = resp.clone();
        caches.open(CACHE_VERSION).then(cache => cache.put(req, copy));
        return resp;
      }).catch(() => caches.match(req))
    );
    return;
  }

  // Cache-first pour fonts, images, icônes
  event.respondWith(
    caches.match(req).then((cached) => cached || fetch(req).then((resp) => {
      const copy = resp.clone();
      if (resp.ok) caches.open(CACHE_VERSION).then(cache => cache.put(req, copy));
      return resp;
    }))
  );
});
