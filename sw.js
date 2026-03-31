const CACHE_NAME = 'creditx-v1';
const ASSETS = [
  '/index.html',
  '/manifest.json'
];

// Instalar y cachear archivos base
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Activar y limpiar caches viejos
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Estrategia: red primero, cache como respaldo
self.addEventListener('fetch', event => {
  // Solo manejar peticiones GET
  if (event.request.method !== 'GET') return;

  // No interceptar peticiones a Firebase (necesitan estar online)
  const url = event.request.url;
  if (url.includes('firebase') || url.includes('googleapis') || url.includes('gstatic')) return;

  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Guardar copia en cache
        const copy = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
