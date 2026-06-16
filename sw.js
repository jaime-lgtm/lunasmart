/* ══════════════════════════════════════════════════════════════
   LUNA SMART — Service Worker
   Hace la app instalable y permite que cargue sin internet.
   Estrategia:
     - App shell (HTML, íconos, fuentes): cache con actualización en segundo plano
     - Datos de Google Sheets / Apps Script: SIEMPRE de red (nunca cacheados)
   ══════════════════════════════════════════════════════════════ */

const CACHE = 'lunasmart-v2';
const SHELL = [
  '/index.html',
  '/manifest.json',
  '/icon-192.png',
  '/icon-512.png',
  '/apple-touch-icon.png',
];

// Instalar: precachear el app shell
self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

// Activar: limpiar caches viejos
self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// Fetch: decidir cómo responder cada petición
self.addEventListener('fetch', (e) => {
  const url = e.request.url;

  // 1. NUNCA cachear datos en vivo (Google Sheets API + Apps Script)
  //    Siempre van a la red para tener datos frescos.
  if (
    url.includes('script.google.com') ||
    url.includes('sheets.googleapis.com') ||
    url.includes('googleusercontent.com') ||
    url.includes('parrotsoftware.io')
  ) {
    return; // dejar que el navegador lo maneje normal (red)
  }

  // 2. Solo manejar GET
  if (e.request.method !== 'GET') return;

  // 3. App shell y assets: network-first con fallback a cache (funciona offline)
  e.respondWith(
    fetch(e.request)
      .then((resp) => {
        // Guardar copia fresca en cache
        const copy = resp.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
        return resp;
      })
      .catch(() => {
        // Sin red: servir desde cache
        return caches.match(e.request).then((cached) => {
          if (cached) return cached;
          // Si es una navegación, devolver el index cacheado
          if (e.request.mode === 'navigate') return caches.match('/index.html');
          return new Response('Sin conexión', { status: 503, statusText: 'Offline' });
        });
      })
  );
});
