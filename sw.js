/* ClimateGuard Mobile - Service Worker
   Cachea el shell de la app para uso offline. Todo (Leaflet, fuentes,
   iconos, datos IUCN) esta embebido en index.html, asi que basta con
   cachear los archivos locales. Los mosaicos de mapa (OSM/Esri) son de
   otro origen: se intentan por red y, si ya se visitaron, salen de cache.

   Estrategia del shell (index.html): RED PRIMERO. Al abrir el icono
   instalado, si hay conexion se descarga la ultima version y se guarda
   en cache; sin conexion se sirve la copia cacheada. Combinado con
   caches.delete de versiones viejas en 'activate' y reg.update() en la
   app, el usuario siempre recibe la version mas reciente. */
const CACHE = 'climateguard-v1.1.1';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  if (url.origin === location.origin) {
    // ¿Es el shell de la app (documento / index.html)?
    const esShell = req.mode === 'navigate'
      || url.pathname.endsWith('/')
      || url.pathname.endsWith('/index.html')
      || url.pathname === '/index.html';
    if (esShell) {
      // RED PRIMERO: siempre la ultima version cuando hay conexion.
      e.respondWith(
        fetch(req).then(res => {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
          return res;
        }).catch(() => caches.match(req).then(hit => hit || caches.match('./index.html')))
      );
    } else {
      // Otros recursos locales (iconos, manifest): cache primero.
      e.respondWith(
        caches.match(req).then(hit => hit || fetch(req).then(res => {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
          return res;
        }).catch(() => caches.match('./index.html')))
      );
    }
  } else {
    // Otro origen (mosaicos de mapa): red primero, cache como respaldo.
    e.respondWith(
      fetch(req).then(res => {
        if (res && res.status === 200 && res.type === 'basic') {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
        }
        return res;
      }).catch(() => caches.match(req))
    );
  }
});
