/* redeploy nonce: v1.9.0 envio confirmado, foto completa JPEG, datos por usuario */
/* ClimateGuard Mobile - Service Worker
   Cachea el shell de la app para uso offline. Casi todo (fuentes, iconos,
   datos IUCN, contorno nacional y areas protegidas WDPA) esta embebido en
   index.html. Leaflet se carga desde CDN, por eso lo pre-cacheamos aqui
   para que el mapa vectorial (contorno RD + WDPA) funcione sin conexion.
   Los mosaicos de mapa (OSM/Esri) son de otro origen: se intentan por red
   y, si ya se visitaron, salen de cache.

   Estrategia del shell (index.html): RED PRIMERO. Al abrir el icono
   instalado, si hay conexion se descarga la ultima version y se guarda en
   cache; sin conexion se sirve la copia cacheada. Combinado con
   caches.delete de versiones viejas en 'activate' y reg.update() en la
   app, el usuario siempre recibe la version mas reciente. */
const CACHE = 'climateguard-v1.9.0';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];
/* Recursos de terceros necesarios para el mapa offline (best-effort) */
const EXTRA = [
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(ASSETS)
        .then(() => Promise.allSettled(EXTRA.map(u => c.add(u).catch(() => {})))))
      .then(() => self.skipWaiting())
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
      // v1.9: con senal debil la red puede tardar mucho; si en 4 s no responde
      // se abre la copia guardada y la descarga sigue en segundo plano.
      const red = fetch(req).then(res => {
        if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
        return res;
      });
      const guardada = () => caches.match(req).then(hit => hit || caches.match('./index.html'));
      e.respondWith(new Promise(resolve => {
        let listo = false;
        const t = setTimeout(() => { guardada().then(h => { if (h && !listo) { listo = true; resolve(h); } }); }, 4000);
        red.then(res => { if (!listo) { listo = true; clearTimeout(t); resolve(res); } })
           .catch(() => { clearTimeout(t); guardada().then(h => { if (!listo) { listo = true; resolve(h || Response.error()); } }); });
      }));
      e.waitUntil(red.catch(() => {}));
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
    // Otro origen (Leaflet CDN, mosaicos de mapa): red primero, cache como respaldo.
    e.respondWith(
      fetch(req).then(res => {
        if (res && res.status === 200 && (res.type === 'basic' || res.type === 'cors')) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
        }
        return res;
      }).catch(() => caches.match(req))
    );
  }
});
