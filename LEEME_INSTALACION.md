# ClimateGuard Mobile — Instalación como app (PWA)

Esta carpeta contiene la app lista para instalarse en el teléfono como una
aplicación (pantalla completa, icono propio, funciona offline). No requiere
tienda de aplicaciones.

## Contenido de la carpeta

- `index.html` — la app (todo embebido: mapa, fuentes, 3.258 especies IUCN)
- `manifest.json` — define nombre, icono y modo pantalla completa
- `sw.js` — service worker (cachea la app para uso offline)
- `icon-192.png`, `icon-512.png` — iconos de la app

**Importante:** los 5 archivos deben estar juntos, en la misma carpeta, y
servirse por una dirección **https://** (o `localhost` para pruebas). Abrir
`index.html` directamente desde el disco (file://) NO activa el GPS ni la
instalación — es una regla de seguridad de los navegadores.

## Cómo publicarlo en HTTPS (elige una opción)

### Opción A — Netlify Drop (más rápido, sin cuenta técnica)
1. Entra a https://app.netlify.com/drop
2. Arrastra esta carpeta completa a la página.
3. Obtienes una dirección https://… en segundos. Listo.

### Opción B — GitHub Pages (gratis, sin instalar git)
1. En github.com crea un repositorio **público** (los privados requieren plan
   de pago para Pages). Ej.: `climateguard`.
2. En el repo: «Add file» → «Upload files». Sube los 7 archivos de esta
   carpeta (incluye `.nojekyll`), no la carpeta en sí. Confirma («Commit»).
3. Settings → Pages → Source: «Deploy from a branch» → Branch: `main`,
   carpeta `/ (root)` → Save.
4. Tras ~1 minuto la app queda en `https://TU-USUARIO.github.io/climateguard/`.
   Esa dirección https:// ya sirve para instalar la app en el teléfono.

Nota: la app usa rutas relativas, así que funciona en la subruta
`/climateguard/` sin cambios. El archivo `.nojekyll` evita el procesado
Jekyll y acelera el despliegue.

### Opción C — Cloudflare Pages / servidor propio de MMARN
Sube estos archivos a un proyecto de Cloudflare Pages, o cópialos a una
carpeta servida por HTTPS en el servidor de MMARN. No hace falta backend.

## Cómo instalarla en el teléfono

Abre la dirección https://… en el navegador del teléfono y:

- **Android (Chrome):** menú ⋮ → «Instalar aplicación» / «Añadir a pantalla
  de inicio». Aparece el icono de ClimateGuard.
- **iPhone/iPad (Safari):** botón Compartir → «Añadir a pantalla de inicio».

Al abrirla desde el icono se ejecuta a pantalla completa, sin barra del
navegador, y con GPS y cámara operativos.

## Permisos

La primera vez pedirá **ubicación** (para las coordenadas) y **cámara**
(para las fotos). Si se rechaza la ubicación, dentro de la app toca el
indicador de ubicación (arriba a la derecha) para ver los pasos de
reactivación según el dispositivo.

## Actualizar la app

Al publicar una versión nueva, sube los archivos y sube el número de versión
en dos sitios para forzar la actualización: `CACHE = 'climateguard-vXX'` en
`sw.js`. Los usuarios recibirán la nueva versión al reabrir la app.
