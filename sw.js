// Service worker de la app "Ruteo Copsa".
// Siempre intenta traer la versión más nueva de internet (así los cambios
// que se suben a GitHub llegan solos); si no hay señal, muestra la última
// copia guardada de la pantalla. No toca los datos de Firebase.
const CACHE = 'ruteo-copsa-v1';

self.addEventListener('install', () => self.skipWaiting());

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  event.respondWith(
    fetch(req)
      .then(res => {
        const copia = res.clone();
        caches.open(CACHE).then(c => c.put(req, copia));
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }))
  );
});
