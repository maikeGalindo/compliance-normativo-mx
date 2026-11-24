// Service worker de limpieza: la app ya no usa PWA/offline.
// Borra cachés viejos y se desregistra para que nadie siga viendo la versión rota.
self.addEventListener('install',function(){self.skipWaiting()});
self.addEventListener('activate',function(e){e.waitUntil(
  caches.keys().then(function(ks){return Promise.all(ks.map(function(k){return caches.delete(k)}))})
  .then(function(){return self.registration.unregister()})
  .then(function(){return self.clients.matchAll()})
  .then(function(cs){cs.forEach(function(c){c.navigate(c.url)})}))});
