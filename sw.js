// Salva l'app sul dispositivo: si apre anche con connessione debole. I dati restano gestiti da Firebase.
const V = 'conti-v1';
const SHELL = ['./', 'index.html', 'css/style.css', 'js/app.js', 'js/excel.js', 'js/backend-local.js', 'js/backend-firebase.js', 'js/firebase-config.js', 'manifest.webmanifest'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== V).map(x => caches.delete(x)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  // prima la rete (così gli aggiornamenti arrivano subito), la copia salvata se manca la connessione
  e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(V).then(ch => ch.put(e.request, c)); return r; }).catch(() => caches.match(e.request).then(m => m || caches.match('index.html'))));
});
