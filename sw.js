// Guarda la "cáscara" de la app para que abra rápido y sin cobertura. Los datos siempre se piden al servidor.
const V='cnsalou-v1';
const SHELL=['./','index.html','config.js','manifest.webmanifest','logo.png','icon-192.png','icon-512.png','apple-touch-icon.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(SHELL)));self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==V).map(k=>caches.delete(k)))));self.clients.claim()});
self.addEventListener('fetch',e=>{
  const u=new URL(e.request.url);
  if(e.request.method!=='GET'||u.origin!==location.origin)return; // API y fuentes: directo a la red
  // red primero (siempre la última versión); si no hay cobertura, la copia guardada
  e.respondWith(fetch(e.request).then(r=>{const c=r.clone();caches.open(V).then(x=>x.put(e.request,c));return r}).catch(()=>caches.match(e.request).then(r=>r||caches.match('index.html'))));
});
