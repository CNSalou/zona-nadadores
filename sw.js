// Guarda la "cáscara" de la app para que abra sin cobertura. Siempre intenta primero la versión más nueva del servidor.
const V='cnsalou-v2';
const SHELL=['./','index.html','config.js','manifest.webmanifest','logo.png','icon-192.png','icon-512.png','apple-touch-icon.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(SHELL.map(u=>new Request(u,{cache:'reload'})))));self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==V).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const u=new URL(e.request.url);
  if(e.request.method!=='GET'||u.origin!==location.origin)return; // API y fuentes: directo a la red
  // red primero, sin usar la copia del navegador (no-cache = comprobar siempre si hay versión nueva); sin cobertura, la copia guardada
  e.respondWith(fetch(e.request.url,{cache:'no-cache',credentials:'same-origin'}).then(r=>{if(r.ok){const c=r.clone();caches.open(V).then(x=>x.put(e.request,c))}return r})
    .catch(()=>caches.match(e.request).then(r=>r||caches.match('index.html'))));
});
