const CACHE_NAME = "workshop-cache-v1";
const urlsToCache = ["/", "/index.html", "/app.js"];

self.addEventListener("install", e=>{
  e.waitUntil(caches.open(CACHE_NAME).then(c=>c.addAll(urlsToCache)));
});

self.addEventListener("fetch", e=>{
  e.respondWith(
    caches.match(e.request).then(r => r || fetch(e.request).catch(()=>r))
  );
});
