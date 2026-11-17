const CACHE = "weather-station-v2";
const ASSETS = [
  "./",
  "./index.html",
  "./app.js",
  "./xlsx.full.min.js",
  "./data_processing.py",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./fonts/msyh.ttc",
];

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE).then(cache => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", event => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);
  if (url.origin !== location.origin) return;  // Pyodide CDN 交给浏览器缓存
  event.respondWith(
    caches.match(event.request).then(hit => hit || fetch(event.request))
  );
});
