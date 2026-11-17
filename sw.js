const CACHE = "weather-station-v4";
const ASSETS = [
  "./",
  "./index.html",
  "./app.js",
  "./xlsx.full.min.js",
  "./data_processing.py",
  "./data_processing_weekly.py",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./fonts/msyh-subset.ttf",
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
