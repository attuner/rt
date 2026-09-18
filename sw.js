const CACHE_NAME = "pulse-radio-v1";
const ASSETS_TO_CACHE = [
  "./radio.html",
  "./manifest.json"
];

// Install: Cache essential shell files
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS_TO_CACHE))
  );
  self.skipWaiting();
});

// Activate: Cleanup old cache versions
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) return caches.delete(key);
        })
      )
    )
  );
  self.clients.claim();
});

// Fetch: Network-first for streaming and API polling; cache fallback for shell
self.addEventListener("fetch", (event) => {
  const url = event.request.url;

  // Never cache live Google Drive media streams or Apps Script API polls
  if (url.includes("script.google.com") || url.includes("drive.google.com")) {
    return;
  }

  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});