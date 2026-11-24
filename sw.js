const CACHE = "at-passport-v2";
const ASSETS = ["./", "./index.html", "./manifest.json"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// Network-first for navigation/HTML so a new deploy is always picked up on
// the next load; fall back to cache only when offline. Other same-origin
// GET requests (fonts, vendor libs, images) stay cache-first since they are
// hashed-by-filename-unlikely-to-change assets and this keeps the app usable
// offline.
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  const url = new URL(e.request.url);
  if (url.origin !== self.location.origin) return;

  const isNavigation =
    e.request.mode === "navigate" ||
    (e.request.destination === "" && e.request.headers.get("accept")?.includes("text/html")) ||
    url.pathname.endsWith("index.html") ||
    url.pathname.endsWith("/");

  if (isNavigation) {
    e.respondWith(
      fetch(e.request)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy));
          return res;
        })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  e.respondWith(
    caches.match(e.request).then(
      (cached) =>
        cached ||
        fetch(e.request)
          .then((res) => {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(e.request, copy));
            return res;
          })
          .catch(() => cached)
    )
  );
});
