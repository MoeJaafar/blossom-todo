// Offline support for the installed app: serve our own files from cache,
// refresh them in the background so updates arrive on the next launch.
const CACHE = "blossom-todo-v2";
const FILES = [
  "taskpane.html",
  "taskpane.css",
  "taskpane.js",
  "cloud.js",
  "firebase-config.js",
  "app.webmanifest",
  "../assets/fonts/PixelifySans.ttf",
  "../assets/img/blossom.svg",
  "../assets/img/blossom-lilac.svg",
  "../assets/img/tile-light.svg",
  "../assets/img/tile-dark.svg",
  "../assets/icons/app-192.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;

  e.respondWith(
    caches.open(CACHE).then(async (cache) => {
      const cached = await cache.match(e.request, { ignoreSearch: true });
      const fresh = fetch(e.request)
        .then((res) => {
          if (res.ok) cache.put(e.request, res.clone());
          return res;
        })
        .catch(() => cached);
      return cached || fresh;
    })
  );
});
