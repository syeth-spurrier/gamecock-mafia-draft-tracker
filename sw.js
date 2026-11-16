const CACHE_NAME = "gcm-tracker-v10.0.0";
const APP_SHELL_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./league-crest.png",
  "./apple-touch-icon.png",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL_ASSETS)).catch(() => {})
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// Network-First strategy for HTML/JSON/assets with instant Offline Cache fallback
self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(req)
      .then((networkRes) => {
        if (networkRes && networkRes.ok && !url.pathname.endsWith("version.json")) {
          const clone = networkRes.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, clone)).catch(() => {});
        }
        return networkRes;
      })
      .catch(() =>
        caches.match(req).then((cached) => cached || caches.match("./index.html"))
      )
  );
});

// Web Push & Notification Click Handlers (Wave 2 OneSignal & Native PWA Support)
self.addEventListener("push", (event) => {
  let payload = {
    title: "Gamecock Mafia Alert",
    body: "A new league update is available in the Draft & Pick Tracker.",
    url: "./"
  };
  try {
    if (event.data) {
      const parsed = event.data.json();
      payload.title = parsed.title || (parsed.headings && parsed.headings.en) || payload.title;
      payload.body = parsed.body || (parsed.contents && parsed.contents.en) || payload.body;
      payload.url = parsed.url || (parsed.data && parsed.data.url) || payload.url;
    }
  } catch (e) {}

  event.waitUntil(
    self.registration.showNotification(payload.title, {
      body: payload.body,
      icon: "./icon-192.png",
      badge: "./icon-192.png",
      data: { url: payload.url }
    })
  );
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetUrl = (event.notification.data && event.notification.data.url) || "./";
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((windowClients) => {
      for (const client of windowClients) {
        if (client.url && "focus" in client) {
          return client.focus();
        }
      }
      if (self.clients.openWindow) {
        return self.clients.openWindow(targetUrl);
      }
    })
  );
});

