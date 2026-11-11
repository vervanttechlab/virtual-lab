// Offline support: cache the app shell and day data; the Apps Script API is never cached.
const VERSION = "dal3-lab-v33";
const SHELL = [
  "./", "index.html", "manifest.webmanifest", "css/app.css",
  "js/app.js", "js/api.js", "js/config.js", "js/ui.js", "js/workbook.js", "js/grid.js", "js/stats.js",
  "js/wbcore.js", "js/check.js", "js/hf-extras.js", "js/export.js", "js/videos.js", "js/sqllab.js", "vendor/exceljs.min.js",
  "vendor/sql-wasm.js", "vendor/sql-wasm.wasm", "data/aling_nena.db",
  "vendor/hyperformula.full.min.js", "vendor/chart.umd.js",
  "data/index.json", "data/day02.json", "data/day03.json", "data/day04.json", "data/day05.json",
  "icons/icon-192.png", "icons/icon-512.png",
  "guide.html", "guide/img/install-android.png", "guide/img/install-iphone.png", "guide/img/signin.png",
  "guide/img/day.png", "guide/img/workbook.png", "guide/img/desktop.png",
];
self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;     // API calls go straight to the network
  // Network first for data and code (so updates arrive), cache as the offline fallback.
  // cache: "no-cache" makes the browser re-check each file with GitHub (a cheap 304 when unchanged), so an update can
  // never mix a new index.html with old modules from the browser's 10-minute HTTP cache.
  const fresh = e.request.mode === "navigate" ? fetch(e.request) : fetch(new Request(e.request, { cache: "no-cache" }));
  e.respondWith(fresh.then((res) => {
    if (res.ok) { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(e.request, copy)); }
    return res;
  }).catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match("index.html"))));
});
