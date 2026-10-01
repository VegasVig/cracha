/* Service worker — Vegas · Entrega de Crachá
   Guarda a tela do app para abrir rápido. As chamadas à API (Apps Script)
   nunca passam pelo cache: sempre vão direto ao servidor. */
const CACHE = "vegas-cracha-v1";
const ARQUIVOS = ["./", "./index.html", "./manifest.json",
  "./icons/icon-96.png", "./icons/icon-192.png", "./icons/icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET" || new URL(req.url).origin !== location.origin) return;

  // Páginas: tenta a rede primeiro (versão sempre atualizada) e usa o cache se estiver sem internet
  if (req.mode === "navigate") {
    e.respondWith(fetch(req).then(r => {
      caches.open(CACHE).then(c => c.put("./index.html", r.clone()));
      return r;
    }).catch(() => caches.match("./index.html")));
    return;
  }
  // Ícones e manifesto: cache primeiro
  e.respondWith(caches.match(req).then(c => c || fetch(req)));
});
