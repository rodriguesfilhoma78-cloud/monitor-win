// Service worker minimo - so existe para o Chrome/Edge aceitarem instalar
// o dashboard como app (PWA). Sem cache: sempre busca da rede, porque os
// dados do monitor sao ao vivo e uma versao em cache seria perigosa.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (event) => event.waitUntil(self.clients.claim()));
self.addEventListener("fetch", (event) => {
  event.respondWith(fetch(event.request));
});
