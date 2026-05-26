// Service Worker fuer Kuchenproduktion
// -----------------------------------
// Strategie: network-first fuer alle gleich-Origin Anfragen (also alle
// HTMLs unter /Wochenplan/). Damit ignoriert der Browser den HTTP-Cache
// (Cache-Control: max-age=600 von GitHub Pages) und holt die Seite immer
// frisch vom Server.
//
// Anfragen an api.github.com (also die ist_produktion-PUTs/GETs) werden
// NICHT abgefangen, weil das eine andere Origin ist.
//
// Effekt fuer den User: nach einem Save sieht ein Reload die neue Version
// SOFORT, ohne 10 Minuten Browser-Cache abwarten zu muessen, ohne Inkognito,
// ohne manuell Cache zu leeren.

self.addEventListener('install', function(event) {
  // Sofort aktiv werden, ohne Wartezeit auf Page-Reload
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  // Alle alten Caches loeschen + sofort die Kontrolle uebernehmen
  event.waitUntil(
    caches.keys().then(function(names) {
      return Promise.all(names.map(function(n) { return caches.delete(n); }));
    }).then(function() {
      return self.clients.claim();
    })
  );
});

self.addEventListener('fetch', function(event) {
  var url = new URL(event.request.url);
  // Nur same-origin Anfragen abfangen — api.github.com nicht beruehren.
  if (url.origin !== self.location.origin) return;

  // Network-first: erst Server fragen mit cache:'no-store'. Bei Netz-Ausfall
  // (z.B. Handy offline) fallback auf das, was der Browser noch hat.
  event.respondWith(
    fetch(event.request, { cache: 'no-store' }).catch(function() {
      return fetch(event.request);  // letzter Fallback ohne no-store
    })
  );
});
