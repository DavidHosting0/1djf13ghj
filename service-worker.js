/**
 * Service Worker für Offline-Funktionalität
 * Cached alle notwendigen Ressourcen für Offline-Nutzung
 */

const CACHE_NAME = 'bernticker-v1';
const STATIC_CACHE_URLS = [
    './',
    './index.html',
    './styles.css',
    './app.js',
    './excelParser.js',
    './csvExporter.js',
    './manifest.json',
    // SheetJS CDN wird nicht gecacht, da es extern ist
    // Bei vollständiger Offline-Nutzung sollte die Library lokal eingebunden werden
];

// Install Event - Cache alle statischen Ressourcen
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => {
                console.log('Service Worker: Caching statische Ressourcen');
                // Nur erfolgreiche Requests cachen
                return Promise.allSettled(
                    STATIC_CACHE_URLS.map((url) => {
                        return fetch(url)
                            .then((response) => {
                                if (response.ok) {
                                    return cache.put(url, response);
                                }
                            })
                            .catch(() => {
                                // Fehler beim Cachen ignorieren
                                console.warn(`Service Worker: Konnte ${url} nicht cachen`);
                            });
                    })
                );
            })
            .then(() => {
                // Service Worker sofort aktivieren
                return self.skipWaiting();
            })
    );
});

// Activate Event - Alte Caches löschen
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys()
            .then((cacheNames) => {
                return Promise.all(
                    cacheNames
                        .filter((cacheName) => cacheName !== CACHE_NAME)
                        .map((cacheName) => {
                            console.log('Service Worker: Lösche alten Cache:', cacheName);
                            return caches.delete(cacheName);
                        })
                );
            })
            .then(() => {
                // Service Worker sofort kontrollieren
                return self.clients.claim();
            })
    );
});

// Fetch Event - Cache-First Strategie für statische Ressourcen
self.addEventListener('fetch', (event) => {
    const { request } = event;
    const url = new URL(request.url);

    // Nur GET-Requests behandeln
    if (request.method !== 'GET') {
        return;
    }

    // Externe Ressourcen (CDN) - Cache-First für Offline-Funktionalität
    if (url.origin !== location.origin) {
        // Für CDN-Ressourcen: Cache-First, dann Network
        event.respondWith(
            caches.match(request)
                .then((cachedResponse) => {
                    if (cachedResponse) {
                        return cachedResponse;
                    }
                    // Wenn nicht im Cache, vom Netzwerk holen und cachen
                    return fetch(request)
                        .then((response) => {
                            // Nur erfolgreiche Responses cachen
                            if (response && response.status === 200) {
                                const responseToCache = response.clone();
                                caches.open(CACHE_NAME)
                                    .then((cache) => {
                                        cache.put(request, responseToCache);
                                    });
                            }
                            return response;
                        });
                })
        );
        return;
    }

    // Für lokale Ressourcen: Cache-First Strategie
    event.respondWith(
        caches.match(request)
            .then((cachedResponse) => {
                if (cachedResponse) {
                    return cachedResponse;
                }

                // Wenn nicht im Cache, vom Netzwerk holen und cachen
                return fetch(request)
                    .then((response) => {
                        // Nur erfolgreiche Responses cachen
                        if (response && response.status === 200 && response.type === 'basic') {
                            const responseToCache = response.clone();
                            caches.open(CACHE_NAME)
                                .then((cache) => {
                                    cache.put(request, responseToCache);
                                });
                        }
                        return response;
                    })
                    .catch(() => {
                        // Wenn Netzwerk fehlschlägt und kein Cache vorhanden,
                        // könnte hier eine Offline-Seite zurückgegeben werden
                        return new Response('Offline - Bitte stellen Sie eine Internetverbindung her.', {
                            status: 503,
                            statusText: 'Service Unavailable',
                            headers: new Headers({
                                'Content-Type': 'text/plain'
                            })
                        });
                    });
            })
    );
});

