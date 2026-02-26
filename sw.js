const CACHE_NAME = 'attendance-v109';
const urlsToCache = [
  './index.html',
  './manifest.json',
  './icons/icon-72x72.png',
  './icons/icon-96x96.png',
  './icons/icon-128x128.png',
  './icons/icon-144x144.png',
  './icons/icon-152x152.png',
  './icons/icon-184x184.png',
  './icons/icon-192x192.png',
  './icons/icon-384x384.png',
  './icons/icon-512x512.png',
  './logo.png',
  'https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css'
];

// Install event - cache essential files
self.addEventListener('install', (event) => {
  console.log('[SW] Installing Service Worker v104...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('[SW] Opened cache');
        return cache.addAll(urlsToCache);
      })
      .catch((error) => {
        console.error('[SW] Cache addAll failed:', error);
      })
  );
  // Force immediate activation
  self.skipWaiting();
});

// Fetch event - serve from cache, fallback to network
self.addEventListener('fetch', (event) => {
  const requestUrl = new URL(event.request.url);
  
  // Skip non-http requests
  if (!requestUrl.protocol.startsWith('http')) {
    return;
  }
  
  // Skip internal Replit routes
  if (requestUrl.origin === location.origin && requestUrl.pathname.startsWith('/_')) {
    return;
  }
  
  // Skip API requests - always go to network
  if (requestUrl.pathname.includes('/api/') || requestUrl.hostname.includes('logic.azure.com')) {
    return;
  }
  
  const isHTMLRequest = event.request.mode === 'navigate' || 
    requestUrl.pathname.endsWith('.html') || 
    requestUrl.pathname === '/' || 
    requestUrl.pathname === '';

  if (isHTMLRequest) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          if (response && response.status === 200) {
            const responseToCache = response.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, responseToCache).catch(() => {});
            });
          }
          return response;
        })
        .catch(() => {
          return caches.match(event.request).then(r => r || caches.match('./index.html'));
        })
    );
    return;
  }

  event.respondWith(
    caches.match(event.request)
      .then((response) => {
        if (response) {
          return response;
        }
        
        const fetchRequest = event.request.clone();
        
        return fetch(fetchRequest).then((response) => {
          if (!response || response.status !== 200 || response.type === 'error') {
            return response;
          }
          
          if (!event.request.url.startsWith('http')) {
            return response;
          }
          
          if (event.request.method !== 'GET') {
            return response;
          }
          
          const responseToCache = response.clone();
          
          caches.open(CACHE_NAME)
            .then((cache) => {
              cache.put(event.request, responseToCache)
                .catch((error) => {
                  console.warn('[SW] Failed to cache:', event.request.url, error);
                });
            });
          
          return response;
        }).catch((error) => {
          console.error('[SW] Fetch failed:', error);
          if (event.request.mode === 'navigate') {
            return caches.match('./index.html');
          }
          return new Response('Offline', { status: 503 });
        });
      })
  );
});

// Activate event - clean old caches
self.addEventListener('activate', (event) => {
  console.log('[SW] Activating Service Worker v104...');
  const cacheWhitelist = [CACHE_NAME];
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheWhitelist.indexOf(cacheName) === -1) {
            console.log('[SW] Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  // Take control of all pages immediately
  self.clients.claim();
});

// Message handler for skip waiting
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

console.log('[SW] Service Worker v104 loaded');
