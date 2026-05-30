// Service Worker - SuperOfertas Cartazes PWA
const CACHE_NAME = 'cartazes-v1';

// Arquivos para cache (interface básica)
const STATIC_ASSETS = [
  '/',
  '/static/style.css',
  '/static/manifest.json',
  '/static/icons/icon-192x192.png',
  '/static/icons/icon-512x512.png',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;800&display=swap'
];

// Instala o SW e faz cache dos assets estáticos
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[SW] Cache aberto');
      // Tenta cachear cada asset individualmente para não falhar tudo
      return Promise.allSettled(
        STATIC_ASSETS.map(url => cache.add(url).catch(err => console.warn('[SW] Falhou ao cachear:', url, err)))
      );
    })
  );
  self.skipWaiting();
});

// Ativa e limpa caches antigos
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames
          .filter(name => name !== CACHE_NAME)
          .map(name => {
            console.log('[SW] Removendo cache antigo:', name);
            return caches.delete(name);
          })
      );
    })
  );
  self.clients.claim();
});

// Estratégia: Network First com fallback para cache
// Garante que os dados (ofertas) sejam sempre atuais da rede
self.addEventListener('fetch', (event) => {
  const { request } = event;
  const url = new URL(request.url);

  // Ignora requisições não-GET e cross-origin (exceto fontes)
  if (request.method !== 'GET') return;
  if (!url.origin.startsWith('http')) return;

  // Para assets estáticos: Cache First (mais rápido)
  if (url.pathname.startsWith('/static/')) {
    event.respondWith(
      caches.match(request).then((cached) => {
        return cached || fetch(request).then((response) => {
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(request, clone));
          }
          return response;
        });
      })
    );
    return;
  }

  // Para páginas HTML: Network First (dados sempre atuais)
  event.respondWith(
    fetch(request)
      .then((response) => {
        // Cacheia a resposta bem-sucedida
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(request, clone));
        }
        return response;
      })
      .catch(() => {
        // Se offline, tenta o cache
        return caches.match(request).then((cached) => {
          if (cached) return cached;
          // Fallback para a página principal
          return caches.match('/');
        });
      })
  );
});

// Notificação de atualização disponível
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
