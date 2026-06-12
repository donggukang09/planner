// 업무 플래너 Service Worker — v3
// HTML은 네트워크 우선(최신 보장), 정적 자원은 캐시 우선
const CACHE = 'planner-v3';
const ASSETS = [
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './favicon.png'
];

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS).catch(()=>{})));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // 구글/외부 API 요청은 항상 네트워크
  if (url.origin !== self.location.origin) return;

  const isHTML = req.mode === 'navigate' || req.headers.get('accept')?.includes('text/html');
  if (isHTML) {
    // 네트워크 우선, 실패 시 캐시
    e.respondWith(
      fetch(req).then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy));
        return res;
      }).catch(() => caches.match(req).then(r => r || caches.match('./index.html')))
    );
  } else {
    // 캐시 우선
    e.respondWith(caches.match(req).then(r => r || fetch(req)));
  }
});

// 알림 클릭 시 앱 포커스
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(clients.matchAll({type:'window'}).then(list => {
    for (const c of list){ if ('focus' in c) return c.focus(); }
    if (clients.openWindow) return clients.openWindow('./index.html');
  }));
});
