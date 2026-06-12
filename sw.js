// 업무 플래너 Service Worker — v4
const CACHE = 'planner-v4';
const ASSETS = ['./index.html','./manifest.json','./icon-192.png','./icon-512.png','./favicon.png'];

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
  if (req.method !== 'GET') return;                       // 쓰기(POST)는 건드리지 않음
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;        // 구글 등 외부 요청은 그대로 통과

  const isHTML = req.mode === 'navigate' || (req.headers.get('accept')||'').includes('text/html');
  if (isHTML) {
    e.respondWith(
      fetch(req)
        .then(res => { const c = res.clone(); caches.open(CACHE).then(ch => ch.put(req, c)); return res; })
        .catch(() => caches.match(req).then(r => r || caches.match('./index.html')).then(r => r || new Response('오프라인', {status:503, headers:{'Content-Type':'text/plain; charset=utf-8'}})))
    );
  } else {
    e.respondWith(
      caches.match(req).then(r => r || fetch(req).catch(() => new Response('', {status:503})))
    );
  }
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(clients.matchAll({type:'window'}).then(list => {
    for (const c of list){ if ('focus' in c) return c.focus(); }
    if (clients.openWindow) return clients.openWindow('./index.html');
  }));
});
