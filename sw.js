/* Service worker: cache toàn bộ app để chạy offline + nhận file chia sẻ (Android Web Share Target). */
const VERSION = 'kk-v1.7.1';
const MASTER_URL = './data/MASTERR.xlsx';
const ASSETS = [
  './', './index.html', './styles.css', './app.js', './parsers.js', './db.js', MASTER_URL,
  './report.js', './vendor/xlsx.full.min.js', './vendor/jspdf.umd.min.js', './vendor/jspdf.plugin.autotable.min.js', './manifest.webmanifest',
  './fonts/Tinos-Regular.ttf', './fonts/Tinos-Bold.ttf', './fonts/Tinos-Italic.ttf',
  './fonts/Montserrat-Regular.ttf', './fonts/Montserrat-SemiBold.ttf',
  './report/logo-cp.png', './report/logo-cp-circle.jpg', './report/cover-store.jpg',
  './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png', './icons/apple-touch-icon.png'
];
const SHARE_CACHE = 'kk-shared';

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== VERSION && k !== SHARE_CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // File được chia sẻ từ app khác (OneDrive...) vào app này
  if (e.request.method === 'POST' && url.pathname.endsWith('/share-target')) {
    e.respondWith((async () => {
      try {
        const form = await e.request.formData();
        const file = form.getAll('file').find(f => f && typeof f === 'object');
        if (file) {
          const c = await caches.open(SHARE_CACHE);
          await c.put('./shared-file', new Response(file, { headers: { 'x-file-name': encodeURIComponent(file.name || 'shared.xlsx') } }));
        }
      } catch (err) { /* bỏ qua, app sẽ báo không có file */ }
      return Response.redirect('./?share=1', 303);
    })());
    return;
  }

  if (e.request.method !== 'GET' || url.origin !== location.origin) return;

  // File master: có mạng thì lấy bản mới nhất (và cập nhật cache), mất mạng thì dùng bản đã cache
  if (url.pathname.endsWith('/data/MASTERR.xlsx')) {
    e.respondWith(fetch(e.request, { cache: 'no-cache' })
      .then(r => { if (r.ok) { const copy = r.clone(); caches.open(VERSION).then(c => c.put(MASTER_URL, copy)); } return r; })
      .catch(() => caches.match(MASTER_URL)));
    return;
  }

  // Trang: ưu tiên cache để mở được khi mất mạng
  if (e.request.mode === 'navigate') {
    e.respondWith(caches.match('./index.html').then(r => r || fetch(e.request)));
    return;
  }
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(r => r || fetch(e.request)));
});
