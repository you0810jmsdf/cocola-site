/* COCoLa DAO PWA Service Worker
   - アプリシェルのオフラインキャッシュ
   - 通知クリック時にアプリを前面化
   実証段階のため Web Push（VAPID）は未使用。通知はページ側のポーリングから出す。 */

var CACHE_NAME = 'cocola-dao-v10';
var CORE_ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  '../assets/icons/icon-192.png',
  '../assets/icons/icon-512.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(CORE_ASSETS).catch(function () { /* 個別失敗は無視 */ });
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (key) {
        if (key !== CACHE_NAME) return caches.delete(key);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

/* data.json は常に最新を取りに行く（ネットワーク優先）。
   それ以外のアプリシェルはネットワーク優先＋キャッシュフォールバック。 */
self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;

  var url = new URL(req.url);
  var isData = url.pathname.indexOf('/data.json') !== -1;

  if (isData) {
    event.respondWith(fetch(req).catch(function () { return caches.match(req); }));
    return;
  }

  event.respondWith(
    fetch(req).then(function (res) {
      if (res && res.status === 200 && url.origin === self.location.origin) {
        var copy = res.clone();
        caches.open(CACHE_NAME).then(function (cache) { cache.put(req, copy); });
      }
      return res;
    }).catch(function () {
      return caches.match(req).then(function (hit) {
        return hit || caches.match('./index.html');
      });
    })
  );
});

/* バックグラウンド通知：サーバー(Web Push)からの push を受信して通知を表示する。
   アプリ／ブラウザが閉じていてもOSが起動して通知を出せる。 */
self.addEventListener('push', function (event) {
  var payload = { title: 'COCoLa DAO', body: '新しいお知らせがあります', url: './index.html', tag: 'cocola-push' };
  if (event.data) {
    try {
      var json = event.data.json();
      if (json && typeof json === 'object') {
        if (json.title) payload.title = json.title;
        if (json.body) payload.body = json.body;
        if (json.url) payload.url = json.url;
        if (json.tag) payload.tag = json.tag;
      }
    } catch (e) {
      var text = event.data.text();
      if (text) payload.body = text;
    }
  }
  event.waitUntil(
    self.registration.showNotification(payload.title, {
      body: payload.body,
      icon: '../assets/icons/icon-192.png',
      badge: '../assets/icons/icon-192.png',
      tag: payload.tag,
      renotify: true,
      data: { url: payload.url }
    })
  );
});

/* 購読がブラウザ側で更新された場合に再購読する（endpoint失効対策の足がかり）。
   再購読後のサーバー再登録はページ側で行うため、ここでは最小限。 */
self.addEventListener('pushsubscriptionchange', function (event) {
  // 詳細な再購読処理はページ側（VAPID公開鍵を持つ）に委ねる。
  // ここではログ目的のみ（将来 Background Sync 連携の拡張点）。
});

/* ページから postMessage('SHOW_NOTIFICATION', payload) を受けて通知を出す
   （PWAがバックグラウンドでもSW経由なら表示できる） */
self.addEventListener('message', function (event) {
  var data = event.data || {};
  if (data.type === 'SHOW_NOTIFICATION' && self.registration.showNotification) {
    var n = data.payload || {};
    self.registration.showNotification(n.title || 'COCoLa DAO', {
      body: n.body || '',
      icon: '../assets/icons/icon-192.png',
      badge: '../assets/icons/icon-192.png',
      tag: n.tag || 'cocola-dao',
      renotify: true,
      data: { url: n.url || './index.html' }
    });
  }
});

self.addEventListener('notificationclick', function (event) {
  event.notification.close();
  var target = (event.notification.data && event.notification.data.url) || './index.html';
  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (list) {
      for (var i = 0; i < list.length; i++) {
        if (list[i].url.indexOf('/dao/') !== -1 && 'focus' in list[i]) return list[i].focus();
      }
      if (self.clients.openWindow) return self.clients.openWindow(target);
    })
  );
});
