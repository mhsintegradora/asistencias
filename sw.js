/**
 * MHS Integradora — sw.js  (Service Worker)
 * ──────────────────────────────────────────────────────────────────────────
 * Habilita:
 *   - Instalación en pantalla de inicio (iOS 16.4+ / Android)
 *   - Caché offline de archivos estáticos
 *   - Push Notifications (requiere suscripción VAPID desde el servidor)
 *   - Vibración y beep al recibir push (vía Notification API)
 * ──────────────────────────────────────────────────────────────────────────
 */

var CACHE_NAME = 'mhs-rh-v1';

/** Archivos que se cachean al instalar el SW */
var STATIC_FILES = [
  './',
  './mobile.html',
  './ceomobile.html',
  './index.html',
  './router.html',
  './api.js',
  './turno-filter.js',
  './unlock-dates.js',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap'
];

/* ── INSTALL: cachear archivos estáticos ─── */
self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(STATIC_FILES.map(function (url) {
        /* Usar Request para ignorar errores en URLs externas */
        return new Request(url, { mode: 'no-cors' });
      })).catch(function (err) {
        console.warn('[SW] Error al cachear archivos estáticos:', err);
      });
    })
  );
  self.skipWaiting();
});

/* ── ACTIVATE: limpiar cachés viejos ─── */
self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== CACHE_NAME; })
            .map(function (k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

/* ── FETCH: servir desde caché, fallback a red ─── */
self.addEventListener('fetch', function (event) {
  var url = event.request.url;

  /* Siempre ir a la red para el GAS Web App (API dinámica) */
  if (url.includes('script.google.com') || url.includes('googleusercontent.com')) {
    return; // dejar pasar sin interceptar
  }

  event.respondWith(
    caches.match(event.request).then(function (cached) {
      if (cached) return cached;
      return fetch(event.request).then(function (response) {
        /* Cachear respuestas válidas de recursos estáticos */
        if (
          response &&
          response.status === 200 &&
          event.request.method === 'GET' &&
          !url.includes('chrome-extension')
        ) {
          var resClone = response.clone();
          caches.open(CACHE_NAME).then(function (cache) {
            try { cache.put(event.request, resClone); } catch (e) {}
          });
        }
        return response;
      }).catch(function () {
        /* Sin red y sin caché: mostrar página offline si existe */
        return caches.match('./mobile.html');
      });
    })
  );
});

/* ── PUSH: recibir notificaciones del servidor ─── */
self.addEventListener('push', function (event) {
  var data = {};
  try {
    data = event.data ? event.data.json() : {};
  } catch (e) {
    data = { title: 'MHS RH Hub', body: event.data ? event.data.text() : '' };
  }

  var title   = data.title   || 'MHS Integradora · RH HUB';
  var body    = data.body    || 'Tienes una notificación pendiente.';
  var icon    = data.icon    || './icons/icon-192.png';
  var badge   = data.badge   || './icons/badge-96.png';
  var tag     = data.tag     || 'mhs-notif';
  var urgent  = data.urgent  || false;
  var url     = data.url     || './mobile.html';

  var options = {
    body:               body,
    icon:               icon,
    badge:              badge,
    tag:                tag,
    renotify:           true,
    requireInteraction: urgent,   // urgente: no se cierra automáticamente
    vibrate:            urgent ? [200, 100, 200, 100, 200] : [100, 50, 100],
    data:               { url: url, tipo: data.tipo || '' },
    actions: urgent ? [
      { action: 'abrir',    title: '👁️ Abrir' },
      { action: 'cerrar',   title: 'Cerrar' }
    ] : [
      { action: 'abrir',    title: 'Ver' }
    ]
  };

  event.waitUntil(
    self.registration.showNotification(title, options)
  );
});

/* ── NOTIFICATIONCLICK: abrir app al tocar la notificación ─── */
self.addEventListener('notificationclick', function (event) {
  event.notification.close();

  if (event.action === 'cerrar') return;

  var targetUrl = (event.notification.data && event.notification.data.url)
    ? event.notification.data.url
    : './mobile.html';

  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true })
      .then(function (clientList) {
        /* Si la app ya está abierta, enfocarla */
        for (var i = 0; i < clientList.length; i++) {
          var client = clientList[i];
          if (client.url.includes('mobile.html') || client.url.includes('mhsintegradora')) {
            client.focus();
            client.postMessage({ type: 'NOTIFICATION_CLICK', data: event.notification.data });
            return;
          }
        }
        /* Si no está abierta, abrir nueva ventana */
        return self.clients.openWindow(targetUrl);
      })
  );
});

/* ── MENSAJE desde la app principal ─── */
self.addEventListener('message', function (event) {
  if (!event.data) return;

  switch (event.data.type) {
    case 'SKIP_WAITING':
      self.skipWaiting();
      break;
    case 'CACHE_CLEAR':
      caches.delete(CACHE_NAME);
      break;
    default:
      break;
  }
});

console.log('[SW] MHS RH Hub service worker activo. Cache:', CACHE_NAME);
