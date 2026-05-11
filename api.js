/**
 * MHS Integradora — api.js
 * Reemplaza google.script.run con llamadas fetch() al GAS Web App.
 *
 * ══════════════════════════════════════════════════════════════
 *  ⚙️  PASO OBLIGATORIO:
 *     Cambia MHS_GAS_URL por la URL de tu Web App de GAS.
 *
 *  Cómo obtenerla:
 *    1. Google Apps Script → Implementar → Administrar implementaciones
 *    2. Nueva implementación → Web App
 *       · Ejecutar como: Yo
 *       · Acceso: Cualquier usuario
 *    3. Copia la URL que termina en /exec
 *    4. Pégala abajo (reemplaza TODO el string, incluyendo las comillas)
 * ══════════════════════════════════════════════════════════════
 */

window.MHS_GAS_URL = 'https://script.google.com/macros/s/AKfycbzzEv3X_MJ3ZnilwzyjSyHe5GQ2AoaNIC5iBp7dv42T-Suloc_lKbrM4ihrlOQvKmWx/exec';

// ── Timeout por petición (ms) ─────────────────────────────────────────────
var MHS_API_TIMEOUT_MS = 30000;   // 30 segundos

// ── Detectar URL sin configurar y mostrar aviso visible ───────────────────
(function checkURL() {
  if (window.MHS_GAS_URL.indexOf('https://script.google.com/macros/s/AKfycbzzEv3X_MJ3ZnilwzyjSyHe5GQ2AoaNIC5iBp7dv42T-Suloc_lKbrM4ihrlOQvKmWx/exec') < 0) return;

  // Mostrar banner de error en cuanto cargue el DOM
  function showConfigError() {
    // Ocultar loader si está visible
    var loader = document.getElementById('loader');
    if (loader) loader.classList.remove('show');

    // Banner de aviso
    var banner = document.createElement('div');
    banner.id = 'mhs-config-error';
    banner.style.cssText = [
      'position:fixed','top:0','left:0','right:0','z-index:99999',
      'background:#c0392b','color:#fff','font-family:sans-serif',
      'font-size:14px','padding:16px 20px','line-height:1.5',
      'box-shadow:0 2px 8px rgba(0,0,0,.4)'
    ].join(';');
    banner.innerHTML =
      '<strong>⚠️ MHS — Falta configurar la URL del servidor</strong><br>' +
      'Edita <code>api.js</code> y reemplaza <code>TU_DEPLOYMENT_ID</code> ' +
      'con la URL de tu Web App de Google Apps Script.<br>' +
      '<small>Implementar → Administrar implementaciones → copiar URL /exec</small>';
    document.body.insertBefore(banner, document.body.firstChild);

    // Mostrar pantalla de login (si existe) en vez del loader
    var loginScreen = document.getElementById('screenLogin');
    if (loginScreen) {
      var screens = document.querySelectorAll('[id^="screen"]');
      screens.forEach(function(s) { s.classList.remove('active'); });
      loginScreen.classList.add('active');
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', showConfigError);
  } else {
    showConfigError();
  }
})();

// ── Shim de google.script.run ─────────────────────────────────────────────
(function (w) {
  'use strict';

  // Si ya existe el objeto nativo de GAS, no sobreescribimos nada.
  if (
    typeof w.google !== 'undefined' &&
    w.google.script &&
    typeof w.google.script.run === 'object' &&
    w.google.script.run !== null
  ) return;

  /**
   * Builder — replica el patrón de GAS:
   *   google.script.run
   *     .withSuccessHandler(fn)
   *     .withFailureHandler(fn)
   *     .apiMiMetodo(payload)
   */
  function Builder(successFn, failureFn) {
    var obj = {
      withSuccessHandler: function (fn) { return Builder(fn, failureFn); },
      withFailureHandler: function (fn) { return Builder(successFn, fn); }
    };

    return new Proxy(obj, {
      get: function (target, prop) {
        if (Object.prototype.hasOwnProperty.call(target, prop)) return target[prop];
        if (prop === 'then' || prop === 'catch' || prop === 'finally') return undefined;

        return function (payload) {
          var body = JSON.stringify(
            Object.assign({ action: prop }, payload || {})
          );

          // Abort controller para timeout
          var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
          var timeoutId  = controller
            ? setTimeout(function() { controller.abort(); }, MHS_API_TIMEOUT_MS)
            : null;

          fetch(w.MHS_GAS_URL, {
            method:  'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body:    body,
            signal:  controller ? controller.signal : undefined
          })
          .then(function (r) {
            if (timeoutId) clearTimeout(timeoutId);
            if (!r.ok) throw new Error('HTTP ' + r.status + ' al llamar ' + prop);
            return r.json();
          })
          .then(function (data) {
            if (successFn) successFn(data);
          })
          .catch(function (err) {
            if (timeoutId) clearTimeout(timeoutId);
            var msg = err && err.name === 'AbortError'
              ? 'Tiempo de espera agotado al llamar ' + prop + ' (>' + (MHS_API_TIMEOUT_MS/1000) + 's)'
              : (err && err.message) || String(err);
            console.warn('[MHS api.js] Error en ' + prop + ':', msg);
            if (failureFn) failureFn({ message: msg });
            // Si no hay failureHandler, ocultar loader para no quedar pegado
            else {
              var loader = document.getElementById('loader');
              if (loader) loader.classList.remove('show');
              console.error('[MHS] Sin failureHandler para ' + prop + '. Error: ' + msg);
            }
          });
        };
      }
    });
  }

  w.google        = w.google || {};
  w.google.script = { run: Builder(null, null) };

  console.info('[MHS api.js] Shim activo → ' + w.MHS_GAS_URL);

})(window);
