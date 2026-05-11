/**
 * MHS Integradora — api.js
 * Reemplaza google.script.run con llamadas fetch() al GAS Web App.
 * Coloca este archivo ANTES de cualquier otro <script> en mobile.html, index.html, etc.
 *
 * ⚙️  CONFIGURACIÓN: cambia MHS_GAS_URL a tu URL de despliegue de GAS.
 *     Ve a GAS → Implementar → Nueva implementación → Web App
 *     Ejecutar como: Yo  |  Acceso: Cualquier usuario
 */

// ─── 1. URL del Web App de GAS ──────────────────────────────────────────────
window.MHS_GAS_URL = 'https://script.google.com/macros/s/TU_DEPLOYMENT_ID/exec';

// ─── 2. Shim de google.script.run ───────────────────────────────────────────
// Solo se instala cuando no estamos dentro del propio entorno de GAS.
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
   * Builder — crea un objeto encadenable igual al original de GAS:
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
        // Devolvemos las propiedades propias normalmente
        if (Object.prototype.hasOwnProperty.call(target, prop)) return target[prop];
        if (prop === 'then' || prop === 'catch' || prop === 'finally') return undefined; // no es una Promise

        // Todo lo demás es el nombre del método de la API
        return function (payload) {
          var body = JSON.stringify(
            Object.assign({ action: prop }, payload || {})
          );

          fetch(w.MHS_GAS_URL, {
            method: 'POST',
            // text/plain evita el preflight CORS en GAS
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: body
          })
          .then(function (r) {
            if (!r.ok) throw new Error('HTTP ' + r.status);
            return r.json();
          })
          .then(function (data) {
            if (successFn) successFn(data);
          })
          .catch(function (err) {
            if (failureFn) failureFn(err);
            else console.warn('[MHS api.js] Error en ' + prop + ':', err);
          });
        };
      }
    });
  }

  w.google       = w.google || {};
  w.google.script = { run: Builder(null, null) };

  console.info('[MHS api.js] google.script.run instalado → ' + w.MHS_GAS_URL);

})(window);
