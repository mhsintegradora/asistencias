/**
 * MHS Integradora — gas-updates.gs  (versión GITHUB)
 * ──────────────────────────────────────────────────────────────────────────
 * Hoja GitHub separada del sistema en producción:
 *   https://docs.google.com/spreadsheets/d/14mjnnmOttQcGvINN-ka2f1HumtaFs1ln97GQ9YwUUj4
 *
 * ARQUITECTURA:
 *   GitHub Pages (HTML/JS)  →  fetch()  →  GAS Web App (code.gs)  →  Google Sheets
 *       Frontend                  HTTP        Backend + lógica          Base de datos
 *
 * Tu GAS sigue siendo el motor completo:
 *   - Lee/escribe Sheets  - Genera PDFs con DriveApp
 *   - Autenticación       - Notificaciones y tickets
 *   - TODO lo que hacía antes, exactamente igual
 *   - Solo cambia el canal: google.script.run → fetch() desde GitHub Pages
 *
 * INSTRUCCIONES DE SETUP:
 *   1. Abre tu hoja GITHUB en Extensiones → Apps Script
 *   2. Pega tu code.gs original completo
 *   3. Agrega al final el contenido de este archivo
 *   4. Implementar → Nueva implementación → Web App
 *      · Ejecutar como: Yo  |  Acceso: Cualquier usuario
 *   5. Copia la URL /exec y ponla en api.js → window.MHS_GAS_URL
 * ──────────────────────────────────────────────────────────────────────────
 */

// ════════════════════════════════════════════════════════════════════════════
//  ⚙️  CONFIGURACIÓN CENTRAL — solo cambiar este ID si mueves la hoja
// ════════════════════════════════════════════════════════════════════════════

var MHS_SHEETS_ID = '14mjnnmOttQcGvINN-ka2f1HumtaFs1ln97GQ9YwUUj4';

/**
 * Abre la hoja correcta sea script ligado o standalone.
 * Úsala en lugar de SpreadsheetApp.getActiveSpreadsheet() si quieres
 * poder cambiar de hoja solo tocando MHS_SHEETS_ID arriba.
 */
function getMHSSpreadsheet_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId() === MHS_SHEETS_ID) return active;
  } catch(e) {}
  return SpreadsheetApp.openById(MHS_SHEETS_ID);
}


// ════════════════════════════════════════════════════════════════════════════
//  1.  doPost — ROUTER HTTP PRINCIPAL
//      Agrega (o reemplaza) este doPost en tu code.gs
// ════════════════════════════════════════════════════════════════════════════

function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    // Guard: si e es undefined (prueba desde editor GAS o acceso GET)
    if (!e || !e.postData) {
      output.setContent(JSON.stringify({ ok: false, error: 'No POST data. Accede desde fetch() o usa el Web App URL.' }));
      return output;
    }
    var raw     = e.postData.contents || '{}';
    var payload = JSON.parse(raw || '{}');
    var action  = String(payload.action || '');
    delete payload.action;

    var ROUTES = {
      // ── Sesión / Auth ───────────────────────────────────────────────
      'apiMobileSessionBootstrap':    apiMobileSessionBootstrap,
      'apiMobileLogin':               apiMobileLogin,

      // ── Pase de lista ───────────────────────────────────────────────
      'apiMobileCargarPase':          apiMobileCargarPase,
      'apiMobileGuardarPase':         apiMobileGuardarPase,
      'apiGetPaseStatus':             apiGetPaseStatus,
      'apiCopiarJornadaAnterior':     apiCopiarJornadaAnterior,

      // ── Registros e incidencias ─────────────────────────────────────
      'apiGetRegistros':              apiGetRegistros,
      'apiRegistrarIncidenciaFormal': apiRegistrarIncidenciaFormal,
      'apiGetIncidencias':            apiGetIncidencias,
      'apiGetCalendarioIncidencias':  apiGetCalendarioIncidencias,

      // ── Turnos eventuales ────────────────────────────────────────────
      'apiRegistrarTurnoEventual':    apiRegistrarTurnoEventual,
      'apiGetTurnosEventuales':       apiGetTurnosEventuales,

      // ── Descansos / Descriptivo ──────────────────────────────────────
      'apiDescriptivoInit':           apiDescriptivoInit,
      'apiDescriptivoGuardar':        apiDescriptivoGuardar,

      // ── Liberación de fechas ─────────────────────────────────────────
      'apiLiberarFecha':              apiLiberarFecha,
      'apiRevocarFecha':              apiRevocarFecha,

      // ── Soporte / Tickets ────────────────────────────────────────────
      'apiEnviarTicket':              apiEnviarTicket,
      'apiGetSoporteHistorial':       apiGetSoporteHistorial,
      'apiGetMisTickets':             apiGetMisTickets,
      'apiCerrarTicket':              apiCerrarTicket,
      'apiGetInboxSoporte':           apiGetInboxSoporte,
      'apiSoportePollV3':             apiSoportePollV3,
      'apiSoporteMobileSync':         apiSoporteMobileSync,
      'apiResponderTicket':           apiResponderTicket,

      // ── Live events ──────────────────────────────────────────────────
      'apiLogLiveEvent':              apiLogLiveEvent,
      'apiGetLiveEvents':             apiGetLiveEvents,

      // ── CEO / Dashboard ──────────────────────────────────────────────
      'apiCEODashboard':              apiCEODashboard,
      'apiCEOResumenPeriodo':         apiCEOResumenPeriodo,
      'apiCEOLive':                   apiCEOLive,

      // ── Exportación PDF ──────────────────────────────────────────────
      // GAS genera el PDF en Drive y devuelve una URL. Sigue igual.
      'apiSupervisorReportePDF':      apiSupervisorReportePDF,
      'apiSupervisorReporteOpciones': apiSupervisorReporteOpciones,
      'apiCEOReportePDF':             apiCEOReportePDF,
      'apiPagosPendientesPDF':        apiPagosPendientesPDF,

      // ── RH Móvil ─────────────────────────────────────────────────────
      'apiRHMobilePanel':             apiRHMobilePanel,
      'apiRHMobileEmpleados':         apiRHMobileEmpleados,
      'apiPagoPendienteCrear':        apiPagoPendienteCrear,
      'apiPagosPendientesListar':     apiPagosPendientesListar,
      'apiPagoPendienteMarcarPagado': apiPagoPendienteMarcarPagado,
      'apiDarBajaEmpleado':           apiDarBajaEmpleado,
      'apiRHAsignacionesCargar':      apiRHAsignacionesCargar,
      'apiRHAsignacionGuardar':       apiRHAsignacionGuardar,
    };

    if (!action) {
      output.setContent(JSON.stringify({ ok:false, error:'Falta el campo "action".' }));
      return output;
    }
    var handler = ROUTES[action];
    if (typeof handler !== 'function') {
      output.setContent(JSON.stringify({ ok:false, error:'Acción desconocida: ' + action }));
      return output;
    }

    var result = handler(payload);
    output.setContent(JSON.stringify(result !== undefined ? result : { ok:true }));

  } catch(err) {
    Logger.log('doPost ERROR: ' + err.message);
    output.setContent(JSON.stringify({ ok:false, error: err.message || String(err) }));
  }

  return output;
}


// ════════════════════════════════════════════════════════════════════════════
//  2.  doGet — Redirige al GitHub Pages (descomenta si lo necesitas)
// ════════════════════════════════════════════════════════════════════════════

/*
function doGet(e) {
  var page  = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'mobile';
  var BASE  = 'https://mhsintegradora.github.io/asistencias/';
  var pages = { mobile:BASE+'mobile.html', index:BASE+'index.html',
                ceo:BASE+'ceomobile.html', router:BASE+'router.html' };
  var url = pages[page] || pages.mobile;
  return ContentService.createTextOutput(
    '<html><head><meta http-equiv="refresh" content="0;url='+url+'"></head></html>'
  ).setMimeType(ContentService.MimeType.HTML);
}
*/


// ════════════════════════════════════════════════════════════════════════════
//  3.  getJornadasPorSede_ — Filtro de turnos por supervisor
//      Agrega esta función a tu code.gs
// ════════════════════════════════════════════════════════════════════════════

/**
 * Devuelve { "OHORAN":["MATUTINO","VESPERTINO"], ... } para el supervisor.
 * Lee la pestaña "Asignaciones" de la hoja:
 *   Columna A: username  |  Columna B: sede  |  Columna C: jornada
 */
function getJornadasPorSede_(username) {
  try {
    var ss    = getMHSSpreadsheet_();
    var sheet = ss.getSheetByName('Asignaciones');
    if (!sheet) { Logger.log('Hoja "Asignaciones" no encontrada'); return {}; }

    var rows = sheet.getDataRange().getValues();
    var map  = {};
    var user = String(username || '').toLowerCase().trim();

    for (var i = 1; i < rows.length; i++) {
      var u = String(rows[i][0]||'').toLowerCase().trim();
      var s = String(rows[i][1]||'').toUpperCase().trim();
      var j = String(rows[i][2]||'').toUpperCase().trim();
      if (u !== user || !s || !j) continue;
      if (!map[s]) map[s] = [];
      if (map[s].indexOf(j) < 0) map[s].push(j);
    }
    return map;
  } catch(err) {
    Logger.log('getJornadasPorSede_ ERROR: ' + err);
    return {};
  }
}


// ════════════════════════════════════════════════════════════════════════════
//  4.  Modificar apiMobileSessionBootstrap — agregar jornadasPorSede
// ════════════════════════════════════════════════════════════════════════════

/*
  Busca tu función apiMobileSessionBootstrap en code.gs.
  Dentro del objeto return, agrega esta línea (marcada con ⭐):

  return {
    ok:              true,
    token:           token,
    role:            ses.role,
    fullName:        ses.fullName,
    sedesAsignadas:  sedesAsig,
    jornadasPorSede: getJornadasPorSede_(ses.username),   // ⭐ AGREGAR
    // ... demás campos existentes ...
  };
*/


// ════════════════════════════════════════════════════════════════════════════
//  5.  _deskGetPrioridad_ — DESBLOQUEO_FECHA siempre es prioridad ALTA
// ════════════════════════════════════════════════════════════════════════════

function _deskGetPrioridad_(payload) {
  var tipo = String(payload.tipo_ticket || '').toUpperCase();
  if (tipo === 'DESBLOQUEO_FECHA')   return 'ALTA';
  if (tipo === 'INCIDENCIA_URGENTE') return 'URGENTE';
  return payload.prioridad_override || 'MEDIA';
}

/*
  En tu apiEnviarTicket reemplaza donde defines la prioridad:
    var prioridad = 'MEDIA';                    ← ANTES
    var prioridad = _deskGetPrioridad_(payload); ← DESPUÉS
*/


// ════════════════════════════════════════════════════════════════════════════
//  6.  Hoja "Asignaciones" — crear en Google Sheets
//      ID: 14mjnnmOttQcGvINN-ka2f1HumtaFs1ln97GQ9YwUUj4
// ════════════════════════════════════════════════════════════════════════════

/*
 Crea una pestaña "Asignaciones" con estas columnas:

 ┌──────────┬──────────────────┬────────────┐
 │ Username │ Sede             │ Jornada    │
 ├──────────┼──────────────────┼────────────┤
 │ fer      │ OHORAN           │ MATUTINO   │
 │ fer      │ OHORAN           │ VESPERTINO │
 │ ivan     │ OHORAN           │ VESPERTINO │
 │ ivan     │ OHORAN           │ NOCTURNO   │
 │ alex     │ OHORAN           │ NOCTURNO   │
 │ fer      │ SAN JOSE TEC     │ MATUTINO   │
 │ ivan     │ SAN JOSE TEC     │ VESPERTINO │
 │ alex     │ SAN JOSE TEC     │ NOCTURNO   │
 └──────────┴──────────────────┴────────────┘

 ✅ Username = igual a la columna "username" de tu hoja de Usuarios
 ✅ Sede     = igual al nombre exacto como aparece en los pases
 ✅ Jornada  = MATUTINO · VESPERTINO · NOCTURNO (en mayúsculas)
*/


// ════════════════════════════════════════════════════════════════════════════
//  7.  ¿Qué sigue funcionando igual? TODO.
// ════════════════════════════════════════════════════════════════════════════

/*
 ✅  Pase de lista          — guardar, cargar, copiar jornada anterior
 ✅  Incidencias formales   — registrar, calendario, detalle por día
 ✅  Turnos eventuales      — todos los pasos del wizard
 ✅  Desbloqueo de fechas   — liberar / revocar (módulo nuevo)
 ✅  Exportar PDF           — GAS genera el PDF en Drive, devuelve URL → frontend la abre
 ✅  Soporte / Chat         — inbox, historial, respuestas, cerrar ticket
 ✅  Live alerts            — polling, toast, vibración, beep
 ✅  CEO Dashboard          — resumen en tiempo real, gráficas
 ✅  RH móvil               — alta/baja, pagos, asignaciones
 ✅  Filtro de jornadas     — Fer/Ivan/Alex ven solo sus turnos por sede (NUEVO)

 DIFERENCIA GitHub Pages vs GAS nativo:
   Antes:  HTML dentro de GAS, servido por doGet() como HtmlService
   Ahora:  HTML en GitHub Pages, llama a GAS via fetch()
   Motor:  code.gs + Sheets = exactamente igual, sin cambios
*/
