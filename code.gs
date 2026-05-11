/***********************
 * MHS RH HUB · Apps Script · Code.gs
 * ─────────────────────────────────────
 * FIXES APLICADOS (vs versión anterior):
 *
 * 1. BUG SEDES: getSedes_() ahora construye la lista UNIENDO
 *    CAT_SEDES + CONTRATOS_2026 automáticamente, con fallback a
 *    abrev derivada si una sede existe en contratos pero no en el catálogo.
 *    → Ya no rompe si agregas sedes nuevas sin actualizar CAT_SEDES.
 *
 * 2. BUG FECHAS (HORA_INICIO / HORA_FIN): en Sheets, las celdas con
 *    formato de hora llegan como objeto Date de JS al usar getValues().
 *    String(dateObj) produce "Mon Jan 01 1899 06:00:00 GMT-0600" que
 *    rompe todo. Se normaliza con formatCellValue_() que detecta Date
 *    y lo convierte a texto HH:MM limpio usando Utilities.formatDate.
 *
 * 3. BUG CP FLOAT: CP llega como 97246.0 (float). Se convierte a
 *    entero/string limpio antes de guardarse.
 *
 * 4. SYNC LIVE: apiGetEmpleadosPorSede() y getSedes_() leen siempre
 *    directamente del spreadsheet real (CONTRATOS_2026) para que
 *    cualquier alta/baja manual desde la hoja se refleje al instante
 *    sin redeployar.
 *
 * 5. ABREV FALLBACK: si una sede no tiene abreviatura en CAT_SEDES
 *    se genera una automáticamente (primeras 3 letras de cada palabra).
 *
 * 6. FECHA ISO SAFE: todas las comparaciones de fecha usan strings
 *    ISO yyyy-MM-dd. Nunca se comparan objetos Date directamente.
 ***********************/

// ──────────────────────────────────────────
//  CONFIGURACIÓN GLOBAL
// ──────────────────────────────────────────
const TZ = "America/Merida";

const SH_CONTRATOS   = "CONTRATOS_2026";
const SH_CAT_SEDES   = "CAT_SEDES";
const SH_PL          = "PASE_LISTA_V2";
const SH_CFG_CODES   = "CONFIG_ASISTENCIAS";
const SH_USERS       = "USUARIOS";
const SH_LOCK        = "CONFIG";
const SH_ESTADO_PASE = "ESTADO_PASE";
const SH_BAJAS       = "BAJAS";
const SH_INCIDENCIAS = "INCIDENCIAS";
const SH_LIVE_EVENTS = "LIVE_EVENTS";

const COL_STATUS_TRAB = "STATUS_TRABAJADOR";
const COL_FECHA_BAJA  = "FECHA_BAJA";
const COL_MOTIVO_BAJA = "MOTIVO_BAJA";
const COL_BAJA_POR    = "BAJA_CAPTURADO_POR";
const COL_BAJA_TS     = "BAJA_TS";

const HDR_ESTADO = ["ID","ESTADO","FECHA_BAJA_ISO","MOTIVO","CAPTURADO_POR","TS_ISO"];
const HDR_PL     = ["FECHA_ISO","QUINCENA","SEDE","ABREV","ID","NOMBRE","CODIGO","CAPTURADO_POR","TS_ISO","KEY"];
const HDR_BAJAS  = ["FECHA_BAJA","SEDE","ID","NOMBRE","MOTIVO","CAPTURADO_POR","TIMESTAMP"];
const HDR_INC    = ["FECHA_ISO","SEDE","ID","NOMBRE","TIPO","OBSERVACIONES","CAPTURADO_POR","TS_ISO"];
const HDR_LIVE   = ["TS_ISO","FECHA_ISO","USERNAME","FULL_NAME","ROLE","TIPO","SEDE","JORNADA","MENSAJE","META_JSON","TARGET_ROLES","LEIDO_POR"];

// ID de la carpeta Drive para exportar PDFs de reportes.
// Este es el ID de tu carpeta "Asistencias AUTOMATIZADAS 2026".
const DRIVE_PDF_FOLDER_ID = "1-S7eh31oHuYa2ioiJvh1l8LFKTi5X_jz";

// ── ID de la hoja de Google Sheets para la versión GITHUB ──────────────────
// Cambia este ID para apuntar a producción cuando sea el momento.
const MHS_SHEETS_ID = "14mjnnmOttQcGvINN-ka2f1HumtaFs1ln97GQ9YwUUj4";

/** Abre la hoja correcta (ligada o standalone). */
function getMHSSpreadsheet_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId() === MHS_SHEETS_ID) return active;
  } catch(e) {}
  return SpreadsheetApp.openById(MHS_SHEETS_ID);
}

// Modelo Claude para análisis de imágenes (escaneo de credenciales)
const CLAUDE_MODEL = "claude-opus-4-6";

// ══════════════════════════════════════════
//  MENÚ
// ══════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("MHS RH HUB")
    .addItem("A) Setup completo (crear TODO)", "setupAll")
    .addItem("B) Sincronizar CONTRATOS_2026 (arreglar columnas)", "menuSyncContratos")
    .addSeparator()
    .addItem("C) Diagnóstico rápido", "diagV2")
    .addToUi();
}

// ══════════════════════════════════════════
//  SETUP TOTAL
// ══════════════════════════════════════════
function setupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureSheetIfMissing(ss, SH_CAT_SEDES,   [["SEDE","ABREV","ULTIMO_FOLIO_2026"]]);
  _ensureSheetIfMissing(ss, SH_CFG_CODES,   [
    ["CODIGO","DESCRIPCION","DESCUENTA"],
    ["A","Asistio",0],["AF","Asistio (Feriado)",0],["F","Falto",1],
    ["PCG","Permiso con goce",0],["PSG","Permiso sin goce",1],
    ["I","Incapacidad",0],["DT","Doble turno",0],
    ["DS","Descanso",0],["FER","Feriado",0],["INH","Inhabilitado",0],
  ]);
  _ensureSheetIfMissing(ss, SH_PL,          [HDR_PL]);
  _ensureSheetIfMissing(ss, SH_USERS,       [
    ["USERNAME","PASSWORD","ROLE","ACTIVE","FULL_NAME","NOTES"],
    ["usuario","1234","USER",true,"Usuario demo",""],
    ["admin","admin123","ADMIN",true,"Admin demo",""],
    ["super","super123","SUPERADMIN",true,"Super demo",""],
    ["ceo","ceo123","CEO",true,"CEO demo",""],
  ]);
  _ensureSheetIfMissing(ss, SH_ESTADO_PASE, [HDR_ESTADO]);
  _ensureSheetIfMissing(ss, SH_BAJAS,       [HDR_BAJAS]);
  _ensureSheetIfMissing(ss, SH_INCIDENCIAS, [HDR_INC]);
  _ensureSheetIfMissing(ss, SH_LIVE_EVENTS, [HDR_LIVE]);

  let lock = ss.getSheetByName(SH_LOCK);
  if (!lock) lock = ss.insertSheet(SH_LOCK);
  if (!lock.getRange("A1").getValue()) lock.getRange("A1").setValue("FECHA_CIERRE");
  if (!lock.getRange("B1").getValue()) {
    const d = new Date(); d.setDate(d.getDate() - 1);
    lock.getRange("B1").setValue(Utilities.formatDate(d, TZ, "yyyy-MM-dd"));
  }

  ensureContratosRHColumns_();

  SpreadsheetApp.getUi().alert(
    "Setup MHS RH HUB completado.\n\n" +
    "Hojas verificadas/creadas:\n" +
    "CONTRATOS_2026 (columnas RH aseguradas)\n" +
    "CAT_SEDES, CONFIG_ASISTENCIAS, PASE_LISTA_V2\n" +
    "USUARIOS, ESTADO_PASE, CONFIG, BAJAS, INCIDENCIAS"
  );
}

function _ensureSheetIfMissing(ss, name, rows) {
  let sh = ss.getSheetByName(name);
  const isNew = !sh;
  if (isNew) sh = ss.insertSheet(name);
  if (isNew && rows && rows.length) {
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, rows[0].length);
  }
  return sh;
}

function diagV2() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const check = [SH_PL, SH_CAT_SEDES, SH_CONTRATOS, SH_USERS, SH_BAJAS, SH_INCIDENCIAS];
  const sedes = getSedes_();
  const msg   = [
    ...check.map(n => n + ": " + (ss.getSheetByName(n) ? "OK" : "FALTA")),
    "",
    "Sedes detectadas: " + sedes.length,
    ...sedes.map(s => "  " + s.sede + " (" + s.abrev + ")")
  ].join("\n");
  SpreadsheetApp.getUi().alert(msg);
}

// ══════════════════════════════════════════
//  FIX #2: NORMALIZAR VALORES DE CELDA
//  Convierte Date -> string seguro, float -> int limpio
// ══════════════════════════════════════════
function formatCellValue_(val) {
  if (val === null || val === undefined || val === "") return "";

  // Si la celda tiene formato de hora/fecha, getValues() devuelve un Date object.
  // String(dateObj) = "Mon Jan 01 1900 06:00:00 GMT-0600" -> ROMPE comparaciones.
  if (val instanceof Date) {
    try {
      const y = val.getFullYear();
      // Fecha base de Sheets para horas puras es 1899/1900
      if (y <= 1900) {
        return Utilities.formatDate(val, TZ, "HH:mm");
      }
      return Utilities.formatDate(val, TZ, "yyyy-MM-dd");
    } catch (e) {
      return "";
    }
  }

  // Float que parece entero: 97246.0 -> "97246"
  if (typeof val === "number") {
    if (Number.isInteger(val)) return String(val);
    if (Math.abs(val - Math.round(val)) < 0.001) return String(Math.round(val));
    return String(val);
  }

  return String(val).trim();
}

// ══════════════════════════════════════════
//  FIX #1 + #4 + #5: SEDES
//  Construye la lista uniendo CAT_SEDES + CONTRATOS_2026 en tiempo real.
//  Si una sede existe en contratos pero no en el catalogo, genera abrev automatica.
// ══════════════════════════════════════════
function getSedes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Leer catálogo CAT_SEDES (fuente de verdad para abreviaturas)
  const catMap = {};
  const shCat  = ss.getSheetByName(SH_CAT_SEDES);
  if (shCat && shCat.getLastRow() >= 2) {
    const v = shCat.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      const sede  = String(v[i][0] || "").trim();
      const abrev = String(v[i][1] || "").trim();
      if (sede && abrev) catMap[sede] = abrev;
    }
  }

  // 2) Leer todas las sedes que realmente tienen empleados en CONTRATOS_2026
  const sedesEnContratos = new Set();
  const shCon = ss.getSheetByName(SH_CONTRATOS);
  if (shCon && shCon.getLastRow() >= 2) {
    const v   = shCon.getDataRange().getValues();
    const hdr = v[0].map(normHdr_); // FIX
    const cS  = hdr.indexOf("SEDE");
    const cSt = hdr.indexOf(COL_STATUS_TRAB);
    if (cS >= 0) {
      for (let i = 1; i < v.length; i++) {
        // Solo incluir empleados activos (no dados de baja)
        const st = cSt >= 0 ? String(v[i][cSt] || "ACTIVO").toUpperCase().trim() : "ACTIVO";
        if (st === "BAJA") continue;
        const s = String(v[i][cS] || "").trim();
        if (s) sedesEnContratos.add(s);
      }
    }
  }

  // 3) Union de ambas fuentes
  const all = new Set([...Object.keys(catMap), ...sedesEnContratos]);

  const out = [];
  all.forEach(sede => {
    let abrev = catMap[sede];
    if (!abrev) {
      // FIX #5: generar abreviatura automatica si no esta en catalogo
      // Toma la primera letra de cada palabra significativa (mas de 2 chars)
      abrev = sede.split(/\s+/)
        .filter(w => w.length > 2 && !["DE","LA","EL","LOS","LAS","DEL"].includes(w.toUpperCase()))
        .map(w => w[0])
        .join("")
        .toUpperCase()
        .substring(0, 4);
      if (!abrev) abrev = sede.substring(0, 3).toUpperCase();
    }
    out.push({ sede, abrev });
  });

  out.sort((a, b) => a.sede.localeCompare(b.sede, "es"));
  return out;
}

function abrevBySede_(sede) {
  if (!sede) return "";
  const found = getSedes_().find(x => x.sede === sede);
  return found ? found.abrev : "";
}

// ══════════════════════════════════════════
//  HELPERS FECHA
// ══════════════════════════════════════════
// Normaliza encabezados de Sheets: elimina espacios, saltos de linea,
// caracteres invisibles y convierte a mayusculas.
// Usar SIEMPRE en lugar de String(h).trim() al leer headers de getValues()
function normHdr_(h) {
  return String(h || "")
    .replace(/[\r\n\t\u00A0\u200B\u200C\u200D\uFEFF]/g, "") // chars invisibles
    .trim()
    .toUpperCase();
}

function todayISO_()  { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd"); }
function isoNow_()    { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss"); }
function norm_(s)     { return String(s || "").trim(); }
function upper_(s)    { return norm_(s).toUpperCase(); }
function isoValid_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim()); }
function makeKey_(f, a, id) { return f + "|" + upper_(a) + "|" + norm_(id); }
function quincenaFromISO_(f) {
  const d = Number(String(f).split("-")[2] || 0);
  return d <= 15 ? "Q1" : "Q2";
}

// ══════════════════════════════════════════
//  USUARIOS
// ══════════════════════════════════════════
function readUsers_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
  if (!sh || sh.getLastRow() < 2) return [];
  const v   = sh.getDataRange().getValues();
  const idx = {};
  v[0].forEach((h, i) => { idx[String(h || "").trim().toUpperCase()] = i; });
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    out.push({
      username: String(r[idx.USERNAME] || "").toLowerCase().trim(),
      password: String(r[idx.PASSWORD] || ""),
      role:     String(r[idx.ROLE]     || "USER").toUpperCase().trim(),
      active:   String(r[idx.ACTIVE]).toLowerCase().trim() === "true",
      fullName: String(r[idx.FULL_NAME] || "")
    });
  }
  return out;
}

// ══════════════════════════════════════════
//  AUTH
// ══════════════════════════════════════════
const TTL = 60 * 60 * 10;

function issueToken_(username, role) {
  const t = Utilities.getUuid().replace(/-/g, "");
  CacheService.getScriptCache().put("T:" + t, JSON.stringify({ u: username, r: role }), TTL);
  return t;
}
function readToken_(token) {
  const key = "T:" + String(token || "").trim();
  const raw = CacheService.getScriptCache().get(key);
  if (!raw) return null;
  // Sesión viva mientras la app esté abierta: cada llamada válida renueva el TTL.
  // CacheService puede limitar el TTL máximo, pero el polling activo mantiene la sesión fresca.
  try { CacheService.getScriptCache().put(key, raw, TTL); } catch(e) {}
  try { return JSON.parse(raw); } catch (e) { return null; }
}
function requireSession_(token) {
  const s = readToken_(token);
  if (!s) throw new Error("No has iniciado sesion.");
  return { username: s.u, role: s.r };
}
function isAdminRole_(role) {
  const r = String(role || "").toUpperCase();
  return r === "ADMIN" || r === "SUPERADMIN" || r === "CEO" || r === "SOPORTE";
}

// ══════════════════════════════════════════
//  REGLAS DE EDICIÓN
// ══════════════════════════════════════════
function canEdit_(role, fechaISO) {
  var r = String(role || "").toUpperCase();
  if (r === "SUPERADMIN") return { ok: true, reason: "", graceExpiry: null };

  var now  = new Date();
  var hoy  = todayISO_();
  var ayerParts = hoy.split("-");
  var ayerD = new Date(Number(ayerParts[0]), Number(ayerParts[1]) - 1, Number(ayerParts[2]) - 1);
  var ayer  = Utilities.formatDate(ayerD, TZ, "yyyy-MM-dd");

  // Verificar si Superadmin liberó esta fecha en CONFIG
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shCfg = ss.getSheetByName(SH_LOCK);
    if (shCfg && shCfg.getLastRow() >= 1) {
      var cfgVals = shCfg.getDataRange().getValues();
      for (var ci = 0; ci < cfgVals.length; ci++) {
        if (String(cfgVals[ci][0] || "").trim() === "LIB_" + fechaISO &&
            String(cfgVals[ci][1] || "").trim() === "1") {
          return { ok: true, reason: "", graceExpiry: null, liberada: true };
        }
      }
    }
  } catch (e_) {}

  if (fechaISO === hoy) {
    // Día actual: OK. Gracia vence mañana a las 12:00 pm (18:00 UTC)
    var hoyParts = hoy.split("-");
    var limiteGracia = new Date(Date.UTC(
      Number(hoyParts[0]), Number(hoyParts[1]) - 1, Number(hoyParts[2]) + 1, 18, 0, 0
    ));
    var graceStr = Utilities.formatDate(limiteGracia, TZ, "yyyy-MM-dd HH:mm");
    return { ok: true, reason: "", graceExpiry: graceStr };
  }

  if (fechaISO === ayer) {
    // Día anterior: OK solo si ahora < 12:00 pm (18:00 UTC) de HOY
    var ayerISOParts = hoy.split("-");
    var limiteHoy = new Date(Date.UTC(
      Number(ayerISOParts[0]), Number(ayerISOParts[1]) - 1, Number(ayerISOParts[2]), 18, 0, 0
    ));
    if (now < limiteHoy) {
      var graceAyer = Utilities.formatDate(limiteHoy, TZ, "yyyy-MM-dd HH:mm");
      return { ok: true, reason: "", graceExpiry: graceAyer };
    }
    return {
      ok: false,
      reason: "El período de gracia para " + ayer + " venció a las 12:00 pm de hoy. Solicita al Superadmin.",
      graceExpiry: null
    };
  }

  if (fechaISO < hoy) {
    return {
      ok: false,
      reason: "Solo puedes asentar el día actual o el día anterior hasta las 12:00 pm. Solicita al Superadmin.",
      graceExpiry: null
    };
  }

  return {
    ok: false,
    reason: "No puedes registrar asistencias para fechas futuras (" + fechaISO + ").",
    graceExpiry: null
  };
}
function isLockedByNomina_(role, fechaISO) {
  var r = String(role || "").toUpperCase();
  if (r === "SUPERADMIN") return { locked: false, reason: "" };

  // FIX V18.1 — Gracia operativa para supervisores USER
  // Antes, si CONFIG!B1 (FECHA_CIERRE) incluía el día anterior,
  // el candado de nómina bloqueaba al USER aunque canEdit_() dijera que
  // todavía estaba dentro del periodo de gracia.
  // Regla correcta: si canEdit_ permite capturar por día actual, día anterior
  // dentro de gracia o fecha liberada por SUPERADMIN, no se aplica el candado.
  try {
    var permisoGracia = canEdit_(r, fechaISO);
    if (permisoGracia && permisoGracia.ok) {
      return { locked: false, reason: "" };
    }
  } catch (eGrace) {}

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_LOCK);
  if (!sh) return { locked: false, reason: "" };
  const key    = String(sh.getRange("A1").getValue() || "").trim();
  const cierre = String(sh.getRange("B1").getValue() || "").trim();
  if (key !== "FECHA_CIERRE" || !isoValid_(cierre)) return { locked: false, reason: "" };
  if (fechaISO <= cierre) return { locked: true, reason: "cerrado hasta " + cierre };
  return { locked: false, reason: "" };
}

// ══════════════════════════════════════════
//  CONTRATOS: asegurar columnas RH
// ══════════════════════════════════════════
function ensureContratosRHColumns_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CONTRATOS);
  if (!sh || sh.getLastColumn() === 0) return;

  // FIX: usar normHdr_ para evitar falsos negativos por chars invisibles
  const hdr  = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(normHdr_);
  const have = new Set(hdr.filter(Boolean));

  // FIX: DIA_DESCANSO incluida — si no esta en el sheet se agrega
  const need = ["DIA_DESCANSO", COL_STATUS_TRAB, COL_FECHA_BAJA, COL_MOTIVO_BAJA, COL_BAJA_POR, COL_BAJA_TS];
  const add  = need.filter(x => !have.has(x));
  if (add.length) sh.getRange(1, sh.getLastColumn() + 1, 1, add.length).setValues([add]);

  // Releer headers actualizados
  const hdr2 = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(normHdr_);
  const isSt = hdr2.indexOf(COL_STATUS_TRAB);
  if (isSt >= 0 && sh.getLastRow() >= 2) {
    const rng = sh.getRange(2, isSt + 1, sh.getLastRow() - 1, 1);
    const v   = rng.getValues();
    let changed = false;
    for (let i = 0; i < v.length; i++) {
      if (!String(v[i][0] || "").trim()) { v[i][0] = "ACTIVO"; changed = true; }
    }
    if (changed) rng.setValues(v);
  }
}

// ══════════════════════════════════════════
//  ESTADO_PASE
// ══════════════════════════════════════════
function readEstadoPaseMap_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ESTADO_PASE);
  if (!sh || sh.getLastRow() < 2) return {};
  const v = sh.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < v.length; i++) {
    const id = String(v[i][0] || "").trim();
    if (id) out[id] = { estado: String(v[i][1] || "ACTIVO").toUpperCase().trim(), row: i + 1 };
  }
  return out;
}
function upsertEstadoPase_(id, estado, fechaBajaISO, motivo, capt) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ESTADO_PASE);
  if (!sh) return;
  const map = readEstadoPaseMap_();
  const ts  = isoNow_();
  const est = String(estado || "ACTIVO").toUpperCase().trim();
  if (map[id] && map[id].row) {
    const r = map[id].row;
    sh.getRange(r, 2).setValue(est);
    sh.getRange(r, 3).setValue(fechaBajaISO || "");
    sh.getRange(r, 4).setValue(motivo || "");
    sh.getRange(r, 5).setValue(capt || "");
    sh.getRange(r, 6).setValue(ts);
  } else {
    sh.appendRow([id, est, fechaBajaISO || "", motivo || "", capt || "", ts]);
  }
}

// ══════════════════════════════════════════
//  FIX #6: LECTURA PASE LISTA con fechas ISO seguras
// ══════════════════════════════════════════
function readDayMarks_(fechaISO, abrev) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  if (!sh || sh.getLastRow() < 2) return {};
  const v   = sh.getDataRange().getValues();
  const AB  = String(abrev).toUpperCase().trim();
  const out = {};
  for (let i = 1; i < v.length; i++) {
    const f = formatCellValue_(v[i][0]); // FIX: nunca comparar Date directamente
    const a = String(v[i][3] || "").trim().toUpperCase();
    if (f !== fechaISO || a !== AB) continue;
    const id  = formatCellValue_(v[i][4]);
    const cod = String(v[i][6] || "").trim().toUpperCase();
    if (id) out[id] = cod;
  }
  return out;
}
function readDayRowsByKey_(fechaISO, abrev) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  if (!sh || sh.getLastRow() < 2) return {};
  const v   = sh.getDataRange().getValues();
  const AB  = String(abrev).toUpperCase().trim();
  const map = {};
  for (let i = 1; i < v.length; i++) {
    const f = formatCellValue_(v[i][0]); // FIX
    const a = String(v[i][3] || "").trim().toUpperCase();
    if (f !== fechaISO || a !== AB) continue;
    const key = String(v[i][9] || "").trim();
    if (key) map[key] = { row: i + 1, codigo: String(v[i][6] || "").trim().toUpperCase() };
  }
  return map;
}

// ══════════════════════════════════════════
//  WEBAPP ROUTER
// ══════════════════════════════════════════

/**
 * doGet — punto de entrada de la WebApp.
 *
 * Regla estable:
 *  1. ?api=xxx      → respuesta JSON
 *  2. ?page=mobile  → sirve mobile.html
 *  3. ?page=desktop → sirve index.html
 *  4. Sin parámetro → sirve router.html
 *
 * Nota: Apps Script no expone el User-Agent en `e`, así que la detección
 * real de dispositivo se hace del lado del cliente en router.html.
 */
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  if (params.api) {
    return json_(apiRouter_(e));
  }

  var page = String(params.page || "").toLowerCase();
  if (page === "mobile") {
    return serveHtml_("mobile", "MHS · Pase de Lista");
  }
  if (page === "desktop") {
    return serveHtml_("index", "MHS RH HUB");
  }
  if (page === "ceo") {
    return serveHtml_("ceo", "MHS · CEO LIVE DASHBOARD");
  }
  if (page === "ceo_mobile") {
    return serveHtml_("ceo_mobile", "MHS · CEO LIVE MOBILE");
  }

  return serveHtml_("router", "MHS Integradora");
}

/** Sirve un archivo HTML con configuración estándar. */
function serveHtml_(filename, title) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .setTitle(title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Devuelve la URL final a la que debe navegar el router.
 * El cambio de ubicación lo hace router.html con top.location.href.
 */
function getRedirectUrl(isMobile) {
  var url = ScriptApp.getService().getUrl();
  return url + "?page=" + (isMobile ? "mobile" : "desktop");
}

function getAppPageUrl(page) {
  var url = ScriptApp.getService().getUrl();
  return url + "?page=" + encodeURIComponent(String(page || "desktop"));
}

function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    if (!e || !e.postData) {
      output.setContent(JSON.stringify({ ok:false, error:'No POST data. Usa fetch() hacia la URL /exec.' }));
      return output;
    }
    var raw     = e.postData.contents || '{}';
    var payload = JSON.parse(raw);
    var action  = String(payload.action || '');
    delete payload.action;

    // Si no viene action, intentar ruta legacy por e.parameter.api (compatibilidad)
    if (!action && e.parameter && e.parameter.api) {
      return json_(apiRouter_(e));
    }

    var ROUTES = {
      'apiMobileSessionBootstrap':    apiMobileSessionBootstrap,
      'apiMobileLogin':               apiLogin,
      'apiMobileCargarPase':          apiMobileCargarPase,
      'apiMobileGuardarPase':         apiGuardarPaseListaMobile,
      'apiGetPaseStatus':             apiGetPaseStatus,
      'apiCopiarJornadaAnterior':     apiSupervisorCopiarJornada,
      'apiGetRegistros':              apiGetRegistros,
      'apiRegistrarIncidenciaFormal': apiRegistrarIncidenciaFormal,
      'apiGetIncidencias':            apiGetIncidencias,
      'apiGetCalendarioIncidencias':  apiGetCalendarioIncidencias,
      'apiRegistrarTurnoEventual':    apiRegistrarTurnoEventual,
      'apiGetTurnosEventuales':       apiGetTurnosEventuales,
      'apiDescriptivoInit':           apiDescriptivoInit,
      'apiDescriptivoGuardar':        apiDescriptivoGuardar,
      'apiLiberarFecha':              apiLiberarFecha,
      'apiRevocarFecha':              apiRevocarFecha,
      'apiEnviarTicket':              apiEnviarTicket,
      'apiGetSoporteHistorial':       apiGetSoporteHistorial,
      'apiGetMisTickets':             apiGetMisTickets,
      'apiCerrarTicket':              apiCerrarTicket,
      'apiGetInboxSoporte':           apiGetInboxSoporte,
      'apiSoportePollV3':             apiSoportePollV3,
      'apiSoporteMobileSync':         apiSoporteMobileSync,
      'apiResponderTicket':           apiResponderTicket,
      'apiLogLiveEvent':              apiLogLiveEvent,
      'apiGetLiveEvents':             apiGetLiveEvents,
      'apiCEODashboard':              apiCEODashboard,
      'apiCEOResumenPeriodo':         apiCEOResumenPeriodo,
      'apiCEOLive':                   apiCEOLive,
      'apiSupervisorReportePDF':      apiSupervisorReportePDF,
      'apiSupervisorReporteOpciones': apiSupervisorReporteOpciones,
      'apiCEOReportePDF':             apiCEOReportePDF,
      'apiPagosPendientesPDF':        apiPagosPendientesPDF,
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
      output.setContent(JSON.stringify({ ok:false, error:'Falta campo "action".' }));
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
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function apiRouter_(e) {
  try {
    var p    = e.parameter || {};
    var api  = String(p.api || "").toLowerCase();
    var body = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    if (api === "setup")           { setupAll(); return { ok: true }; }
    if (api === "sedes")           return { ok: true, sedes: getSedes_() };
    if (api === "login")           return apiLogin_(body);
    if (api === "logout")          return apiLogout_(body);
    if (api === "reporte")         return apiReporteDiario(body);
    if (api === "exportar")        return apiExportarQuincena(body);
    if (api === "baja")            return apiDarBajaEmpleado(body);
    if (api === "alta")            return apiAltaEmpleado(body);
    if (api === "incidencia")      return apiRegistrarIncidencia(body);
    if (api === "incformal")       return apiRegistrarIncidenciaFormal(body);
    if (api === "foto")            return apiAnalizarFoto(body);
    if (api === "pdf")             return apiExportarPDFReporte(body);
    if (api === "sheets")          return apiExportarSheets(body);
    if (api === "desccanso")       return apiCambiarDiaDescanso(body);
    if (api === "diag")            return apiDiagContratos(body);
    if (api === "liberarfecha")    return apiLiberarFecha(body);
    if (api === "incidenciassede") return apiGetIncidenciasSede(body);
    if (api === "incidenciascal")  return apiGetIncidenciasCalendario(body);
    if (api === "turnoevental")    return apiRegistrarTurnoEventual(body);
    if (api === "turnoslist")      return apiGetTurnosEventuales(body);
    if (api === "exportinc")       return apiExportarIncidencias(body);
    if (api === "exportte")        return apiExportarTurnosEventuales(body);
    if (api === "ceodashboard")    return apiGetCEODashboard(body);
    if (api === "exportceopdf")    return apiExportCEOPDF(body);
    if (api === "compensacion")    return apiCompensarDescanso(body);
    if (api === "faltas_por_sede") return apiGetFaltasPorSede(body);
    if (api === "sinmarcar")       return apiGetSinMarcar(body);
    if (api === "cdt_crear")       return apiCDTCrear(body);
    if (api === "cdt_listar")      return apiCDTListar(body);
    if (api === "cdt_cancelar")    return apiCDTCancelar(body);
    if (api === "cdt_check")       return apiCDTCheckFecha(body);
    if (api === "cdt_porsede")     return apiCDTGetPorSede(body);
    if (api === "registros")       return apiGetRegistros(body);
    if (api === "soporte")         return apiEnviarSoporte(body);
    if (api === "soporte_hist")    return apiGetSoporteHistorial(body);
    if (api === "soporte_pendiente") return apiGetPendienteSoporte(body);
    if (api === "soporte_responder")  return apiResponderSoporte(body);
    if (api === "soporte_leido")      return apiMarcarLeidoSoporte(body);
    if (api === "soporte_inbox")      return apiGetInboxSoporte(body);
    if (api === "rhpro_export")    return apiRHProExport(body);
    if (api === "rhpro_dashboard") return apiRHProDashboard(body);
    if (api === "geocode")         return apiReverseGeocode(body);
    if (api === "exportnomina")    return apiExportarNomina(body);
    if (api === "ticket")          return apiEnviarTicket(body);
    if (api === "sede_jornada")    return apiGetSedeJornadaUsuario(body);
    if (api === "pendientes_asist")return apiGetPendientesAsistencia(body);
    if (api === "cerrar_ticket")   return apiCerrarTicket(body);
    if (api === "live_log")        return apiLogLiveEvent(body);
    if (api === "live_events")     return apiGetLiveEvents(body);
    if (api === "rhpro_get_asign")   return apiRHProGetAsignaciones(body);
    if (api === "rhpro_save_asign")  return apiRHProGuardarAsignacion(body);
    if (api === "rhpro_export_asign")return apiRHProExportarAsignaciones(body);
    if (api === "emps_jornada")    return apiGetEmpleadosPorSedeJornada(body);
    if (api === "mis_tickets")     return apiGetMisTickets(body);
    if (api === "sup_reporte_pdf")return apiSupervisorReportePDF(body);
    if (api === "sup_reporte_opts")return apiSupervisorReporteOpciones(body);
    return { ok: false, error: "API no valida" };
  } catch (err) {
    return { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
}

// ══════════════════════════════════════════
//  ENDPOINTS PUBLICOS (google.script.run)
// ══════════════════════════════════════════

function apiGetSedes() {
  // FIX #1 + #4: siempre en vivo desde CONTRATOS_2026 + CAT_SEDES
  return getSedes_();
}

function apiLogin(payload) {
  const r = apiLogin_(payload || {});
  return r.ok
    ? { ok: true, token: r.token, role: r.role, username: r.username }
    : { ok: false, message: r.message || r.error || "Login fallido" };
}
function apiLogout(payload) { apiLogout_(payload || {}); return { ok: true }; }

function apiGetPermisosEdicion(payload) {
  const token    = String((payload && payload.token) || "").trim();
  const fechaISO = String((payload && payload.fecha)  || "").trim();
  const ses = requireSession_(token);
  if (!isoValid_(fechaISO)) throw new Error("Fecha invalida (yyyy-MM-dd)");
  const perm = canEdit_(ses.role, fechaISO);
  const lock = isLockedByNomina_(ses.role, fechaISO);
  const ok   = perm.ok && !lock.locked;
  const reason = !perm.ok ? perm.reason : (lock.locked ? "Nomina cerrada: " + lock.reason : "");
  return {
    canEdit:      !!ok,
    isAdmin:      isAdminRole_(ses.role),
    isSuperAdmin: String(ses.role).toUpperCase() === "SUPERADMIN",
    reason:       reason || "",
    graceExpiry:  perm.graceExpiry || null,
    liberada:     perm.liberada    || false
  };
}

function apiGetCandadoNomina(payload) {
  const token    = String((payload && payload.token) || "").trim();
  const fechaISO = String((payload && payload.fecha)  || "").trim();
  const ses = requireSession_(token);
  if (!isoValid_(fechaISO)) throw new Error("Fecha invalida");
  const lock = isLockedByNomina_(ses.role, fechaISO);
  return { locked: !!lock.locked, reason: lock.reason || "" };
}

// ══════════════════════════════════════════
//  FIX #1+#2+#4: EMPLEADOS POR SEDE
//  Lee CONTRATOS_2026 en vivo, normaliza valores de celda
// ══════════════════════════════════════════
function apiGetEmpleadosPorSede(sede) {
  const sedeN = String(sede || "").trim();
  if (!sedeN) throw new Error("Falta sede");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_CONTRATOS);
  if (!sh) throw new Error("No existe CONTRATOS_2026");

  ensureContratosRHColumns_();
  if (sh.getLastRow() < 2) return [];

  const raw = sh.getDataRange().getValues();
  const hdr = raw[0].map(normHdr_); // FIX: normalizar headers robustamente
  const idx = {};
  hdr.forEach((h, i) => { if(h) idx[h] = i; });

  const cID   = idx["ID"];
  const cSEDE = idx["SEDE"];
  const cNOM  = idx["NOMBRE_TRABAJADOR"];
  const cSTAT = idx[COL_STATUS_TRAB];
  const cTURN = (idx["JORNADA"] !== undefined) ? idx["JORNADA"]
              : (idx["TURNO"]   !== undefined) ? idx["TURNO"] : -1;
  const cDIA  = (idx["DIA_DESCANSO"] !== undefined) ? idx["DIA_DESCANSO"] : -1;

  if (cID === undefined || cSEDE === undefined || cNOM === undefined)
    throw new Error("CONTRATOS_2026 requiere columnas: ID, SEDE, NOMBRE_TRABAJADOR");

  const estadoPase = readEstadoPaseMap_();
  const out = [];

  for (let i = 1; i < raw.length; i++) {
    const r = raw[i];

    const rowSede = String(r[cSEDE] || "").trim();
    if (rowSede !== sedeN) continue;

    // FIX #2: normalizar ID (puede ser numero entero o float)
    const id = formatCellValue_(r[cID]);
    if (!id) continue;

    const st = cSTAT !== undefined
      ? String(r[cSTAT] || "ACTIVO").toUpperCase().trim()
      : "ACTIVO";
    if (st === "BAJA") continue;

    if (estadoPase[id] && estadoPase[id].estado === "BAJA") continue;

    const nombre = String(r[cNOM] || "").trim();
    const turno  = cTURN >= 0 ? String(r[cTURN] || "").trim() : "";

    var diaDescanso = cDIA >= 0 ? String(r[cDIA] || "").trim() : "";
    out.push({ id, nombre, turno, diaDescanso });
  }

  out.sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
  return out;
}

function apiGetPaseListaDia(fechaISO, sede) {
  const fecha = String(fechaISO || "").trim();
  const sedeN = String(sede     || "").trim();
  if (!isoValid_(fecha)) throw new Error("Fecha invalida (yyyy-MM-dd)");
  if (!sedeN) throw new Error("Falta sede");
  const abrev = abrevBySede_(sedeN);
  if (!abrev) throw new Error("La sede '" + sedeN + "' no tiene ABREV. Revisa CAT_SEDES o ejecuta Setup.");
  return readDayMarks_(fecha, abrev);
}

function apiGuardarPaseLista(payload) {
  const token    = String((payload && payload.token) || "");
  const ses      = requireSession_(token);
  const fechaISO = String((payload && payload.fecha) || "").trim();
  const sede     = String((payload && payload.sede)  || "").trim();
  const lista    = Array.isArray(payload && payload.lista) ? payload.lista : [];

  if (!isoValid_(fechaISO)) throw new Error("Fecha invalida");
  if (!sede) throw new Error("Falta sede");

  const perm = canEdit_(ses.role, fechaISO);
  if (!perm.ok) throw new Error(perm.reason);
  const lock = isLockedByNomina_(ses.role, fechaISO);
  if (lock.locked) throw new Error("Nomina cerrada: " + lock.reason);

  const abrev = abrevBySede_(sede);
  if (!abrev) throw new Error("La sede '" + sede + "' no tiene ABREV. Revisa CAT_SEDES.");

  const pl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  if (!pl) throw new Error("No existe PASE_LISTA_V2. Ejecuta Setup primero.");

  const existing = readDayRowsByKey_(fechaISO, abrev);
  const appended = [];
  const updated  = [];
  const q    = quincenaFromISO_(fechaISO);
  const capt = ses.username;
  const ts   = isoNow_();

  for (var i = 0; i < lista.length; i++) {
    var it     = lista[i];
    var id     = String(it.id     || "").trim();
    var nombre = String(it.nombre || "").trim();
    var codigo = String(it.codigo || "").trim().toUpperCase();
    if (!id || !codigo) continue;
    var key = makeKey_(fechaISO, abrev, id);
    if (existing[key]) {
      if (String(ses.role).toUpperCase() !== "SUPERADMIN")
        throw new Error("Registro ya asentado (ID " + id + "). Solo SUPERADMIN puede corregir.");
      updated.push({ row: existing[key].row, codigoNew: codigo });
    } else {
      appended.push({ key: key, id: id, nombre: nombre, codigo: codigo });
    }
  }

  for (var j = 0; j < updated.length; j++) {
    pl.getRange(updated[j].row, 7).setValue(updated[j].codigoNew);
    pl.getRange(updated[j].row, 8).setValue(capt);
    pl.getRange(updated[j].row, 9).setValue(ts);
  }
  if (appended.length) {
    var start = pl.getLastRow() + 1;
    var rows  = appended.map(function(a) {
      return [fechaISO, q, sede, abrev, a.id, a.nombre, a.codigo, capt, ts, a.key];
    });
    pl.getRange(start, 1, rows.length, HDR_PL.length).setValues(rows);
  }

  return { ok: true, appended: appended.length, updated: updated.length, user: capt };
}

// ══════════════════════════════════════════
//  LOGIN
// ══════════════════════════════════════════
function apiLogin_(body) {
  const u = String((body && body.username) || "").toLowerCase().trim();
  const p = String((body && body.password) || "");
  if (!u || !p) return { ok: false, message: "Falta usuario/contrasena" };
  const found = readUsers_().find(function(x) { return x.username === u && x.active; });
  if (!found)            return { ok: false, message: "Usuario no existe o inactivo" };
  if (found.password !== p) return { ok: false, message: "Contrasena incorrecta" };
  return { ok: true, token: issueToken_(u, found.role), role: found.role, username: u };
}
function apiLogout_(body) {
  const t = String((body && body.token) || "").trim();
  if (t) CacheService.getScriptCache().remove("T:" + t);
  return { ok: true };
}

// ══════════════════════════════════════════
//  REPORTE DIARIO
// ══════════════════════════════════════════
function apiReporteDiario(payload) {
  const token = String((payload && payload.token) || "");
  const ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede ver reportes.");

  const fecha  = String((payload && payload.fecha)  || "").trim();
  const codigo = String((payload && payload.codigo) || "ALL").trim().toUpperCase();
  if (!isoValid_(fecha)) throw new Error("Fecha invalida");

  var sedes = [];
  if (Array.isArray(payload && payload.sedes) && payload.sedes.length) {
    sedes = payload.sedes.map(function(s) { return String(s || "").trim(); }).filter(Boolean);
  } else if (payload && payload.sede) {
    sedes = [String(payload.sede).trim()].filter(Boolean);
  } else {
    sedes = getSedes_().map(function(x) { return x.sede; });
  }

  const pl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  if (!pl || pl.getLastRow() < 2) return { ok: true, fecha: fecha, items: [], counts: {}, totalAll: 0 };

  const v     = pl.getDataRange().getValues();
  const mapAb = {};
  getSedes_().forEach(function(s) { mapAb[s.sede] = String(s.abrev || "").trim().toUpperCase(); });
  const wanted = new Set(sedes.map(function(s) { return mapAb[s]; }).filter(Boolean));

  const counts = { A: 0, AF: 0, F: 0, PCG: 0, PSG: 0, I: 0, DT: 0, DS: 0, FER: 0, INH: 0 };
  const itemsAll = [];

  for (var i = 1; i < v.length; i++) {
    const f = formatCellValue_(v[i][0]); // FIX #6
    if (f !== fecha) continue;
    const abrev = String(v[i][3] || "").trim().toUpperCase();
    if (wanted.size && !wanted.has(abrev)) continue;
    const cod = String(v[i][6] || "").trim().toUpperCase();
    if (counts[cod] !== undefined) counts[cod]++;
    itemsAll.push({
      sede:   String(v[i][2] || "").trim(),
      abrev:  abrev,
      id:     formatCellValue_(v[i][4]),
      nombre: String(v[i][5] || "").trim(),
      codigo: cod
    });
  }

  const totalAll = itemsAll.length;
  const items    = (codigo === "ALL") ? itemsAll : itemsAll.filter(function(x) { return x.codigo === codigo; });
  return { ok: true, fecha: fecha, items: items, counts: counts, totalAll: totalAll };
}

// ══════════════════════════════════════════
//  HELPERS EXPORT / MATRICES / COLORES
// ══════════════════════════════════════════
function getCodeColor_(code) {
  var c = String(code || "").trim().toUpperCase();
  return CODE_COLORS_[c] || { bg: "#FFFFFF", fg: "#666666" };
}

function _htmlEsc_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function _uniqueKeepOrder_(arr) {
  var out = [];
  var seen = {};
  (arr || []).forEach(function(x) {
    var k = String(x || "").trim();
    if (!k || seen[k]) return;
    seen[k] = true;
    out.push(k);
  });
  return out;
}

function buildAttendanceMatrix_(payload) {
  var fechas = _uniqueKeepOrder_((payload && payload.fechas) || []).filter(isoValid_);
  if (!fechas.length) throw new Error("No hay fechas válidas para construir la matriz.");

  var sede = String((payload && payload.sede) || "").trim();
  var sedesList = sede ? [sede] : getSedes_().map(function(s) { return s.sede; });

  var empleados = [];
  sedesList.forEach(function(sd) {
    apiGetEmpleadosPorSede(sd).forEach(function(emp) {
      empleados.push({
        id: String(emp.id || "").trim(),
        nombre: String(emp.nombre || "").trim(),
        turno: String(emp.turno || "").trim(),
        sede: sd
      });
    });
  });

  empleados.sort(function(a, b) {
    if (a.sede !== b.sede) return a.sede.localeCompare(b.sede, "es");
    return a.nombre.localeCompare(b.nombre, "es");
  });

  var pl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  var marks = {};
  var counts = { A:0, AF:0, F:0, PCG:0, PSG:0, I:0, DT:0, DS:0, FER:0, INH:0 };

  if (pl && pl.getLastRow() >= 2) {
    var v = pl.getDataRange().getValues();
    var wantedSedes = {};
    sedesList.forEach(function(s) { wantedSedes[s] = true; });
    var wantedDates = {};
    fechas.forEach(function(f) { wantedDates[f] = true; });

    for (var i = 1; i < v.length; i++) {
      var f = formatCellValue_(v[i][0]);
      if (!wantedDates[f]) continue;
      var rowSede = String(v[i][2] || "").trim();
      if (!wantedSedes[rowSede]) continue;
      var id  = formatCellValue_(v[i][4]);
      var cod = String(v[i][6] || "").trim().toUpperCase();
      if (!id) continue;
      marks[rowSede + "|" + id + "|" + f] = cod;
      if (counts[cod] !== undefined) counts[cod]++;
    }
  }

  var headers = ["ID", "NOMBRE", "SEDE", "JORNADA"].concat(fechas);
  var rows = [];
  var codeMatrix = [];

  empleados.forEach(function(emp) {
    var row = [emp.id, emp.nombre, emp.sede, emp.turno || ""];
    var codes = [];
    fechas.forEach(function(fecha) {
      var cod = marks[emp.sede + "|" + emp.id + "|" + fecha] || "";
      row.push(cod);
      codes.push(cod);
    });
    rows.push(row);
    codeMatrix.push(codes);
  });

  return {
    headers: headers,
    rows: rows,
    fechas: fechas,
    empleados: empleados,
    codeMatrix: codeMatrix,
    sedes: sedesList,
    counts: counts
  };
}

function paintCodeMatrix_(sheet, startRow, startCol, matrix) {
  if (!matrix || !matrix.length || !matrix[0].length) return;
  var bgs = [], fgs = [];
  for (var r = 0; r < matrix.length; r++) {
    var bgRow = [], fgRow = [];
    for (var c = 0; c < matrix[r].length; c++) {
      var color = getCodeColor_(matrix[r][c]);
      bgRow.push(color.bg);
      fgRow.push(color.fg);
    }
    bgs.push(bgRow);
    fgs.push(fgRow);
  }
  var rng = sheet.getRange(startRow, startCol, matrix.length, matrix[0].length);
  rng.setBackgrounds(bgs);
  rng.setFontColors(fgs);
  rng.setHorizontalAlignment("center");
  rng.setFontWeight("bold");
}

function buildAttendancePdfHtml_(payload, matrixInfo, username) {
  var fechas = matrixInfo.fechas || [];
  var rows   = matrixInfo.rows || [];
  var counts = matrixInfo.counts || {};
  var sedeLabel = String((payload && payload.sede) || "").trim() || "Todas las sedes";
  var label = String((payload && payload.label) || (fechas[0] || ""));
  var tipo  = String((payload && payload.tipo) || "dia");
  var tipoLabel = tipo === "dia" ? "Reporte Diario"
                : tipo === "semana" ? "Reporte Semanal"
                : "Reporte Mensual";

  var chips = [];
  function chip(lbl, val, color) {
    return "<div class='chip'><span class='lbl'>" + _htmlEsc_(lbl) + "</span><span class='val' style='color:" + color + "'>" + val + "</span></div>";
  }
  chips.push(chip("Total empleados", rows.length, "#1a3a6c"));
  chips.push(chip("Asistencias", (counts.A||0)+(counts.AF||0), "#276221"));
  chips.push(chip("Faltas", counts.F||0, "#990000"));
  if (counts.PCG) chips.push(chip("PCG", counts.PCG, "#1155CC"));
  if (counts.PSG) chips.push(chip("PSG", counts.PSG, "#783F04"));
  if (counts.I)   chips.push(chip("Incapacidad", counts.I, "#7F6000"));
  if (counts.DS)  chips.push(chip("Descanso", counts.DS, "#4A1942"));

  var thead = "<tr><th>ID</th><th>Nombre</th><th>Sede</th><th>Jornada</th>";
  fechas.forEach(function(f) { thead += "<th>" + _htmlEsc_(f.substring(8)) + "</th>"; });
  thead += "</tr>";

  var tbody = "";
  rows.forEach(function(r) {
    tbody += "<tr>"
      + "<td>" + _htmlEsc_(r[0]) + "</td>"
      + "<td>" + _htmlEsc_(r[1]) + "</td>"
      + "<td>" + _htmlEsc_(r[2]) + "</td>"
      + "<td>" + _htmlEsc_(r[3]) + "</td>";
    for (var i = 4; i < r.length; i++) {
      var cod = String(r[i] || "").trim().toUpperCase();
      var color = getCodeColor_(cod);
      tbody += "<td style='background:" + color.bg + ";color:" + color.fg + ";font-weight:700;text-align:center'>" + _htmlEsc_(cod || "—") + "</td>";
    }
    tbody += "</tr>";
  });

  if (!tbody) {
    tbody = "<tr><td colspan='" + (4 + fechas.length) + "' style='text-align:center;color:#666;padding:16px'>Sin datos para el período seleccionado.</td></tr>";
  }

  return "<!DOCTYPE html><html><head><meta charset='UTF-8'>"
    + "<style>"
    + "*{box-sizing:border-box} body{font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#1a1a1a;padding:18px}"
    + "h1{font-size:18px;color:#1a3a6c;margin:0 0 4px 0} .sub{color:#555;margin-bottom:10px;font-size:11px}"
    + ".chips{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0 16px 0} .chip{border:1px solid #dde;padding:6px 10px;border-radius:999px;background:#f7f9ff}"
    + ".chip .lbl{font-size:9px;color:#666;margin-right:6px;text-transform:uppercase} .chip .val{font-size:14px;font-weight:700}"
    + "table{width:100%;border-collapse:collapse;font-size:9px} th,td{border:1px solid #d9e2f3;padding:4px 6px} th{background:#1a3a6c;color:#fff}"
    + "tbody tr:nth-child(even) td{background:#fafcff} .foot{margin-top:10px;color:#888;font-size:9px;text-align:right}"
    + "</style></head><body>"
    + "<h1>MHS Integradora</h1>"
    + "<div class='sub'>" + _htmlEsc_(tipoLabel) + " · Período: " + _htmlEsc_(label) + " · Sede: " + _htmlEsc_(sedeLabel) + " · Generado: " + _htmlEsc_(isoNow_()) + " · Por: " + _htmlEsc_(username || "") + "</div>"
    + "<div class='chips'>" + chips.join("") + "</div>"
    + "<table><thead>" + thead + "</thead><tbody>" + tbody + "</tbody></table>"
    + "<div class='foot'>MHS RH HUB 2026</div>"
    + "</body></html>";
}


// ══════════════════════════════════════════
//  EXPORTAR QUINCENA
// ══════════════════════════════════════════
function apiExportarQuincena(payload) {
  const fechaISO = String((payload && payload.fecha) || "").trim();
  const sede     = String((payload && payload.sede)  || "").trim();
  if (!isoValid_(fechaISO)) throw new Error("Fecha invalida");
  if (!sede) throw new Error("Falta sede");

  const q      = quincenaFromISO_(fechaISO);
  const parts  = fechaISO.split("-");
  const Y = parts[0], M = parts[1];
  const startDay = (q === "Q1") ? 1 : 16;
  const endDay   = (q === "Q1") ? 15 : daysInMonth_(Number(Y), Number(M));
  const fechas   = [];
  for (var d = startDay; d <= endDay; d++) {
    fechas.push(Y + "-" + M + "-" + String(d).padStart(2, "0"));
  }

  const matrix = buildAttendanceMatrix_({ fechas: fechas, sede: sede });

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const abrev = abrevBySede_(sede) || "SEDE";
  const name = "EXPORT_" + String(abrev).toUpperCase() + "_" + q + "_" + Y + M;
  var sh     = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name); else sh.clear();

  sh.getRange(1, 1).setValue("MHS Integradora · Exportación quincenal")
    .setFontWeight("bold").setFontSize(13).setFontColor("#1a3a6c");
  sh.getRange(1, 3).setValue("Sede: " + sede).setFontColor("#444444");
  sh.getRange(1, 5).setValue("Quincena: " + q + " · " + Y + "-" + M).setFontColor("#444444");
  sh.getRange(1, 8).setValue("Generado: " + isoNow_()).setFontColor("#888888");

  const header = matrix.headers;
  sh.getRange(2, 1, 1, header.length).setValues([header]);
  sh.getRange(2, 1, 1, header.length)
    .setBackground("#1a3a6c")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sh.setFrozenRows(2);

  if (matrix.rows.length) {
    sh.getRange(3, 1, matrix.rows.length, header.length).setValues(matrix.rows);
    paintCodeMatrix_(sh, 3, 5, matrix.codeMatrix);

    sh.getRange(3, 1, matrix.rows.length, 4)
      .setBackground("#ffffff")
      .setFontColor("#1a1a1a");
  } else {
    sh.getRange(3, 1).setValue("Sin datos para la quincena seleccionada.");
  }

  sh.autoResizeColumns(1, header.length);
  sh.setColumnWidth(2, 230);
  sh.setColumnWidth(3, 170);
  sh.setColumnWidth(4, 90);
  for (var c = 5; c <= header.length; c++) sh.setColumnWidth(c, 44);

  // Leyenda
  var legendRow = Math.max(5, matrix.rows.length + 5);
  sh.getRange(legendRow, 1).setValue("LEYENDA").setFontWeight("bold").setFontColor("#1a3a6c");
  var legend = [
    ["A","Asistió"],["AF","Asistió Feriado"],["F","Faltó"],["PCG","Permiso con goce"],
    ["PSG","Permiso sin goce"],["I","Incapacidad"],["DT","Doble turno"],
    ["DS","Descanso"],["FER","Feriado"],["INH","Inhabilitado"]
  ];
  legend.forEach(function(item, i) {
    var row = legendRow + 1 + i;
    var color = getCodeColor_(item[0]);
    sh.getRange(row, 1).setValue(item[0]).setBackground(color.bg).setFontColor(color.fg).setFontWeight("bold");
    sh.getRange(row, 2).setValue(item[1]);
  });

  return {
    ok: true,
    quincena: q,
    sheetName: name,
    start: fechas[0],
    end: fechas[fechas.length - 1],
    spreadsheetUrl: ss.getUrl() + "#gid=" + sh.getSheetId()
  };
}
function daysInMonth_(year, month) { return new Date(year, month, 0).getDate(); }

// ══════════════════════════════════════════
//  BAJA DE EMPLEADO
// ══════════════════════════════════════════
function apiDarBajaEmpleado(payload) {
  const token     = String((payload && payload.token)     || "");
  const ses       = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede dar bajas.");

  const id        = String((payload && payload.id)        || "").trim();
  const sede      = String((payload && payload.sede)      || "").trim();
  const fechaBaja = String((payload && payload.fechaBaja) || todayISO_()).trim();
  const motivo    = String((payload && payload.motivo)    || "").trim();

  if (!id)   throw new Error("Falta ID.");
  if (!sede) throw new Error("Falta SEDE.");
  if (!isoValid_(fechaBaja)) throw new Error("Fecha baja invalida (yyyy-MM-dd).");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_CONTRATOS);
  if (!sh) throw new Error("No existe " + SH_CONTRATOS + ".");

  ensureContratosRHColumns_();

  const data = sh.getDataRange().getValues();
  const hdr  = data[0].map(normHdr_); // FIX
  const idx  = {};
  hdr.forEach(function(h, i) { if(h) idx[h] = i; });

  const cID = idx["ID"], cSEDE = idx["SEDE"], cNOM = idx["NOMBRE_TRABAJADOR"],
        cST = idx[COL_STATUS_TRAB], cFB = idx[COL_FECHA_BAJA],
        cMB = idx[COL_MOTIVO_BAJA], cBP = idx[COL_BAJA_POR], cBT = idx[COL_BAJA_TS];

  if ([cID, cSEDE, cNOM, cST].some(function(x) { return x === undefined; }))
    throw new Error("CONTRATOS_2026 necesita: ID, SEDE, NOMBRE_TRABAJADOR, " + COL_STATUS_TRAB);

  var rowIndex = -1, nombre = "";
  for (var i = 1; i < data.length; i++) {
    // FIX #2: comparar ID normalizado
    if (formatCellValue_(data[i][cID]) === id && String(data[i][cSEDE] || "").trim() === sede) {
      rowIndex = i + 1;
      nombre   = String(data[i][cNOM] || "").trim();
      break;
    }
  }
  if (rowIndex < 0) throw new Error("Empleado no encontrado. ID='" + id + "' SEDE='" + sede + "'");

  const ts = isoNow_();
  sh.getRange(rowIndex, cST + 1).setValue("BAJA");
  if (cFB !== undefined) sh.getRange(rowIndex, cFB + 1).setValue(fechaBaja);
  if (cMB !== undefined) sh.getRange(rowIndex, cMB + 1).setValue(motivo);
  if (cBP !== undefined) sh.getRange(rowIndex, cBP + 1).setValue(ses.username || "");
  if (cBT !== undefined) sh.getRange(rowIndex, cBT + 1).setValue(ts);

  const shB = _ensureSheetIfMissing(ss, SH_BAJAS, [HDR_BAJAS]);
  shB.appendRow([fechaBaja, sede, id, nombre, motivo, ses.username || "", ts]);
  try { upsertEstadoPase_(id, "BAJA", fechaBaja, motivo, ses.username || ""); } catch (_) {}

  return { ok: true, message: "Baja aplicada: " + nombre + " (ID " + id + ") en " + sede + "." };
}

// ══════════════════════════════════════════
//  ALTA DE EMPLEADO
//  FIX #2: HORA_INICIO/HORA_FIN se guardan como texto, nunca como Date
// ══════════════════════════════════════════
function apiAltaEmpleado(payload) {
  const token = String((payload && payload.token) || "");
  const ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede dar altas.");

  const nombre     = String((payload && payload.nombre)     || "").trim().toUpperCase();
  const rfc        = String((payload && payload.rfc)        || "").trim().toUpperCase();
  const sede       = String((payload && payload.sede)       || "").trim();
  const jornada    = String((payload && payload.jornada)    || "").trim();
  // FIX #2: horas siempre como texto plano, nunca como Date
  const horaInicio = String((payload && payload.horaInicio) || "").trim();
  const horaFin    = String((payload && payload.horaFin)    || "").trim();
  const domicilio  = String((payload && payload.domicilio)  || "").trim();
  const cp         = String((payload && payload.cp)         || "").trim();
  const descanso   = String((payload && payload.descanso)   || "").trim().toUpperCase();

  if (!nombre) throw new Error("Nombre es obligatorio.");
  if (!sede)   throw new Error("Sede es obligatoria.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_CONTRATOS);
  if (!sh) throw new Error("No existe " + SH_CONTRATOS + ".");

  ensureContratosRHColumns_();

  const data = sh.getDataRange().getValues();
  const hdr  = data[0].map(normHdr_); // FIX
  const idx  = {};
  hdr.forEach(function(h, i) { if(h) idx[h] = i; });

  // ID maximo + 1
  const ids    = data.slice(1).map(function(r) { return Number(formatCellValue_(r[idx["ID"]]) || 0); }).filter(function(n) { return n > 0; });
  const nextID = (Math.max.apply(null, ids.concat([0])) + 1);

  // Folio de sede desde CAT_SEDES
  const shSedes = ss.getSheetByName(SH_CAT_SEDES);
  var folio = nextID;
  if (shSedes && shSedes.getLastRow() >= 2) {
    const vs = shSedes.getDataRange().getValues();
    for (var i = 1; i < vs.length; i++) {
      if (String(vs[i][0] || "").trim() === sede) {
        folio = Number(vs[i][2] || 0) + 1;
        shSedes.getRange(i + 1, 3).setValue(folio);
        break;
      }
    }
  }

  const abrev      = abrevBySede_(sede);
  const contratoID = abrev
    ? "MHS/" + abrev + String(folio).padStart(3, "0") + "/2026"
    : "MHS/" + nextID + "/2026";

  const newRow = hdr.map(function(h) {
    switch (h) {
      case "ID":                 return nextID;
      case "SEDE":               return sede;
      case "SEGMENTO_ORIGINAL":  return sede;
      case "CONTRATO_ID":        return contratoID;
      case "NOMBRE_TRABAJADOR":  return nombre;
      case "RFC":                return rfc;
      case "DOMICILIO_COMPLETO": return domicilio;
      case "CP":                 return cp ? (Number(cp) || cp) : "";
      // FIX #2: guardar como texto, nunca como objeto Date
      case "HORA_INICIO":        return horaInicio;
      case "HORA_FIN":           return horaFin;
      case "JORNADA":            return jornada;
      case "DIA_DESCANSO":       return descanso;
      case "STATUS_PDF":         return "PENDIENTE";
      case COL_STATUS_TRAB:      return "ACTIVO";
      default:                   return "";
    }
  });

  sh.appendRow(newRow);

  const shLog = ss.getSheetByName("LOG_CONTRATOS");
  if (shLog) shLog.appendRow([isoNow_(), "ALTA", nextID, sh.getLastRow(), contratoID, "Alta por " + ses.username]);

  return { ok: true, message: "Alta registrada: " + nombre + " (ID " + nextID + ", Contrato " + contratoID + ").", id: nextID, contratoID: contratoID };
}

// ══════════════════════════════════════════
//  INCIDENCIA MANUAL
// ══════════════════════════════════════════
function apiRegistrarIncidencia(payload) {
  const token  = String((payload && payload.token)  || "");
  const ses    = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede registrar incidencias.");

  const sede   = String((payload && payload.sede)   || "").trim();
  const id     = String((payload && payload.id)     || "").trim();
  const nombre = String((payload && payload.nombre) || "").trim();
  const fecha  = String((payload && payload.fecha)  || "").trim();
  const tipo   = String((payload && payload.tipo)   || "").trim().toUpperCase();
  const obs    = String((payload && payload.obs)    || "").trim();

  if (!sede || !id || !fecha || !tipo) throw new Error("Faltan campos: sede, id, fecha, tipo.");
  if (!isoValid_(fecha)) throw new Error("Fecha invalida (yyyy-MM-dd).");

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shInc = _ensureSheetIfMissing(ss, SH_INCIDENCIAS, [HDR_INC]);
  const ts    = isoNow_();
  shInc.appendRow([fecha, sede, id, nombre, tipo, obs, ses.username || "", ts]);

  // Sincronizar al PASE_LISTA_V2
  try {
    const abrev = abrevBySede_(sede);
    if (abrev) {
      const pl = ss.getSheetByName(SH_PL);
      if (pl) {
        const key      = makeKey_(fecha, abrev, id);
        const existing = readDayRowsByKey_(fecha, abrev);
        const q        = quincenaFromISO_(fecha);
        if (existing[key]) {
          pl.getRange(existing[key].row, 7).setValue(tipo);
          pl.getRange(existing[key].row, 8).setValue(ses.username || "");
          pl.getRange(existing[key].row, 9).setValue(ts);
        } else {
          pl.appendRow([fecha, q, sede, abrev, id, nombre, tipo, ses.username || "", ts, key]);
        }
      }
    }
  } catch (_) {}

  return { ok: true, message: "Incidencia " + tipo + " registrada para " + (nombre || id) + " el " + fecha + "." };
}

// ══════════════════════════════════════════
//  ANALIZAR FOTO (Claude Vision)
// ══════════════════════════════════════════
function apiAnalizarFoto(payload) {
  const token       = String((payload && payload.token)       || "");
  const imageBase64 = String((payload && payload.imageBase64) || "");
  const sede        = String((payload && payload.sede)        || "").trim();

  if (!imageBase64) throw new Error("Falta imagen (imageBase64).");
  requireSession_(token);

  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) {
    return {
      ok: true,
      text: "Para habilitar el analisis de fotos, configura CLAUDE_API_KEY en: Configuracion del proyecto > Propiedades del script.",
      id: null, nombre: null
    };
  }

  var contexto = "";
  try {
    if (sede) {
      const emps = apiGetEmpleadosPorSede(sede);
      if (emps.length) {
        contexto = "\n\nEmpleados activos en la sede '" + sede + "':\n" +
          emps.slice(0, 80).map(function(e) { return "ID:" + e.id + " | " + e.nombre; }).join("\n");
      }
    }
  } catch (_) {}

  const prompt = "Eres un asistente de RH de MHS Integradora. Analiza la imagen de una credencial o identificacion de empleado.\n" +
    "Extrae: nombre completo y numero de ID/empleado que aparezca en la imagen.\n" +
    contexto + "\n" +
    "Si detectas un nombre o numero que coincide con la lista, indica el ID y nombre exactos.\n" +
    "Responde UNICAMENTE con JSON sin backticks:\n" +
    '{"nombre":"NOMBRE EN MAYUSCULAS o null","id":"NUMERO o null","texto_visible":"texto detectado","confianza":"alta|media|baja"}';

  try {
    const resp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      headers: {
        "x-api-key":         apiKey,
        "anthropic-version": "2023-06-01",
        "content-type":      "application/json"
      },
      payload: JSON.stringify({
        model:      CLAUDE_MODEL,
        max_tokens: 512,
        messages: [{
          role:    "user",
          content: [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
            { type: "text",  text:  prompt }
          ]
        }]
      }),
      muteHttpExceptions: true
    });

    const data   = JSON.parse(resp.getContentText());
    const text   = (data.content || []).map(function(c) { return c.text || ""; }).join("").trim();
    var   parsed = null;
    try {
      const m = text.match(/\{[\s\S]*\}/);
      if (m) parsed = JSON.parse(m[0]);
    } catch (_) {}

    return {
      ok:     true,
      text:   parsed ? (parsed.texto_visible || text) + " (confianza: " + (parsed.confianza || "?") + ")" : text,
      id:     parsed && parsed.id     ? String(parsed.id).trim()     : null,
      nombre: parsed && parsed.nombre ? String(parsed.nombre).trim() : null
    };
  } catch (err) {
    return { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
}

// ══════════════════════════════════════════
//  EXPORTAR PDF A DRIVE
// ══════════════════════════════════════════
function apiExportarPDFReporte(payload) {
  var token  = String((payload && payload.token)  || "");
  var ses    = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede exportar PDFs.");

  var fechas = [];
  if (Array.isArray(payload && payload.fechas) && payload.fechas.length) {
    fechas = payload.fechas;
  } else {
    var f = String((payload && payload.fecha) || todayISO_()).trim();
    if (!isoValid_(f)) throw new Error("Fecha invalida.");
    fechas = [f];
  }

  var sede = String((payload && payload.sede) || "").trim();
  var tipo = String((payload && payload.tipo) || (fechas.length > 1 ? "semana" : "dia"));

  var matrix = buildAttendanceMatrix_({ fechas: fechas, sede: sede });
  var html = buildAttendancePdfHtml_(payload, matrix, ses.username || "");

  try {
    var folder   = DriveApp.getFolderById(DRIVE_PDF_FOLDER_ID);
    var label    = String((payload && payload.label) || fechas[0] || "reporte").replace(/ /g, "_").replace(/\//g, "-");
    var fileName = "Reporte_" + tipo.toUpperCase() + "_" + label + "_" + (sede || "TODAS").replace(/ /g, "_") + ".pdf";

    var docFile = DriveApp.createFile(
      Utilities.newBlob(html, MimeType.HTML, "temp.html")
    );
    var pdfBlob = DriveApp.getFileById(docFile.getId()).getAs(MimeType.PDF);
    pdfBlob.setName(fileName);
    var pdfFile = folder.createFile(pdfBlob);
    docFile.setTrashed(true);

    return { ok: true, message: fileName + " guardado en Drive.", url: pdfFile.getUrl() };
  } catch (err) {
    return { ok: false, error: "Error al crear PDF: " + (err.message || err) + ". Alternativa: usa Exportar Sheets." };
  }
}

// ── Helper: construir el HTML del reporte ──
function buildReportHTML_(labelRango, sedeLabel, tipo, now, username,
                           total, asist, faltas, counts, headerCols, tablaHtml, esFechaMult) {
  var tipoLabel = tipo === "dia" ? "Reporte Diario"
                : tipo === "semana" ? "Reporte Semanal"
                : "Reporte Mensual";

  var sumBoxes = ""
    + "<div class='sum-box'><div class='val'>" + total + "</div><div class='lbl'>Total registros</div></div>"
    + "<div class='sum-box' style='border-color:#bbf7d0'><div class='val' style='color:#16a34a'>" + asist + "</div><div class='lbl'>Asistencias</div></div>"
    + "<div class='sum-box' style='border-color:#fecaca'><div class='val' style='color:#dc2626'>" + faltas + "</div><div class='lbl'>Faltas</div></div>";
  if (counts.I)   sumBoxes += "<div class='sum-box'><div class='val' style='color:#d97706'>" + counts.I   + "</div><div class='lbl'>Incapacidad</div></div>";
  if (counts.PCG) sumBoxes += "<div class='sum-box'><div class='val' style='color:#2563eb'>" + counts.PCG + "</div><div class='lbl'>PCG</div></div>";
  if (counts.PSG) sumBoxes += "<div class='sum-box'><div class='val' style='color:#dc2626'>" + counts.PSG + "</div><div class='lbl'>PSG</div></div>";
  if (counts.DS)  sumBoxes += "<div class='sum-box'><div class='val' style='color:#0891b2'>" + counts.DS  + "</div><div class='lbl'>Descanso</div></div>";
  if (counts.DT)  sumBoxes += "<div class='sum-box'><div class='val' style='color:#7c3aed'>" + counts.DT  + "</div><div class='lbl'>Doble turno</div></div>";

  var emptyRow = "<tr><td colspan='6' style='text-align:center;color:#888;padding:20px'>Sin registros para este período.</td></tr>";

  return "<!DOCTYPE html><html><head><meta charset='UTF-8'>"
    + "<style>"
    + "*{box-sizing:border-box;margin:0;padding:0}"
    + "body{font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#1a1a1a;padding:24px}"
    + ".header{border-bottom:3px solid #1a3a6c;padding-bottom:14px;margin-bottom:18px;"
    + "  display:flex;align-items:flex-start;justify-content:space-between;gap:20px}"
    + ".header-left h1{font-size:18px;color:#1a3a6c;font-weight:800;margin-bottom:3px}"
    + ".header-left .tipo{font-size:12px;color:#1a3a6c;font-weight:600;letter-spacing:.5px;"
    + "  background:#e8f0ff;padding:2px 8px;border-radius:99px;display:inline-block}"
    + ".header-left .rango{font-size:11px;color:#444;margin-top:4px}"
    + ".header-left .sede{font-size:10px;color:#666;margin-top:2px}"
    + ".header-right{text-align:right;color:#666;font-size:10px;line-height:1.7;flex-shrink:0}"
    + ".summary{display:flex;gap:10px;margin-bottom:18px;flex-wrap:wrap}"
    + ".sum-box{border:1px solid #dde;padding:8px 14px;border-radius:8px;"
    + "  min-width:90px;text-align:center;background:#f7f9ff}"
    + ".sum-box .val{font-size:20px;font-weight:800;color:#1a3a6c}"
    + ".sum-box .lbl{font-size:9px;color:#777;text-transform:uppercase;letter-spacing:.5px;margin-top:1px}"
    + "table{width:100%;border-collapse:collapse;font-size:10.5px}"
    + "thead th{background:#1a3a6c;color:#fff;padding:7px 9px;text-align:left}"
    + "tbody td{border-bottom:1px solid #eef;padding:5px 9px;vertical-align:middle}"
    + "tbody tr:nth-child(even) td{background:#f7f9ff}"
    + ".footer{margin-top:18px;text-align:right;color:#aaa;font-size:9px;"
    + "  border-top:1px solid #eee;padding-top:8px}"
    + "@media print{body{padding:10px}.footer{position:fixed;bottom:10px;right:10px}}"
    + "</style></head><body>"
    + "<div class='header'>"
    + "  <div class='header-left'>"
    + "    <h1>MHS Integradora</h1>"
    + "    <div class='tipo'>" + tipoLabel + "</div>"
    + "    <div class='rango'>Período: " + labelRango + "</div>"
    + "    <div class='sede'>Sede: " + sedeLabel + "</div>"
    + "  </div>"
    + "  <div class='header-right'>"
    + "    Generado: " + now + "<br>"
    + "    Por: " + username + "<br>"
    + "    Sistema RH HUB 2026"
    + "  </div>"
    + "</div>"
    + "<div class='summary'>" + sumBoxes + "</div>"
    + "<table>"
    + "  <thead><tr>" + headerCols + "</tr></thead>"
    + "  <tbody>" + (tablaHtml || emptyRow) + "</tbody>"
    + "</table>"
    + "<div class='footer'>MHS Integradora &middot; Sistema RH HUB 2026 &middot; " + now + "</div>"
    + "</body></html>";
}

// ══════════════════════════════════════════
//  EXPORTAR A GOOGLE SHEETS (con colores)
// ══════════════════════════════════════════
// Colores por código
var CODE_COLORS_ = {
  A:   { bg: "#C6EFCE", fg: "#276221" },  // verde
  AF:  { bg: "#D9EAD3", fg: "#276221" },  // verde claro
  F:   { bg: "#F4CCCC", fg: "#990000" },  // rojo
  PSG: { bg: "#FCE5CD", fg: "#783F04" },  // naranja
  PCG: { bg: "#CFE2F3", fg: "#1155CC" },  // azul claro
  I:   { bg: "#FFF2CC", fg: "#7F6000" },  // amarillo
  DT:  { bg: "#D0E0E3", fg: "#134F5C" },  // azul gris
  DS:  { bg: "#EAD1DC", fg: "#4A1942" },  // rosa
  FER: { bg: "#D9D2E9", fg: "#20124D" },  // morado
  INH: { bg: "#EFEFEF", fg: "#666666" },  // gris
};

function apiExportarSheets(payload) {
  var token  = String((payload && payload.token)  || "");
  var ses    = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede exportar Sheets.");

  var fechas = [];
  if (Array.isArray(payload && payload.fechas) && payload.fechas.length) {
    fechas = payload.fechas;
  } else {
    var f = String((payload && payload.fecha) || todayISO_()).trim();
    fechas = [f];
  }
  fechas = _uniqueKeepOrder_(fechas).filter(isoValid_);
  if (!fechas.length) throw new Error("No hay fechas válidas para exportar.");

  var tipo  = String((payload && payload.tipo)  || (fechas.length > 1 ? "semana" : "dia"));
  var sede  = String((payload && payload.sede)  || "").trim();
  var label = String((payload && payload.label) || fechas[0]).trim();

  var matrix = buildAttendanceMatrix_({ fechas: fechas, sede: sede });

  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var tipoLabel = tipo === "dia" ? "DIA" : (tipo === "semana" ? "SEM" : "MES");
  var shName = "RPT_" + tipoLabel + "_" + label.replace(/ /g,"-").replace(/\//g,"-").substring(0,20);
  var existingSheet = ss.getSheetByName(shName);
  if (existingSheet) ss.deleteSheet(existingSheet);
  var sh = ss.insertSheet(shName);

  sh.getRange(1,1).setValue("MHS Integradora - Reporte de Asistencias")
    .setFontSize(13).setFontWeight("bold").setFontColor("#1a3a6c");
  sh.getRange(1,2).setValue(tipo==="dia"?"Reporte Diario":tipo==="semana"?"Reporte Semanal":"Reporte Mensual")
    .setFontColor("#1a3a6c").setFontWeight("bold");
  sh.getRange(1,4).setValue("Período: " + label).setFontColor("#444444");
  sh.getRange(1,5).setValue("Sede: " + (sede||"Todas")).setFontColor("#444444");
  sh.getRange(1,6).setValue("Generado: " + isoNow_()).setFontColor("#888888");

  var headers = matrix.headers.concat(["TOTAL A", "TOTAL F", "TOTAL I"]);
  sh.getRange(2, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 1, 1, headers.length)
    .setBackground("#1a3a6c")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var dataRows = [];
  var codeMatrix = [];
  matrix.rows.forEach(function(r, idx) {
    var codes = matrix.codeMatrix[idx] || [];
    var cntA=0, cntF=0, cntI=0;
    codes.forEach(function(c) {
      if (c==="A"||c==="AF") cntA++;
      if (c==="F") cntF++;
      if (c==="I") cntI++;
    });
    dataRows.push(r.concat([cntA, cntF, cntI]));
    codeMatrix.push(codes);
  });

  if (dataRows.length) {
    sh.getRange(3, 1, dataRows.length, headers.length).setValues(dataRows);
    paintCodeMatrix_(sh, 3, 5, codeMatrix);

    sh.getRange(3, 1, dataRows.length, 4).setBackground("#ffffff").setFontColor("#1a1a1a");
    sh.getRange(3, headers.length - 2, dataRows.length, 3).setHorizontalAlignment("center").setFontWeight("bold");
  } else {
    sh.getRange(3,1).setValue("Sin datos para este período.");
  }

  // Totales coloreados
  for (var r = 0; r < dataRows.length; r++) {
    var row = dataRows[r];
    var a = row[headers.length - 3], f = row[headers.length - 2], i = row[headers.length - 1];
    if (a > 0) sh.getRange(3+r, headers.length - 2).setBackground("#C6EFCE").setFontColor("#276221");
    if (f > 0) sh.getRange(3+r, headers.length - 1).setBackground("#F4CCCC").setFontColor("#990000");
    if (i > 0) sh.getRange(3+r, headers.length).setBackground("#FFF2CC").setFontColor("#7F6000");
  }

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 60);
  sh.setColumnWidth(2, 230);
  sh.setColumnWidth(3, 170);
  sh.setColumnWidth(4, 90);
  for (var c = 5; c < 5 + matrix.fechas.length; c++) sh.setColumnWidth(c, 44);
  sh.autoResizeColumns(Math.max(5, headers.length - 2), 3);

  // Leyenda
  var legendStartRow = Math.max(5, dataRows.length + 5);
  sh.getRange(legendStartRow, 1).setValue("LEYENDA DE COLORES").setFontWeight("bold").setFontColor("#1a3a6c");
  var legendItems = [
    ["A",  "Asistió"],["AF", "Asistió (Feriado)"],["F", "Faltó"],["PCG","Permiso con goce"],
    ["PSG","Permiso sin goce"],["I","Incapacidad"],["DT","Doble turno"],["DS","Descanso"],["FER","Feriado"],["INH","Inhabilitado"]
  ];
  legendItems.forEach(function(item, idx) {
    var lr = legendStartRow + 1 + idx;
    var color = getCodeColor_(item[0]);
    sh.getRange(lr, 1).setValue(item[0]).setBackground(color.bg).setFontColor(color.fg).setFontWeight("bold").setHorizontalAlignment("center");
    sh.getRange(lr, 2).setValue(item[1]).setFontColor("#333333");
  });

  return { ok: true, message: "Hoja '" + shName + "' creada.", url: ss.getUrl() + "#gid=" + sh.getSheetId() };
}



// ══════════════════════════════════════════
// ══════════════════════════════════════════
//  CAMBIAR DIA DE DESCANSO FIJO
//  Modifica DIA_DESCANSO en CONTRATOS_2026
//  y registra el cambio en hoja CAMBIOS_DESCANSO
// ══════════════════════════════════════════
function apiCambiarDiaDescanso(payload) {
  var token    = String((payload && payload.token)    || "");
  var ses      = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede cambiar dias de descanso.");

  var id       = String((payload && payload.id)       || "").trim();
  var sede     = String((payload && payload.sede)     || "").trim();
  var diaNuevo = String((payload && payload.diaNuevo) || "").trim().toUpperCase();
  var motivo   = String((payload && payload.motivo)   || "").trim();

  if (!id)       throw new Error("Falta ID del trabajador.");
  if (!sede)     throw new Error("Falta SEDE.");
  if (!diaNuevo) throw new Error("Falta el nuevo dia de descanso.");

  var diasValidos = ["LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO","DOMINGO","SABADO Y DOMINGO"];
  if (diasValidos.indexOf(diaNuevo) < 0)
    throw new Error("Dia invalido: " + diaNuevo);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CONTRATOS);
  if (!sh) throw new Error("No existe " + SH_CONTRATOS + ".");

  var data = sh.getDataRange().getValues();
  // FIX: usar normHdr_ para tolerar espacios ocultos, saltos de linea, etc.
  var hdr  = data[0].map(normHdr_);
  var idx  = {};
  hdr.forEach(function(h,i){ if(h) idx[h]=i; });

  var cID   = idx["ID"];
  var cSEDE = idx["SEDE"];
  var cNOM  = idx["NOMBRE_TRABAJADOR"];
  var cDESC = idx["DIA_DESCANSO"];

  if (cID===undefined || cSEDE===undefined || cNOM===undefined)
    throw new Error("CONTRATOS_2026 requiere columnas: ID, SEDE, NOMBRE_TRABAJADOR.");
  if (cDESC === undefined) {
    // Diagnóstico: mostrar los headers reales para debugging
    var hdrsStr = hdr.filter(function(h){ return h; }).join(" | ");
    throw new Error(
      "No se encontro columna DIA_DESCANSO. Cabeceras detectadas: [" + hdrsStr + "]. " +
      "Verifica que la columna se llame exactamente DIA_DESCANSO en la hoja CONTRATOS_2026."
    );
  }

  // Buscar la fila del empleado
  var rowIndex = -1, nombre = "", diaAnterior = "";
  for (var i = 1; i < data.length; i++) {
    if (formatCellValue_(data[i][cID]) === id && String(data[i][cSEDE]||"").trim() === sede) {
      rowIndex    = i + 1;
      nombre      = String(data[i][cNOM]  || "").trim();
      diaAnterior = String(data[i][cDESC] || "").trim();
      break;
    }
  }
  if (rowIndex < 0)
    throw new Error("Trabajador no encontrado. ID=" + id + " SEDE=" + sede);

  // Actualizar DIA_DESCANSO en la fila
  sh.getRange(rowIndex, cDESC + 1).setValue(diaNuevo);

  // Registrar el cambio en hoja CAMBIOS_DESCANSO (la crea si no existe)
  var SH_CAMBIOS = "CAMBIOS_DESCANSO";
  var shC = ss.getSheetByName(SH_CAMBIOS);
  if (!shC) {
    shC = ss.insertSheet(SH_CAMBIOS);
    shC.getRange(1,1,1,7).setValues([["TIMESTAMP","SEDE","ID","NOMBRE","DIA_ANTERIOR","DIA_NUEVO","MOTIVO","CAPTURADO_POR"]]);
    shC.setFrozenRows(1);
    shC.autoResizeColumns(1,8);
  }
  var ts = isoNow_();
  shC.appendRow([ts, sede, id, nombre, diaAnterior, diaNuevo, motivo, ses.username||""]);

  return {
    ok: true,
    message: "Dia de descanso actualizado: " + nombre + " (" + id + ") "
           + diaAnterior + " -> " + diaNuevo + "."
  };
}

// ══════════════════════════════════════════
//  SINCRONIZAR CONTRATOS_2026
//  Asegura que todas las columnas necesarias
//  existen y tienen los valores correctos.
//  Ejecutar desde el menú cuando hay problemas.
// ══════════════════════════════════════════

// Wrapper para el menú
function menuSyncContratos() {
  ensureContratosRHColumns_();
  fixDiaDescansoColumn_();
  SpreadsheetApp.getUi().alert(
    "Sincronizacion completada.\n\n" +
    "Se verificaron y corrigieron:\n" +
    "- Columna DIA_DESCANSO\n" +
    "- Columna STATUS_TRABAJADOR\n" +
    "- Columnas FECHA_BAJA, MOTIVO_BAJA, BAJA_CAPTURADO_POR, BAJA_TS\n\n" +
    "Ejecuta el Diagnostico rapido para confirmar."
  );
}

// Asegura que DIA_DESCANSO tiene datos válidos
// y que la columna está bien posicionada.
function fixDiaDescansoColumn_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_CONTRATOS);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol === 0) return;

  const allVals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const hdr     = allVals[0].map(normHdr_);
  const idx     = {};
  hdr.forEach(function(h, i) { if (h) idx[h] = i; });

  const cDESC = idx["DIA_DESCANSO"];

  // Si no existe la columna, agregarla ya lo hizo ensureContratosRHColumns_
  // Aquí solo normalizamos los valores existentes (typos, espacios extra)
  if (cDESC === undefined) return;

  const VALID_DIAS = {
    "LUNES":1,"MARTES":1,"MIERCOLES":1,"JUEVES":1,
    "VIERNES":1,"SABADO":1,"DOMINGO":1,
    "SABADO Y DOMINGO":1,"SABADO Y DOMNINGO":1
  };

  const col    = cDESC + 1;
  const rng    = sh.getRange(2, col, lastRow - 1, 1);
  const vals   = rng.getValues();
  let changed  = false;

  for (let i = 0; i < vals.length; i++) {
    var raw = String(vals[i][0] || "").trim();
    if (!raw || raw === "__________") continue;

    // Normalizar typo SABADO Y DOMNINGO → SABADO Y DOMINGO
    var norm = raw.toUpperCase()
      .replace("SABADO Y DOMNINGO", "SABADO Y DOMINGO")
      .replace("SABADO Y  DOMINGO", "SABADO Y DOMINGO")
      .trim();

    if (norm !== raw) {
      vals[i][0] = norm;
      changed = true;
    }
  }

  if (changed) rng.setValues(vals);
}

// Endpoint que el frontend puede llamar para diagnóstico de columnas
function apiDiagContratos(payload) {
  var token = String((payload && payload.token) || "");
  requireSession_(token);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CONTRATOS);
  if (!sh) return { ok: false, error: "No existe " + SH_CONTRATOS };

  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(normHdr_);
  var idx = {};
  hdr.forEach(function(h,i){ if(h) idx[h]=i; });

  return {
    ok: true,
    columnas: hdr.filter(Boolean),
    tiene_DIA_DESCANSO: idx["DIA_DESCANSO"] !== undefined,
    tiene_STATUS_TRABAJADOR: idx[COL_STATUS_TRAB] !== undefined,
    col_DIA_DESCANSO: idx["DIA_DESCANSO"] !== undefined ? idx["DIA_DESCANSO"] + 1 : null,
    total_filas: sh.getLastRow() - 1
  };
}


// ══════════════════════════════════════════
function apiLiberarFecha(payload) {
  var token    = String((payload && payload.token)   || "");
  var ses      = requireSession_(token);
  if (String(ses.role).toUpperCase() !== "SUPERADMIN")
    throw new Error("Solo SUPERADMIN puede liberar fechas.");

  var fechaISO = String((payload && payload.fecha)   || "").trim();
  var liberar  = (payload && payload.liberar !== false); // default true

  if (!isoValid_(fechaISO)) throw new Error("Fecha invalida (yyyy-MM-dd).");

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sh    = ss.getSheetByName(SH_LOCK);
  if (!sh) { sh = ss.insertSheet(SH_LOCK); }

  var clave = "LIB_" + fechaISO;
  var v     = sh.getDataRange().getValues();
  var found = false;
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0] || "").trim() === clave) {
      sh.getRange(i + 1, 2).setValue(liberar ? "1" : "0");
      found = true;
      break;
    }
  }
  if (!found && liberar) {
    sh.appendRow([clave, "1", isoNow_(), ses.username]);
  }

  return {
    ok: true,
    fecha: fechaISO,
    liberada: liberar,
    message: liberar
      ? "Fecha " + fechaISO + " liberada. Los supervisores pueden asentar asistencias."
      : "Liberacion de " + fechaISO + " revocada."
  };
}

// ══════════════════════════════════════════
function apiRegistrarIncidenciaFormal(payload) {
  var token       = String((payload && payload.token)        || "");
  var ses         = requireSession_(token);

  var sede        = String((payload && payload.sede)         || "").trim();
  var id          = String((payload && payload.id)           || "").trim();
  var nombre      = String((payload && payload.nombre)       || "").trim();
  var fecha       = String((payload && payload.fecha)        || "").trim();
  var tipo        = String((payload && payload.tipo)         || "").trim().toUpperCase();
  var obs         = String((payload && payload.obs)          || "").trim();
  if (obs.length > 800) obs = obs.substring(0, 800);
  var cubreId     = String((payload && payload.cubre_id)     || "").trim();
  var cubreNombre = String((payload && payload.cubre_nombre) || "").trim();
  var autoriza    = String((payload && payload.autoriza)     || "").trim();

  if (!sede || !id || !fecha || !tipo)
    throw new Error("Faltan campos requeridos: sede, id, fecha, tipo.");
  if (!isoValid_(fecha))
    throw new Error("Fecha invalida (yyyy-MM-dd).");

  var tiposConAutorizacion = ["PCG","PSG","DT"];
  if (tiposConAutorizacion.indexOf(tipo) >= 0 && !autoriza)
    throw new Error("Se requiere el responsable que autoriza para el tipo " + tipo + ".");

  var ss    = SpreadsheetApp.getActiveSpreadsheet();

  // ── Asegurar que la hoja INCIDENCIAS tiene las columnas v2 ──
  // Se hace aquí dentro para no depender de const globales
  var shInc = ss.getSheetByName(SH_INCIDENCIAS);
  if (!shInc) {
    shInc = ss.insertSheet(SH_INCIDENCIAS);
    shInc.getRange(1, 1, 1, 11).setValues([[
      "FECHA_ISO","SEDE","ID","NOMBRE","TIPO","OBSERVACIONES",
      "CUBRE_ID","CUBRE_NOMBRE","AUTORIZA","CAPTURADO_POR","TS_ISO"
    ]]);
    shInc.setFrozenRows(1);
    shInc.autoResizeColumns(1, 11);
  } else {
    // Agregar columnas nuevas si faltan
    var curHdr = shInc.getRange(1, 1, 1, shInc.getLastColumn()).getValues()[0].map(normHdr_);
    var curSet = {};
    curHdr.forEach(function(h) { if (h) curSet[h] = true; });
    var newCols = [];
    if (!curSet["CUBRE_ID"])     newCols.push("CUBRE_ID");
    if (!curSet["CUBRE_NOMBRE"]) newCols.push("CUBRE_NOMBRE");
    if (!curSet["AUTORIZA"])     newCols.push("AUTORIZA");
    if (newCols.length) {
      shInc.getRange(1, shInc.getLastColumn() + 1, 1, newCols.length).setValues([newCols]);
    }
  }

  // ── Leer headers actualizados y construir fila ──
  var hdrNow = shInc.getRange(1, 1, 1, shInc.getLastColumn()).getValues()[0].map(normHdr_);
  var hdrIdx = {};
  hdrNow.forEach(function(h, i) { if (h) hdrIdx[h] = i + 1; });

  var ts      = isoNow_();
  var numCols = shInc.getLastColumn();
  var newRow  = [];
  for (var ci = 0; ci < numCols; ci++) newRow.push("");

  function setC(name, val) {
    var idx = hdrIdx[name];
    if (idx) newRow[idx - 1] = val;
  }
  setC("FECHA_ISO",     fecha);
  setC("SEDE",          sede);
  setC("ID",            id);
  setC("NOMBRE",        nombre);
  setC("TIPO",          tipo);
  setC("OBSERVACIONES", obs);
  setC("CUBRE_ID",      cubreId);
  setC("CUBRE_NOMBRE",  cubreNombre);
  setC("AUTORIZA",      autoriza);
  setC("CAPTURADO_POR", ses.username || "");
  setC("TS_ISO",        ts);

  shInc.appendRow(newRow);

  // ── Sincronizar al PASE_LISTA_V2 ──
  try {
    var abrev = abrevBySede_(sede);
    if (abrev) {
      var pl = ss.getSheetByName(SH_PL);
      if (pl) {
        var key      = makeKey_(fecha, abrev, id);
        var existing = readDayRowsByKey_(fecha, abrev);
        var q        = quincenaFromISO_(fecha);
        if (existing[key]) {
          pl.getRange(existing[key].row, 7).setValue(tipo);
          pl.getRange(existing[key].row, 8).setValue(ses.username || "");
          pl.getRange(existing[key].row, 9).setValue(ts);
        } else {
          pl.appendRow([fecha, q, sede, abrev, id, nombre, tipo, ses.username || "", ts, key]);
        }
      }
    }
  } catch (e_) {}

  return {
    ok: true,
    message: "Incidencia " + tipo + " registrada para " + (nombre || id) + " el " + fecha + "."
      + (cubreNombre ? " Cubre: " + cubreNombre + "." : "")
      + (autoriza    ? " Autoriza: " + autoriza + "."  : "")
  };
}

// ══════════════════════════════════════════

function getTurnosAsIncidencias_(opts) {
  opts = opts || {};
  var sede = String(opts.sede || "").trim();
  var fechaInicio = String(opts.fechaInicio || "").trim();
  var fechaFin = String(opts.fechaFin || "").trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_TURNOS || "TURNOS_EVENTUALES");
  if (!sh || sh.getLastRow() < 2) return [];

  var v = sh.getDataRange().getValues();
  var hdr = v[0].map(normHdr_);
  var idx = {};
  hdr.forEach(function(h, i) { if (h) idx[h] = i; });

  function sg(row, col) {
    return col !== undefined ? String(row[col] || "").trim() : "";
  }

  var out = [];
  for (var i = 1; i < v.length; i++) {
    var rowFecha = formatCellValue_(v[i][idx["FECHA_ISO"] !== undefined ? idx["FECHA_ISO"] : 0]);
    var rowSede  = sg(v[i], idx["SEDE"]);

    if (sede && rowSede !== sede) continue;
    if (fechaInicio && rowFecha < fechaInicio) continue;
    if (fechaFin && rowFecha > fechaFin) continue;

    out.push({
      fecha:       rowFecha,
      id:          formatCellValue_(v[i][idx["ID"] !== undefined ? idx["ID"] : 2]),
      nombre:      sg(v[i], idx["NOMBRE"]),
      tipo:        "DT",
      obs:         sg(v[i], idx["OBS"]),
      cubreId:     "",
      cubreNombre: sg(v[i], idx["CUBRE_A"]),
      autoriza:    sg(v[i], idx["AUTORIZA"]),
      captPor:     sg(v[i], idx["CAPTURADO_POR"]),
      origen:      "TURNOS_EVENTUALES",
      turno:       sg(v[i], idx["TURNO"]),
      externo:     sg(v[i], idx["ES_EXTERNO"])
    });
  }

  return out;
}

function mergeIncidenciasConTurnos_(incRows, turnRows) {
  var a = Array.isArray(incRows) ? incRows.slice() : [];
  var b = Array.isArray(turnRows) ? turnRows.slice() : [];
  var seen = {};
  var out = [];

  function keyOf_(x) {
    return [
      String(x.fecha || "").trim(),
      String(x.sede || "").trim(),
      String(x.id || "").trim(),
      String(x.tipo || "").trim().toUpperCase(),
      String(x.origen || "INCIDENCIAS").trim()
    ].join("|");
  }

  a.forEach(function(x) {
    if (!x.origen) x.origen = "INCIDENCIAS";
    var k = keyOf_(x);
    if (!seen[k]) {
      seen[k] = true;
      out.push(x);
    }
  });

  b.forEach(function(x) {
    var k = keyOf_(x);
    if (!seen[k]) {
      seen[k] = true;
      out.push(x);
    }
  });

  out.sort(function(x, y) {
    if (x.fecha < y.fecha) return -1;
    if (x.fecha > y.fecha) return 1;
    return String(x.nombre || "").localeCompare(String(y.nombre || ""), "es");
  });

  return out;
}

function apiGetIncidenciasSede(payload) {
  var token       = String((payload && payload.token)       || "");
  requireSession_(token);

  var sede        = String((payload && payload.sede)        || "").trim();
  var fechaInicio = String((payload && payload.fechaInicio) || todayISO_()).trim();
  var fechaFin    = String((payload && payload.fechaFin)    || todayISO_()).trim();

  if (!sede)                   throw new Error("Falta sede.");
  if (!isoValid_(fechaInicio)) throw new Error("fechaInicio invalida.");
  if (!isoValid_(fechaFin))    throw new Error("fechaFin invalida.");

  var TIPOS_VALIDOS = { F:1, PCG:1, PSG:1, I:1, DS:1, DT:1, FER:1, INH:1 };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_INCIDENCIAS);
  var incRows = [];

  if (sh && sh.getLastRow() >= 2) {
    var v   = sh.getDataRange().getValues();
    var hdr = v[0].map(normHdr_);
    var idx = {};
    hdr.forEach(function(h, i) { if (h) idx[h] = i; });

    function safeGet(row, col) {
      return col !== undefined ? String(row[col] || "").trim() : "";
    }

    for (var i = 1; i < v.length; i++) {
      var rowSede  = safeGet(v[i], idx["SEDE"]);
      var rowFecha = formatCellValue_(v[i][idx["FECHA_ISO"] !== undefined ? idx["FECHA_ISO"] : 0]);
      var tipo     = safeGet(v[i], idx["TIPO"]).toUpperCase();

      if (rowSede !== sede) continue;
      if (rowFecha < fechaInicio || rowFecha > fechaFin) continue;
      if (!TIPOS_VALIDOS[tipo]) continue;

      incRows.push({
        fecha:       rowFecha,
        sede:        rowSede,
        id:          formatCellValue_(v[i][idx["ID"] !== undefined ? idx["ID"] : 2]),
        nombre:      safeGet(v[i], idx["NOMBRE"]),
        tipo:        tipo,
        obs:         safeGet(v[i], idx["OBSERVACIONES"]),
        cubreId:     safeGet(v[i], idx["CUBRE_ID"]),
        cubreNombre: safeGet(v[i], idx["CUBRE_NOMBRE"]),
        autoriza:    safeGet(v[i], idx["AUTORIZA"]),
        captPor:     safeGet(v[i], idx["CAPTURADO_POR"]),
        tsISO:       safeGet(v[i], idx["TS_ISO"]),
        origen:      "INCIDENCIAS"
      });
    }
  }

  var turnRows = getTurnosAsIncidencias_({
    sede: sede,
    fechaInicio: fechaInicio,
    fechaFin: fechaFin
  });

  var merged = mergeIncidenciasConTurnos_(incRows, turnRows);

  return { ok: true, items: merged };
}

function apiGetIncidenciasCalendario(payload) {
  var token = String((payload && payload.token) || "");
  requireSession_(token);

  var month = String((payload && payload.month) || "").trim(); // yyyy-MM
  var sede  = String((payload && payload.sede)  || "").trim();

  if (!/^\d{4}-\d{2}$/.test(month)) {
    throw new Error("Mes inválido. Usa formato yyyy-MM.");
  }

  var y = Number(month.substring(0, 4));
  var m = Number(month.substring(5, 7));
  var fechaInicio = month + "-01";
  var fechaFin = month + "-" + String(daysInMonth_(y, m)).padStart(2, "0");

  var TIPOS_VALIDOS = { F:1, PCG:1, PSG:1, I:1, DS:1, DT:1, FER:1, INH:1 };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_INCIDENCIAS);
  var incRows = [];

  if (sh && sh.getLastRow() >= 2) {
    var v   = sh.getDataRange().getValues();
    var hdr = v[0].map(normHdr_);
    var idx = {};
    hdr.forEach(function(h, i) { if (h) idx[h] = i; });

    function safeGet(row, col) {
      return col !== undefined ? String(row[col] || "").trim() : "";
    }

    for (var i = 1; i < v.length; i++) {
      var rowFecha = formatCellValue_(v[i][idx["FECHA_ISO"] !== undefined ? idx["FECHA_ISO"] : 0]);
      var rowSede  = safeGet(v[i], idx["SEDE"]);
      var tipo     = safeGet(v[i], idx["TIPO"]).toUpperCase();

      if (rowFecha < fechaInicio || rowFecha > fechaFin) continue;
      if (sede && sede !== "Todas las sedes" && rowSede !== sede) continue;
      if (!TIPOS_VALIDOS[tipo]) continue;

      incRows.push({
        fecha:       rowFecha,
        sede:        rowSede,
        id:          formatCellValue_(v[i][idx["ID"] !== undefined ? idx["ID"] : 2]),
        nombre:      safeGet(v[i], idx["NOMBRE"]),
        tipo:        tipo,
        obs:         safeGet(v[i], idx["OBSERVACIONES"]),
        cubreId:     safeGet(v[i], idx["CUBRE_ID"]),
        cubreNombre: safeGet(v[i], idx["CUBRE_NOMBRE"]),
        autoriza:    safeGet(v[i], idx["AUTORIZA"]),
        captPor:     safeGet(v[i], idx["CAPTURADO_POR"]),
        tsISO:       safeGet(v[i], idx["TS_ISO"]),
        origen:      "INCIDENCIAS"
      });
    }
  }

  var turnRows = getTurnosAsIncidencias_({
    sede: (sede && sede !== "Todas las sedes") ? sede : "",
    fechaInicio: fechaInicio,
    fechaFin: fechaFin
  });

  var merged = mergeIncidenciasConTurnos_(incRows, turnRows);

  return { ok: true, items: merged };
}


// ══════════════════════════════════════════
var SH_TURNOS = "TURNOS_EVENTUALES";
var HDR_TURNOS = [
  "FECHA_ISO","NOMBRE","ID","SEDE","TURNO",
  "CUBRE_A","OBS","AUTORIZA","ES_EXTERNO",
  "CAPTURADO_POR","TS_ISO"
];

function _ensureTurnosSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_TURNOS);
  if (!sh) {
    sh = ss.insertSheet(SH_TURNOS);
    sh.getRange(1, 1, 1, HDR_TURNOS.length).setValues([HDR_TURNOS]);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, HDR_TURNOS.length);
  }
  return sh;
}

function apiRegistrarTurnoEventual(payload) {
  var token     = String((payload && payload.token)     || "");
  var ses       = requireSession_(token);

  var fecha     = String((payload && payload.fecha)     || "").trim();
  var nombre    = String((payload && payload.nombre)    || "").trim().toUpperCase();
  var id        = String((payload && payload.id)        || "").trim();
  var sede      = String((payload && payload.sede)      || "").trim();
  var turno     = String((payload && payload.turno)     || "MATUTINO").trim().toUpperCase();
  var cubreA    = String((payload && payload.cubre_a)   || "").trim();
  var obs       = String((payload && payload.obs)       || "").trim();
  if (obs.length > 800) obs = obs.substring(0, 800);
  var autoriza  = String((payload && payload.autoriza)  || "").trim();
  var esExterno = (payload && payload.es_externo) ? "SI" : "NO";

  if (!fecha)    throw new Error("Falta la fecha del turno.");
  if (!nombre)   throw new Error("Falta el nombre de la persona.");
  if (!sede)     throw new Error("Falta la sede.");
  if (!autoriza) throw new Error("Falta el responsable que autoriza.");
  if (!isoValid_(fecha)) throw new Error("Fecha invalida (yyyy-MM-dd).");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = _ensureTurnosSheet_();
  var ts = isoNow_();

  // 1) Guarda en TURNOS_EVENTUALES
  sh.appendRow([fecha, nombre, id, sede, turno, cubreA, obs, autoriza, esExterno, ses.username || "", ts]);

  // 2) Espejo en INCIDENCIAS como DT
  var shInc = ss.getSheetByName(SH_INCIDENCIAS);
  if (!shInc) {
    shInc = ss.insertSheet(SH_INCIDENCIAS);
    shInc.getRange(1, 1, 1, 11).setValues([[
      "FECHA_ISO","SEDE","ID","NOMBRE","TIPO","OBSERVACIONES",
      "CUBRE_ID","CUBRE_NOMBRE","AUTORIZA","CAPTURADO_POR","TS_ISO"
    ]]);
    shInc.setFrozenRows(1);
    shInc.autoResizeColumns(1, 11);
  } else {
    var curHdr = shInc.getRange(1, 1, 1, shInc.getLastColumn()).getValues()[0].map(normHdr_);
    var curSet = {};
    curHdr.forEach(function(h) { if (h) curSet[h] = true; });
    var newCols = [];
    if (!curSet["CUBRE_ID"])     newCols.push("CUBRE_ID");
    if (!curSet["CUBRE_NOMBRE"]) newCols.push("CUBRE_NOMBRE");
    if (!curSet["AUTORIZA"])     newCols.push("AUTORIZA");
    if (newCols.length) {
      shInc.getRange(1, shInc.getLastColumn() + 1, 1, newCols.length).setValues([newCols]);
    }
  }

  var hdrNow = shInc.getRange(1, 1, 1, shInc.getLastColumn()).getValues()[0].map(normHdr_);
  var hdrIdx = {};
  hdrNow.forEach(function(h, i) { if (h) hdrIdx[h] = i + 1; });

  var newRow = [];
  for (var ci = 0; ci < shInc.getLastColumn(); ci++) newRow.push("");

  function setC(name, val) {
    var idx = hdrIdx[name];
    if (idx) newRow[idx - 1] = val;
  }

  setC("FECHA_ISO",     fecha);
  setC("SEDE",          sede);
  setC("ID",            id);
  setC("NOMBRE",        nombre);
  setC("TIPO",          "DT");
  setC("OBSERVACIONES", obs || ("Turno eventual " + turno));
  setC("CUBRE_ID",      "");
  setC("CUBRE_NOMBRE",  cubreA);
  setC("AUTORIZA",      autoriza);
  setC("CAPTURADO_POR", ses.username || "");
  setC("TS_ISO",        ts);

  shInc.appendRow(newRow);

  // 3) Sincroniza al pase de lista como DT
  try {
    var abrev = abrevBySede_(sede);
    if (abrev) {
      var pl = ss.getSheetByName(SH_PL);
      if (pl) {
        var key      = makeKey_(fecha, abrev, id);
        var existing = readDayRowsByKey_(fecha, abrev);
        var q        = quincenaFromISO_(fecha);
        if (existing[key]) {
          pl.getRange(existing[key].row, 7).setValue("DT");
          pl.getRange(existing[key].row, 8).setValue(ses.username || "");
          pl.getRange(existing[key].row, 9).setValue(ts);
        } else {
          pl.appendRow([fecha, q, sede, abrev, id, nombre, "DT", ses.username || "", ts, key]);
        }
      }
    }
  } catch (e_) {}

  return {
    ok: true,
    message: "Turno eventual registrado: " + nombre + " en " + sede
      + " (" + turno + ") el " + fecha
      + (cubreA ? ", cubre a " + cubreA : "")
      + ". Autoriza: " + autoriza + "."
  };
}

function apiGetTurnosEventuales(payload) {
  var token = String((payload && payload.token) || "");
  requireSession_(token);

  var sede      = String((payload && payload.sede)      || "").trim();
  var fechaIni  = String((payload && payload.fechaInicio) || "").trim();
  var fechaFin  = String((payload && payload.fechaFin)    || "").trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_TURNOS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };

  var v   = sh.getDataRange().getValues();
  var hdr = v[0].map(normHdr_);
  var idx = {};
  hdr.forEach(function(h, i) { if (h) idx[h] = i; });

  function sg(row, col) { return col !== undefined ? String(row[col] || "").trim() : ""; }

  var out = [];
  for (var i = 1; i < v.length; i++) {
    var rowFecha = formatCellValue_(v[i][idx["FECHA_ISO"] !== undefined ? idx["FECHA_ISO"] : 0]);
    var rowSede  = sg(v[i], idx["SEDE"]);
    if (sede && rowSede !== sede) continue;
    if (fechaIni && rowFecha < fechaIni) continue;
    if (fechaFin && rowFecha > fechaFin) continue;
    out.push({
      fecha:    rowFecha,
      nombre:   sg(v[i], idx["NOMBRE"]),
      id:       formatCellValue_(v[i][idx["ID"] !== undefined ? idx["ID"] : 2]),
      sede:     rowSede,
      turno:    sg(v[i], idx["TURNO"]),
      cubreA:   sg(v[i], idx["CUBRE_A"]),
      obs:      sg(v[i], idx["OBS"]),
      autoriza: sg(v[i], idx["AUTORIZA"]),
      externo:  sg(v[i], idx["ES_EXTERNO"]),
      captPor:  sg(v[i], idx["CAPTURADO_POR"])
    });
  }

  out.sort(function(a, b) {
    if (a.fecha < b.fecha) return 1;
    if (a.fecha > b.fecha) return -1;
    return 0;
  });

  return { ok: true, items: out };
}

// ══════════════════════════════════════════
function apiExportarIncidencias(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede exportar incidencias.");

  var month = String((payload && payload.month) || "").trim(); // yyyy-MM
  var sede  = String((payload && payload.sede)  || "").trim();

  if (!/^\d{4}-\d{2}$/.test(month)) {
    throw new Error("Mes inválido. Usa formato yyyy-MM.");
  }

  var data = apiGetIncidenciasCalendario({
    token: token,
    month: month,
    sede: sede
  });

  var items = (data && data.items) ? data.items : [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tagSede = (!sede || sede === "Todas las sedes") ? "TODAS" : sede.replace(/[^\w]+/g, "_").substring(0, 20);
  var shName = "RPT_INCID_" + month.replace("-", "") + "_" + tagSede;

  var old = ss.getSheetByName(shName);
  if (old) ss.deleteSheet(old);
  var sh = ss.insertSheet(shName);

  sh.getRange(1,1).setValue("MHS Integradora - Reporte de Incidencias")
    .setFontSize(13).setFontWeight("bold").setFontColor("#1a3a6c");
  sh.getRange(1,3).setValue("Mes: " + month).setFontColor("#444444");
  sh.getRange(1,5).setValue("Sede: " + ((!sede || sede === "Todas las sedes") ? "Todas las sedes" : sede)).setFontColor("#444444");
  sh.getRange(1,7).setValue("Generado: " + isoNow_()).setFontColor("#888888");

  var headers = [
    "FECHA_ISO","SEDE","ID","NOMBRE","TIPO","OBSERVACIONES",
    "CUBRE_ID","CUBRE_NOMBRE","AUTORIZA","CAPTURADO_POR","TS_ISO","ORIGEN"
  ];
  sh.getRange(2,1,1,headers.length).setValues([headers]);
  sh.getRange(2,1,1,headers.length)
    .setBackground("#1a3a6c")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sh.setFrozenRows(2);

  if (!items.length) {
    sh.getRange(3,1).setValue("Sin incidencias para el período seleccionado.");
    sh.autoResizeColumns(1, headers.length);
    return { ok: true, message: "Reporte creado sin registros.", url: ss.getUrl() + "#gid=" + sh.getSheetId() };
  }

  var rows = items.map(function(it) {
    return [
      String(it.fecha || "").trim(),
      String(it.sede || "").trim(),
      String(it.id || "").trim(),
      String(it.nombre || "").trim(),
      String(it.tipo || "").trim().toUpperCase(),
      String(it.obs || "").trim(),
      String(it.cubreId || "").trim(),
      String(it.cubreNombre || "").trim(),
      String(it.autoriza || "").trim(),
      String(it.captPor || "").trim(),
      String(it.tsISO || "").trim(),
      String(it.origen || "INCIDENCIAS").trim()
    ];
  });

  sh.getRange(3,1,rows.length,headers.length).setValues(rows);

  // Pintar columna TIPO
  var tipoVals = rows.map(function(r) { return [r[4]]; });
  for (var i = 0; i < tipoVals.length; i++) {
    var color = getCodeColor_(tipoVals[i][0]);
    sh.getRange(3 + i, 5).setBackground(color.bg).setFontColor(color.fg).setFontWeight("bold");
  }

  sh.autoResizeColumns(1, headers.length);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(4, 230);
  sh.setColumnWidth(6, 260);
  sh.setColumnWidth(8, 180);
  sh.setColumnWidth(9, 160);
  sh.setColumnWidth(10, 140);
  sh.setColumnWidth(12, 140);

  // Resumen arriba
  var counts = { F:0, PCG:0, PSG:0, I:0, DS:0, DT:0, FER:0, INH:0 };
  rows.forEach(function(r) {
    var t = String(r[4] || "").trim().toUpperCase();
    if (counts[t] !== undefined) counts[t]++;
  });

  var resumenRow = rows.length + 5;
  sh.getRange(resumenRow, 1).setValue("RESUMEN").setFontWeight("bold").setFontColor("#1a3a6c");
  var resumen = [
    ["F", counts.F],
    ["PCG", counts.PCG],
    ["PSG", counts.PSG],
    ["I", counts.I],
    ["DS", counts.DS],
    ["DT", counts.DT],
    ["FER", counts.FER],
    ["INH", counts.INH]
  ];

  resumen.forEach(function(x, idx) {
    var rr = resumenRow + 1 + idx;
    var color = getCodeColor_(x[0]);
    sh.getRange(rr, 1).setValue(x[0]).setBackground(color.bg).setFontColor(color.fg).setFontWeight("bold");
    sh.getRange(rr, 2).setValue(x[1]);
  });

  return {
    ok: true,
    message: "Reporte de incidencias exportado correctamente.",
    url: ss.getUrl() + "#gid=" + sh.getSheetId(),
    sheetName: shName
  };
}

// ══════════════════════════════════════════
function apiExportarTurnosEventuales(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede exportar turnos eventuales.");

  var mes   = String((payload && payload.mes)  || todayISO_().substring(0,7)).trim();
  var sede  = String((payload && payload.sede) || "").trim();
  var label = String((payload && payload.label)|| mes).trim();

  var p = mes.split("-");
  var y = p[0], m = p[1];
  var dias = new Date(Number(y), Number(m), 0).getDate();
  var ini  = y+"-"+m+"-01";
  var fin  = y+"-"+m+"-"+String(dias).padStart(2,"0");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("TURNOS_EVENTUALES");
  if (!sh || sh.getLastRow() < 2) throw new Error("No hay datos en TURNOS_EVENTUALES.");

  var v   = sh.getDataRange().getValues();
  var hdr = v[0].map(normHdr_);
  var idx = {};
  hdr.forEach(function(h, i) { if (h) idx[h] = i; });

  function sg(row, col) { return col !== undefined ? String(row[col]||"").trim() : ""; }

  var rows = [];
  for (var i = 1; i < v.length; i++) {
    var rowFecha = formatCellValue_(v[i][idx["FECHA_ISO"]!==undefined ? idx["FECHA_ISO"] : 0]);
    var rowSede  = sg(v[i], idx["SEDE"]);
    if (rowFecha < ini || rowFecha > fin) continue;
    if (sede && rowSede !== sede) continue;
    rows.push({
      fecha:    rowFecha,
      nombre:   sg(v[i], idx["NOMBRE"]),
      id:       formatCellValue_(v[i][idx["ID"]!==undefined ? idx["ID"] : 2]),
      sede:     rowSede,
      turno:    sg(v[i], idx["TURNO"]),
      cubreA:   sg(v[i], idx["CUBRE_A"]),
      obs:      sg(v[i], idx["OBS"]),
      autoriza: sg(v[i], idx["AUTORIZA"]),
      externo:  sg(v[i], idx["ES_EXTERNO"]),
      captPor:  sg(v[i], idx["CAPTURADO_POR"]),
      ts:       sg(v[i], idx["TS_ISO"])
    });
  }

  // Crear hoja
  var shName = "TE_"+y+m+(sede ? "_"+String(sede).substring(0,10).replace(/ /g,"_") : "_TODAS");
  var existing = ss.getSheetByName(shName);
  if (existing) ss.deleteSheet(existing);
  var out = ss.insertSheet(shName);

  var headers = ["FECHA","NOMBRE","ID","SEDE","TURNO","CUBRE_A","OBSERVACIONES","AUTORIZA","EXTERNO","CAPTURADO_POR","TIMESTAMP"];
  out.getRange(1, 1, 1, headers.length).setValues([headers])
     .setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold");
  out.setFrozenRows(1);
  out.getRange(1, headers.length+2).setValue("Turnos Eventuales · "+label+" · Generado: "+isoNow_()+" · Por: "+ses.username)
     .setFontColor("#888888").setFontSize(10);

  if (rows.length) {
    var dataRows = rows.map(function(r) {
      return [r.fecha, r.nombre, r.id, r.sede, r.turno, r.cubreA,
              r.obs, r.autoriza, r.externo, r.captPor, r.ts];
    });
    out.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    // Colorear turno
    var TURNO_COLORS = {MATUTINO:"#C6EFCE",VESPERTINO:"#CFE2F3",NOCTURNO:"#EAD1DC",MIXTO:"#FFF2CC",COMPLETO:"#F4CCCC"};
    for (var ri=0;ri<rows.length;ri++){
      var bg = TURNO_COLORS[rows[ri].turno]||"#FAFAFA";
      out.getRange(ri+2, 5, 1, 1).setBackground(bg).setFontWeight("bold");
      if(rows[ri].externo==="SI") out.getRange(ri+2, 9, 1, 1).setBackground("#FFF2CC").setFontWeight("bold");
    }
  } else {
    out.getRange(2,1).setValue("Sin turnos eventuales para el período seleccionado.");
  }

  out.autoResizeColumns(1, headers.length);
  out.setColumnWidth(2, 200); // nombre
  out.setColumnWidth(7, 280); // obs
  out.setColumnWidth(8, 180); // autoriza

  return {
    ok: true,
    message: "Hoja '"+shName+"' creada con "+rows.length+" turnos eventuales.",
    url: ss.getUrl()+"#gid="+out.getSheetId()
  };
}

// ══════════════════════════════════════════
//  CEO LIVE DASHBOARD
// ══════════════════════════════════════════
function getUserMapByUsername_() {
  var users = readUsers_();
  var map = {};
  users.forEach(function(u) {
    var key = String(u.username || "").toLowerCase().trim();
    if (!key) return;
    map[key] = {
      username: key,
      fullName: String(u.fullName || "").trim() || key,
      role: String(u.role || "").trim().toUpperCase()
    };
  });
  return map;
}

function getRangoCEO_(payload) {
  var periodo = String((payload && payload.periodo) || "mes").trim().toLowerCase();

  if (periodo === "dia") {
    var fecha = String((payload && payload.fecha) || todayISO_()).trim();
    if (!isoValid_(fecha)) throw new Error("Fecha invalida para CEO.");
    return { periodo: "dia", fechaInicio: fecha, fechaFin: fecha, label: fecha };
  }

  if (periodo === "semana") {
    var fi = String((payload && payload.fechaInicio) || "").trim();
    var ff = String((payload && payload.fechaFin) || "").trim();
    if (!isoValid_(fi) || !isoValid_(ff)) throw new Error("Rango semanal invalido.");
    if (fi > ff) throw new Error("fechaInicio no puede ser mayor a fechaFin.");
    return { periodo: "semana", fechaInicio: fi, fechaFin: ff, label: fi + " al " + ff };
  }

  var month = String((payload && payload.month) || "").trim();
  if (!/^\d{4}-\d{2}$/.test(month)) month = todayISO_().substring(0, 7);

  var y = Number(month.substring(0, 4));
  var m = Number(month.substring(5, 7));
  var fi2 = month + "-01";
  var ff2 = month + "-" + String(daysInMonth_(y, m)).padStart(2, "0");

  return { periodo: "mes", fechaInicio: fi2, fechaFin: ff2, label: month };
}

function getPaseRowsRango_(fechaInicio, fechaFin, sede) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  if (!sh || sh.getLastRow() < 2) return [];

  var v = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var fecha = formatCellValue_(v[i][0]);
    var rowSede = String(v[i][2] || "").trim();
    var codigo = String(v[i][6] || "").trim().toUpperCase();
    if (fecha < fechaInicio || fecha > fechaFin) continue;
    if (sede && rowSede !== sede) continue;
    out.push({
      fecha: fecha,
      sede: rowSede,
      abrev: String(v[i][3] || "").trim(),
      id: formatCellValue_(v[i][4]),
      nombre: String(v[i][5] || "").trim(),
      codigo: codigo,
      capturadoPor: String(v[i][7] || "").trim(),
      tsISO: String(v[i][8] || "").trim()
    });
  }
  return out;
}


function getPeriodoAnteriorCEO_(rango) {
  var fi = String((rango && rango.fechaInicio) || "").trim();
  var ff = String((rango && rango.fechaFin) || "").trim();
  if (!fi || !ff) return { fechaInicio: fi, fechaFin: ff, label: "anterior" };

  function diffDays_(a, b) {
    var pa = a.split("-"), pb = b.split("-");
    var da = new Date(Number(pa[0]), Number(pa[1]) - 1, Number(pa[2]));
    var db = new Date(Number(pb[0]), Number(pb[1]) - 1, Number(pb[2]));
    return Math.round((db.getTime() - da.getTime()) / 86400000);
  }

  if (String(rango.periodo || "") === "mes") {
    var y = Number(fi.substring(0,4));
    var m = Number(fi.substring(5,7));
    m -= 1;
    if (m <= 0) { m = 12; y -= 1; }
    var ym = y + "-" + String(m).padStart(2, "0");
    return {
      fechaInicio: ym + "-01",
      fechaFin: ym + "-" + String(daysInMonth_(y, m)).padStart(2, "0"),
      label: ym
    };
  }

  var days = diffDays_(fi, ff) + 1;
  return {
    fechaInicio: dateAddDaysIso_(fi, -days),
    fechaFin: dateAddDaysIso_(fi, -1),
    label: "anterior"
  };
}

function dateAddDaysIso_(iso, days) {
  var p = String(iso || todayISO_()).split("-");
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  d.setDate(d.getDate() + Number(days || 0));
  return Utilities.formatDate(d, TZ, "yyyy-MM-dd");
}

function buildCEOPdfHtml_(data) {
  data = data || {};
  var k = data.kpis || {};
  var sedes = data.sedes || [];
  var comparativoSede = data.comparativoPorSede || [];
  var supervisors = data.supervisores || [];
  var compare = data.comparativo || {};
  var trends = data.trends || [];
  var rango = data.rango || {};
  var generatedAt = data.generatedAt || isoNow_();

  function esc_(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function chip_(label, value, color) {
    return "<div class='chip'><div class='lbl'>" + esc_(label) + "</div><div class='val' style='color:" + color + "'>" + esc_(value) + "</div></div>";
  }

  var chips = [
    chip_("Asistencia global", (k.asistenciaPct || 0) + "%", "#276221"),
    chip_("Ausentismo", (k.ausentismoPct || 0) + "%", "#990000"),
    chip_("Incidencias", k.totalIncidencias || 0, "#b45f06"),
    chip_("Sedes críticas", k.sedesCriticas || 0, "#5b0f00"),
    chip_("Costo estimado", k.costoEstimado || "$0", "#0b5394")
  ].join("");

  var sedesRows = sedes.slice(0, 12).map(function(s) {
    return "<tr><td>" + esc_(s.sede) + "</td><td>" + esc_(s.asistenciaPct) + "%</td><td>" + esc_(s.faltas) + "</td><td>" + esc_(s.incidencias) + "</td></tr>";
  }).join("") || "<tr><td colspan='4'>Sin datos</td></tr>";

  var compRows = comparativoSede.slice(0, 12).map(function(s) {
    var delta = Number(s.deltaAsistenciaPct || 0);
    var sign = delta > 0 ? "+" : "";
    return "<tr><td>" + esc_(s.sede) + "</td><td>" + esc_(s.actualAsistenciaPct) + "%</td><td>" + esc_(s.prevAsistenciaPct) + "%</td><td>" + sign + esc_(delta) + "%</td><td>" + esc_(s.actualIncidencias) + "</td><td>" + esc_(s.prevIncidencias) + "</td></tr>";
  }).join("") || "<tr><td colspan='6'>Sin comparativo</td></tr>";

  var supRows = supervisors.slice(0, 10).map(function(s) {
    return "<tr><td>" + esc_(s.fullName || s.username) + "</td><td>" + esc_(s.registros) + "</td><td>" + esc_(s.incidencias) + "</td><td>" + esc_(s.nivel) + "</td></tr>";
  }).join("") || "<tr><td colspan='4'>Sin datos</td></tr>";

  var trendRows = trends.map(function(t) {
    return "<tr><td>" + esc_(t.label) + "</td><td>" + esc_(t.start) + " a " + esc_(t.end) + "</td><td>" + esc_(t.asistenciaPct) + "%</td><td>" + esc_(t.faltas) + "</td><td>" + esc_(t.incidencias) + "</td></tr>";
  }).join("") || "<tr><td colspan='5'>Sin tendencia</td></tr>";

  var q1 = compare.q1 || {};
  var q2 = compare.q2 || {};
  var deltaQ = Number(compare.deltaAsistenciaPct || 0);
  var signQ = deltaQ > 0 ? "+" : "";

  return "<!DOCTYPE html><html><head><meta charset='UTF-8'><style>" +
    "body{font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;padding:22px;font-size:10px}" +
    "h1{font-size:20px;color:#0b2d5c;margin:0 0 4px} .sub{color:#555;margin-bottom:12px}" +
    ".chips{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 18px} .chip{border:1px solid #d8e2f0;border-radius:12px;padding:8px 12px;background:#f7faff;min-width:120px}" +
    ".lbl{font-size:9px;color:#667;text-transform:uppercase;margin-bottom:4px}.val{font-size:16px;font-weight:700}" +
    ".sec{margin-top:18px} .sec h2{font-size:13px;color:#0b2d5c;margin:0 0 8px}" +
    "table{width:100%;border-collapse:collapse;font-size:9px} th,td{border:1px solid #d9e2f3;padding:5px 6px;text-align:left} th{background:#0b2d5c;color:#fff}" +
    ".qbox{display:flex;gap:12px;margin-top:8px} .qcard{flex:1;border:1px solid #d9e2f3;border-radius:12px;padding:10px;background:#fafcff}" +
    ".foot{margin-top:14px;color:#888;font-size:9px;text-align:right}" +
    "</style></head><body>" +
    "<h1>MHS CEO LIVE</h1>" +
    "<div class='sub'>Reporte ejecutivo · Periodo: " + esc_(rango.label || "") + " · Generado: " + esc_(generatedAt) + "</div>" +
    "<div class='chips'>" + chips + "</div>" +

    "<div class='sec'><h2>Comparativo Q1 vs Q2</h2><div class='qbox'>" +
    "<div class='qcard'><b>Q1</b><br>Asistencia: " + esc_(q1.asistenciaPct || 0) + "%<br>Faltas: " + esc_(q1.faltas || 0) + "<br>Incidencias: " + esc_(q1.incidencias || 0) + "</div>" +
    "<div class='qcard'><b>Q2</b><br>Asistencia: " + esc_(q2.asistenciaPct || 0) + "%<br>Faltas: " + esc_(q2.faltas || 0) + "<br>Incidencias: " + esc_(q2.incidencias || 0) + "</div>" +
    "<div class='qcard'><b>Delta asistencia</b><br><span style='font-size:18px;font-weight:700'>" + signQ + esc_(deltaQ) + "%</span></div>" +
    "</div></div>" +

    "<div class='sec'><h2>Mapa por sede</h2><table><thead><tr><th>Sede</th><th>Asistencia</th><th>Faltas</th><th>Incidencias</th></tr></thead><tbody>" + sedesRows + "</tbody></table></div>" +
    "<div class='sec'><h2>Comparativo por sede</h2><table><thead><tr><th>Sede</th><th>Actual</th><th>Anterior</th><th>Δ Asistencia</th><th>Incid. actuales</th><th>Incid. anteriores</th></tr></thead><tbody>" + compRows + "</tbody></table></div>" +
    "<div class='sec'><h2>Ranking de supervisores</h2><table><thead><tr><th>Supervisor</th><th>Registros</th><th>Incidencias</th><th>Nivel</th></tr></thead><tbody>" + supRows + "</tbody></table></div>" +
    "<div class='sec'><h2>Tendencia 7 / 15 / 30 días</h2><table><thead><tr><th>Ventana</th><th>Periodo</th><th>Asistencia</th><th>Faltas</th><th>Incidencias</th></tr></thead><tbody>" + trendRows + "</tbody></table></div>" +
    "<div class='foot'>MHS CEO LIVE · " + esc_(generatedAt) + "</div>" +
    "</body></html>";
}

function apiExportCEOPDF(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  var role = String(ses.role || "").toUpperCase();
  if (!(role === "CEO" || role === "ADMIN" || role === "SUPERADMIN")) {
    throw new Error("Solo CEO / ADMIN / SUPERADMIN puede exportar PDF ejecutivo.");
  }

  var data = apiGetCEODashboard(payload || {});
  var html = buildCEOPdfHtml_(data);
  var blob = Utilities.newBlob(html, "text/html", "ceo_dashboard.html")
    .getAs("application/pdf")
    .setName("CEO_MHS_" + String((data.rango && data.rango.label) || todayISO_()).replace(/[^\w-]+/g, "_") + ".pdf");

  var folder = null;
  try {
    folder = DriveApp.getFolderById(DRIVE_PDF_FOLDER_ID);
  } catch (e) {}

  var file = folder ? folder.createFile(blob) : DriveApp.createFile(blob);
  return {
    ok: true,
    url: file.getUrl(),
    fileName: file.getName()
  };
}

// ══════════════════════════════════════════
//  CEO · COSTO POR FALTA SEGÚN SEDE
// ══════════════════════════════════════════
function normalizeSedeCosto_(sede) {
  return String(sede || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getCostoFaltaBySede_(sede) {
  var s = normalizeSedeCosto_(sede);

  // Ohorán y variantes
  if (s.indexOf("OHORAN") >= 0) return 850.00;

  // UNEME y variantes (ej. UNEME ONCOLOGIA)
  if (s.indexOf("UNEME") >= 0) return 850.00;

  // Todas las demás sedes
  return 771.20;
}

function calcularCostoFaltasCEO_(rows) {
  rows = rows || [];

  var total = 0;
  var detallePorSede = {};

  rows.forEach(function(r) {
    var codigo = String(r.codigo || "").trim().toUpperCase();
    var sede = String(r.sede || "SIN SEDE").trim();

    // Solo las faltas generan costo
    if (codigo !== "F") return;

    var costo = getCostoFaltaBySede_(sede);
    total += costo;

    if (!detallePorSede[sede]) {
      detallePorSede[sede] = {
        sede: sede,
        faltas: 0,
        costoUnitario: costo,
        costoTotal: 0
      };
    }

    detallePorSede[sede].faltas++;
    detallePorSede[sede].costoTotal += costo;
  });

  return {
    total: total,
    detallePorSede: Object.keys(detallePorSede).map(function(k) {
      return detallePorSede[k];
    }).sort(function(a, b) {
      return b.costoTotal - a.costoTotal;
    })
  };
}

function apiGetCEODashboard(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  var role = String(ses.role || "").toUpperCase();
  if (!(role === "CEO" || role === "ADMIN" || role === "SUPERADMIN")) {
    throw new Error("Solo CEO / ADMIN / SUPERADMIN puede ver el dashboard CEO.");
  }

  var rango = getRangoCEO_(payload || {});
  var sede = String((payload && payload.sede) || "").trim();
  var paseRows = getPaseRowsRango_(rango.fechaInicio, rango.fechaFin, sede);
  var userMap = getUserMapByUsername_();
  var users = readUsers_();

  var NEGATIVE_CODES = { F:1, PCG:1, PSG:1, I:1, DT:1, FER:1, INH:1 };
  var ALL_CODES = { A:1, AF:1, F:1, PCG:1, PSG:1, I:1, DT:1, DS:1, FER:1, INH:1 };

  function summarizeRows_(rows) {
    var counts = { A:0, AF:0, F:0, PCG:0, PSG:0, I:0, DT:0, DS:0, FER:0, INH:0 };
    (rows || []).forEach(function(r) {
      var cod = String(r.codigo || "").trim().toUpperCase();
      if (counts[cod] !== undefined) counts[cod]++;
    });

    var total = (rows || []).length;
    var asist = (counts.A || 0) + (counts.AF || 0);
    var faltas = counts.F || 0;
    var negativas = (counts.F || 0) + (counts.PCG || 0) + (counts.PSG || 0) + (counts.I || 0) + (counts.DT || 0) + (counts.FER || 0) + (counts.INH || 0);

    return {
      totalRegistros: total,
      asistenciaPct: total ? Number(((asist / total) * 100).toFixed(1)) : 0,
      ausentismoPct: total ? Number(((faltas / total) * 100).toFixed(1)) : 0,
      faltas: faltas,
      incidencias: negativas,
      counts: counts
    };
  }

  function summarizeRowsBySede_(rows) {
    var map = {};
    (rows || []).forEach(function(r) {
      var sd = String(r.sede || "SIN SEDE").trim();
      if (!map[sd]) map[sd] = [];
      map[sd].push(r);
    });
    return Object.keys(map).map(function(sd) {
      var s = summarizeRows_(map[sd]);
      return { sede: sd, asistenciaPct: s.asistenciaPct, faltas: s.faltas, incidencias: s.incidencias };
    });
  }

  function dateAddDays_(iso, days) {
    var p = String(iso || todayISO_()).split("-");
    var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    d.setDate(d.getDate() + days);
    return Utilities.formatDate(d, TZ, "yyyy-MM-dd");
  }

  var counts = { A:0, AF:0, F:0, PCG:0, PSG:0, I:0, DT:0, DS:0, FER:0, INH:0 };
  var negCounts = { F:0, PCG:0, PSG:0, I:0, DT:0, FER:0, INH:0 };
  var porSede = {};
  var porSupervisor = {};

  paseRows.forEach(function(r) {
    var cod = String(r.codigo || "").trim().toUpperCase();
    if (!ALL_CODES[cod]) return;

    var sd = String(r.sede || "SIN SEDE").trim();
    var capt = String(r.capturadoPor || "").toLowerCase().trim();
    var userInfo = userMap[capt] || {
      username: capt || "sin_usuario",
      fullName: capt || "Sin supervisor",
      role: ""
    };
    var supName = userInfo.fullName || userInfo.username || "Sin supervisor";

    if (counts[cod] !== undefined) counts[cod]++;
    if (NEGATIVE_CODES[cod]) negCounts[cod]++;

    if (!porSede[sd]) {
      porSede[sd] = { sede: sd, total: 0, A:0, AF:0, F:0, PCG:0, PSG:0, I:0, DT:0, DS:0, FER:0, INH:0 };
    }
    porSede[sd].total++;
    if (porSede[sd][cod] !== undefined) porSede[sd][cod]++;

    if (!porSupervisor[supName]) {
      porSupervisor[supName] = {
        username: userInfo.username,
        fullName: supName,
        registros: 0,
        incidencias: 0,
        role: userInfo.role
      };
    }
    porSupervisor[supName].registros++;
    if (NEGATIVE_CODES[cod]) porSupervisor[supName].incidencias++;
  });

  var totalRegistros = paseRows.length;
  var totalAsist = (counts.A || 0) + (counts.AF || 0);
  var totalFaltas = counts.F || 0;
  var totalIncidencias = (negCounts.F || 0) + (negCounts.PCG || 0) + (negCounts.PSG || 0) + (negCounts.I || 0) + (negCounts.DT || 0) + (negCounts.FER || 0) + (negCounts.INH || 0);

  var asistenciaPct = totalRegistros ? Number(((totalAsist / totalRegistros) * 100).toFixed(1)) : 0;
  var ausentismoPct = totalRegistros ? Number(((totalFaltas / totalRegistros) * 100).toFixed(1)) : 0;

  var costos = calcularCostoFaltasCEO_(paseRows);
  var costoEstimado = "$" + Number(costos.total || 0).toLocaleString("es-MX", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });

  var sedesArr = Object.keys(porSede).map(function(k) {
    var x = porSede[k];
    var asist = x.A + x.AF;
    var incid = x.F + x.PCG + x.PSG + x.I + x.DT + x.FER + x.INH;
    var pct = x.total ? Number(((asist / x.total) * 100).toFixed(1)) : 0;
    return {
      sede: x.sede,
      total: x.total,
      asistenciaPct: pct,
      faltas: x.F,
      incidencias: incid,
      detalle: {
        F:x.F, PCG:x.PCG, PSG:x.PSG, I:x.I, DT:x.DT, FER:x.FER, INH:x.INH, DS:x.DS, A:x.A, AF:x.AF
      }
    };
  }).sort(function(a, b) { return a.asistenciaPct - b.asistenciaPct; });

  var supArr = Object.keys(porSupervisor).map(function(k) {
    var x = porSupervisor[k];
    var ratio = x.registros ? Number((x.incidencias / x.registros).toFixed(3)) : 0;
    var badge = "info";
    var nivel = "Activo";
    if (x.registros >= 80 && ratio < 0.08) { badge = "ok"; nivel = "Excelente"; }
    else if (x.registros >= 30 && ratio < 0.18) { badge = "warn"; nivel = "Regular"; }
    else if (x.registros > 0) { badge = "bad"; nivel = "Bajo"; }
    return {
      username:x.username,
      fullName:x.fullName,
      registros:x.registros,
      incidencias:x.incidencias,
      ratioIncidencia:ratio,
      badge:badge,
      nivel:nivel
    };
  }).sort(function(a, b) {
    if (b.registros !== a.registros) return b.registros - a.registros;
    return a.ratioIncidencia - b.ratioIncidencia;
  });

  var supervisorSinCaptura = users
    .filter(function(u) {
      var roleU = String(u.role || "").toUpperCase();
      return u.active && roleU !== "CEO";
    })
    .map(function(u) {
      var registros = 0;
      Object.keys(porSupervisor).forEach(function(k) {
        var x = porSupervisor[k];
        if (String(x.username || "").toLowerCase() === String(u.username || "").toLowerCase()) {
          registros += Number(x.registros || 0);
        }
      });
      return {
        username: String(u.username || "").toLowerCase().trim(),
        fullName: String(u.fullName || "").trim() || String(u.username || "").trim(),
        role: String(u.role || "").toUpperCase(),
        registros: registros
      };
    })
    .filter(function(x) { return x.registros === 0; })
    .sort(function(a, b) { return a.fullName.localeCompare(b.fullName, "es"); });

  var alertas = [];
  sedesArr.forEach(function(s) {
    if (s.asistenciaPct < 80) {
      alertas.push({ icon:"🔴", title:s.sede + " con asistencia baja", detail:"Asistencia de " + s.asistenciaPct + "% en el periodo." });
    }
    if (s.faltas >= 5) {
      alertas.push({ icon:"⚠️", title:s.sede + " concentra faltas", detail:"Se detectaron " + s.faltas + " faltas en el periodo." });
    }
    if (s.incidencias >= 10) {
      alertas.push({ icon:"📌", title:s.sede + " acumula incidencias", detail:"Se detectaron " + s.incidencias + " incidencias operativas en el periodo." });
    }
  });

  supervisorSinCaptura.slice(0, 6).forEach(function(sup) {
    alertas.push({
      icon: "👤",
      title: sup.fullName + " sin captura",
      detail: "No registró movimientos en el periodo seleccionado."
    });
  });

  var sedesCriticas = sedesArr.filter(function(s) {
    return s.asistenciaPct < 80 || s.faltas >= 5 || s.incidencias >= 10;
  }).length;

  var semaforo = "ok";
  if (sedesCriticas >= 3 || asistenciaPct < 80) semaforo = "bad";
  else if (sedesCriticas >= 1 || asistenciaPct < 90) semaforo = "warn";

  var compareMonth = String((payload && payload.month) || "").trim();
  if (!/^\d{4}-\d{2}$/.test(compareMonth)) compareMonth = rango.fechaInicio.substring(0, 7);
  var y = Number(compareMonth.substring(0, 4));
  var m = Number(compareMonth.substring(5, 7));

  var q1Rows = getPaseRowsRango_(compareMonth + "-01", compareMonth + "-15", sede);
  var q2Rows = getPaseRowsRango_(compareMonth + "-16", compareMonth + "-" + String(daysInMonth_(y, m)).padStart(2, "0"), sede);
  var q1 = summarizeRows_(q1Rows);
  var q2 = summarizeRows_(q2Rows);

  var prevRango = getPeriodoAnteriorCEO_(rango);
  var prevRows = getPaseRowsRango_(prevRango.fechaInicio, prevRango.fechaFin, sede);
  var prevBySede = {};
  summarizeRowsBySede_(prevRows).forEach(function(x) { prevBySede[x.sede] = x; });

  var comparativoPorSede = sedesArr.map(function(s) {
    var prev = prevBySede[s.sede] || { asistenciaPct:0, faltas:0, incidencias:0 };
    return {
      sede: s.sede,
      actualAsistenciaPct: s.asistenciaPct,
      prevAsistenciaPct: prev.asistenciaPct || 0,
      deltaAsistenciaPct: Number(((s.asistenciaPct || 0) - (prev.asistenciaPct || 0)).toFixed(1)),
      actualFaltas: s.faltas,
      prevFaltas: prev.faltas || 0,
      actualIncidencias: s.incidencias,
      prevIncidencias: prev.incidencias || 0
    };
  }).sort(function(a, b) { return a.deltaAsistenciaPct - b.deltaAsistenciaPct; });

  var trendAnchor = rango.fechaFin || todayISO_();
  function buildTrend_(days) {
    var start = dateAddDays_(trendAnchor, -(days - 1));
    var rows = getPaseRowsRango_(start, trendAnchor, sede);
    var s = summarizeRows_(rows);
    return {
      label: String(days) + " dias",
      start: start,
      end: trendAnchor,
      totalRegistros: s.totalRegistros,
      asistenciaPct: s.asistenciaPct,
      faltas: s.faltas,
      incidencias: s.incidencias
    };
  }
  var trends = [buildTrend_(7), buildTrend_(15), buildTrend_(30)];

  return {
    ok: true,
    generatedAt: isoNow_(),
    rango: rango,
    kpis: {
      plantillaActiva: totalRegistros,
      asistenciaPct: asistenciaPct,
      ausentismoPct: ausentismoPct,
      totalIncidencias: totalIncidencias,
      sedesCriticas: sedesCriticas,
      costoEstimado: costoEstimado,
      totalRegistros: totalRegistros,
      totalFaltas: totalFaltas,
      semaforo: semaforo
    },
    sedes: sedesArr,
    supervisores: supArr,
    alertas: alertas,
    supervisorSinCaptura: supervisorSinCaptura,
    comparativo: {
      month: compareMonth,
      q1: { asistenciaPct:q1.asistenciaPct, faltas:q1.faltas, incidencias:q1.incidencias, totalRegistros:q1.totalRegistros },
      q2: { asistenciaPct:q2.asistenciaPct, faltas:q2.faltas, incidencias:q2.incidencias, totalRegistros:q2.totalRegistros }
    },
    comparativoPorSede: comparativoPorSede,
    trends: trends,
    incidencias: {
      F: negCounts.F || 0,
      PCG: negCounts.PCG || 0,
      PSG: negCounts.PSG || 0,
      I: negCounts.I || 0,
      DT: negCounts.DT || 0,
      FER: negCounts.FER || 0,
      INH: negCounts.INH || 0
    },
    costoFaltasDetalle: costos.detallePorSede,
    costoFaltasRegla: "OHORAN / SALUD HOSPITAL OHORAN / UNEME ONCOLOGIA = 850.00 · demás sedes = 771.20"
  };
}
function apiExportCEOPDF(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  var role = String(ses.role || "").toUpperCase();
  if (!(role === "CEO" || role === "ADMIN" || role === "SUPERADMIN")) {
    throw new Error("Solo CEO / ADMIN / SUPERADMIN puede exportar PDF ejecutivo.");
  }

  var data = apiGetCEODashboard(payload || {});
  var html = buildCEOPdfHtml_(data);

  var blob = HtmlService.createHtmlOutput(html).getBlob().getAs("application/pdf");
  var rango = data.rango || {};
  var nombre = "CEO_" + String((rango.label || todayISO_())).replace(/[^\w-]+/g, "_") + "_" + Utilities.formatDate(new Date(), TZ, "yyyyMMdd_HHmmss") + ".pdf";
  blob.setName(nombre);

  var file;
  try {
    var folder = DriveApp.getFolderById(DRIVE_PDF_FOLDER_ID);
    file = folder.createFile(blob);
  } catch (e) {
    file = DriveApp.createFile(blob);
  }

  return {
    ok: true,
    fileId: file.getId(),
    url: file.getUrl(),
    name: nombre
  };
}

// ══════════════════════════════════════════════════════════════
//  MÓDULO: COMPENSACIÓN DE DESCANSO
// ══════════════════════════════════════════════════════════════
function apiGetFaltasPorSede(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  var sede        = String((payload && payload.sede)        || "").trim();
  var fechaInicio = String((payload && payload.fechaInicio) || todayISO_()).trim();
  var fechaFin    = String((payload && payload.fechaFin)    || fechaInicio).trim();
  if (!sede)                   throw new Error("Falta sede.");
  if (!isoValid_(fechaInicio)) throw new Error("fechaInicio invalida.");
  if (!isoValid_(fechaFin))    throw new Error("fechaFin invalida.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_PL);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };
  var diaDescansoMap = {};
  var shCon = ss.getSheetByName(SH_CONTRATOS);
  if (shCon && shCon.getLastRow() >= 2) {
    var cVals = shCon.getDataRange().getValues();
    var cHdr  = cVals[0].map(normHdr_);
    var cIdx  = {};
    cHdr.forEach(function(h, i) { if (h) cIdx[h] = i; });
    if (cIdx["ID"] !== undefined && cIdx["DIA_DESCANSO"] !== undefined) {
      for (var ci = 1; ci < cVals.length; ci++) {
        var cid = formatCellValue_(cVals[ci][cIdx["ID"]]);
        var cdd = String(cVals[ci][cIdx["DIA_DESCANSO"]] || "").trim().toUpperCase();
        if (cid) diaDescansoMap[cid] = cdd;
      }
    }
  }
  var v   = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var fecha   = formatCellValue_(v[i][0]);
    var rowSede = String(v[i][2] || "").trim();
    var codigo  = String(v[i][6] || "").trim().toUpperCase();
    if (rowSede !== sede) continue;
    if (fecha < fechaInicio || fecha > fechaFin) continue;
    if (codigo !== "F") continue;
    var id     = formatCellValue_(v[i][4]);
    var nombre = String(v[i][5] || "").trim();
    out.push({ fecha: fecha, sede: rowSede, id: id, nombre: nombre, codigo: codigo,
               diaDescanso: diaDescansoMap[id] || "", rowNum: i + 1 });
  }
  out.sort(function(a, b) {
    if (a.fecha < b.fecha) return -1; if (a.fecha > b.fecha) return 1;
    return String(a.nombre||"").localeCompare(String(b.nombre||""), "es");
  });
  return { ok: true, items: out };
}

function apiCompensarDescanso(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede aplicar compensaciones.");
  var sede           = String((payload && payload.sede)           || "").trim();
  var id             = String((payload && payload.id)             || "").trim();
  var nombre         = String((payload && payload.nombre)         || "").trim();
  var fechaFalta     = String((payload && payload.fechaFalta)     || "").trim();
  var fechaTrabajo   = String((payload && payload.fechaTrabajo)   || "").trim();
  var diaDescansoOrig= String((payload && payload.diaDescansoOrig)|| "").trim().toUpperCase();
  var codigoNuevo    = String((payload && payload.codigoNuevo)    || "A").trim().toUpperCase();
  var motivo         = String((payload && payload.motivo)         || "").trim();
  var autoriza       = String((payload && payload.autoriza)       || "").trim();
  if (!sede)                     throw new Error("Falta sede.");
  if (!id)                       throw new Error("Falta ID.");
  if (!isoValid_(fechaFalta))    throw new Error("fechaFalta invalida.");
  if (!isoValid_(fechaTrabajo))  throw new Error("fechaTrabajo invalida.");
  if (!autoriza)                 throw new Error("Falta el responsable que autoriza.");
  if (codigoNuevo !== "A" && codigoNuevo !== "AF") codigoNuevo = "A";
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var abrev = abrevBySede_(sede);
  if (!abrev) throw new Error("La sede '" + sede + "' no tiene ABREV.");
  var pl = ss.getSheetByName(SH_PL);
  if (!pl) throw new Error("No existe PASE_LISTA_V2.");
  var key = makeKey_(fechaFalta, abrev, id);
  var v   = pl.getDataRange().getValues();
  var targetRow = -1, codigoAnterior = "";
  for (var i = 1; i < v.length; i++) {
    var rowKey = String(v[i][9] || "").trim();
    if (rowKey === key) { codigoAnterior = String(v[i][6]||"").trim().toUpperCase(); targetRow = i+1; break; }
  }
  if (targetRow < 0) {
    var AB = abrev.toUpperCase();
    for (var j = 1; j < v.length; j++) {
      if (formatCellValue_(v[j][0]) === fechaFalta && String(v[j][3]||"").trim().toUpperCase() === AB && formatCellValue_(v[j][4]) === id) {
        codigoAnterior = String(v[j][6]||"").trim().toUpperCase(); targetRow = j+1; break;
      }
    }
  }
  if (targetRow < 0) throw new Error("No se encontró el registro de falta para ID " + id + " el " + fechaFalta);
  if (codigoAnterior !== "F") throw new Error("El código actual es '" + codigoAnterior + "', no es F.");
  var ts = isoNow_(); var capt = ses.username || "";
  pl.getRange(targetRow, 7).setValue(codigoNuevo);
  pl.getRange(targetRow, 8).setValue(capt);
  pl.getRange(targetRow, 9).setValue(ts);
  var HDR_COMP_LOCAL = ["FECHA_FALTA","FECHA_TRABAJO","SEDE","ID","NOMBRE","DIA_DESCANSO_ORIGINAL","CODIGO_ANTERIOR","CODIGO_NUEVO","MOTIVO","AUTORIZA","CAPTURADO_POR","TS_ISO"];
  var shComp = ss.getSheetByName("COMPENSACIONES") || (function(){ var s=ss.insertSheet("COMPENSACIONES"); s.getRange(1,1,1,HDR_COMP_LOCAL.length).setValues([HDR_COMP_LOCAL]); s.setFrozenRows(1); return s; })();
  shComp.appendRow([fechaFalta, fechaTrabajo, sede, id, nombre, diaDescansoOrig, codigoAnterior, codigoNuevo, motivo, autoriza, capt, ts]);
  try { var col = {bg:"#C6EFCE",fg:"#276221"}; pl.getRange(targetRow,7).setBackground(col.bg).setFontColor(col.fg).setFontWeight("bold"); } catch(_){}
  return { ok:true, message: nombre + " — Falta del " + fechaFalta + " compensada a " + codigoNuevo + ". Autoriza: " + autoriza + ".", codigoAnterior:codigoAnterior, codigoNuevo:codigoNuevo };
}

function apiGetSinMarcar(payload) {
  var token = String((payload && payload.token) || "");
  requireSession_(token);
  var sede  = String((payload && payload.sede)  || "").trim();
  var fecha = String((payload && payload.fecha) || "").trim();
  if (!sede)             throw new Error("Falta sede.");
  if (!isoValid_(fecha)) throw new Error("Fecha invalida.");
  var todos  = apiGetEmpleadosPorSede(sede);
  var marcas = readDayMarks_(fecha, abrevBySede_(sede));
  var sinMarcar = todos.filter(function(emp) { return !marcas[emp.id]; });
  return { ok:true, total:todos.length, marcados:todos.length-sinMarcar.length, sinMarcar:sinMarcar.length, items:sinMarcar };
}


// ══════════════════════════════════════════════════════════════
//  MÓDULO: CAMBIO DE DESCANSO TEMPORAL (CDT)
// ══════════════════════════════════════════════════════════════
var SH_CDT  = "CAMBIOS_DESCANSO_TEMP";
var HDR_CDT = ["FECHA_INICIO","FECHA_FIN","SEDE","ID","NOMBRE","DIA_DESCANSO_ORIG","DIA_DESCANSO_TEMP","MOTIVO","AUTORIZA","CAPTURADO_POR","TS_ISO","ACTIVO"];

function _ensureCDTSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CDT);
  if (!sh) {
    sh = ss.insertSheet(SH_CDT);
    sh.getRange(1,1,1,HDR_CDT.length).setValues([HDR_CDT]);
    sh.getRange(1,1,1,HDR_CDT.length).setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1); sh.autoResizeColumns(1,HDR_CDT.length);
  }
  return sh;
}

function apiCDTCrear(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN.");
  var sede        = String((payload && payload.sede)           || "").trim();
  var id          = String((payload && payload.id)             || "").trim();
  var nombre      = String((payload && payload.nombre)         || "").trim();
  var fechaInicio = String((payload && payload.fechaInicio)    || "").trim();
  var fechaFin    = String((payload && payload.fechaFin)       || fechaInicio).trim();
  var diaOrig     = String((payload && payload.diaDescansoOrig)|| "").trim().toUpperCase();
  var diaTemp     = String((payload && payload.diaDescansoTemp)|| "").trim().toUpperCase();
  var motivo      = String((payload && payload.motivo)         || "").trim().substring(0,800);
  var autoriza    = String((payload && payload.autoriza)       || "").trim();
  if (!sede)                   throw new Error("Falta sede.");
  if (!id)                     throw new Error("Falta ID.");
  if (!isoValid_(fechaInicio)) throw new Error("fechaInicio invalida.");
  if (!isoValid_(fechaFin))    throw new Error("fechaFin invalida.");
  if (fechaFin < fechaInicio)  throw new Error("fechaFin no puede ser anterior a fechaInicio.");
  if (!diaTemp)                throw new Error("Falta día de descanso temporal.");
  if (!autoriza)               throw new Error("Falta responsable que autoriza.");
  var DIAS_VALIDOS = ["LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO","DOMINGO","SABADO Y DOMINGO","NINGUNO"];
  if (DIAS_VALIDOS.indexOf(diaTemp) < 0) throw new Error("Día temporal inválido: " + diaTemp);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ts = isoNow_(); var capt = ses.username || "";
  var abrev = abrevBySede_(sede);
  var shCDT = _ensureCDTSheet_();
  shCDT.appendRow([fechaInicio, fechaFin, sede, id, nombre, diaOrig, diaTemp, motivo, autoriza, capt, ts, "SI"]);
  var registrados = [];
  var pl    = ss.getSheetByName(SH_PL);
  var shInc = ss.getSheetByName("INCIDENCIAS");
  if (pl && abrev) {
    var curParts = fechaInicio.split("-").map(Number);
    var curD = new Date(curParts[0], curParts[1]-1, curParts[2]);
    var finParts = fechaFin.split("-").map(Number);
    var finD = new Date(finParts[0], finParts[1]-1, finParts[2]);
    var DIAS_NOMBRE = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"];
    while (curD <= finD) {
      var fechaISO = Utilities.formatDate(curD, TZ, "yyyy-MM-dd");
      var dow = curD.getDay();
      var diaHoy = DIAS_NOMBRE[dow];
      var descansaTemp = (diaTemp === diaHoy) || (diaTemp === "SABADO Y DOMINGO" && (dow===0||dow===6));
      if (descansaTemp) {
        var key = makeKey_(fechaISO, abrev, id);
        var q   = quincenaFromISO_(fechaISO);
        var existing = readDayRowsByKey_(fechaISO, abrev);
        var obsText = "Cambio descanso temporal" + (motivo ? ": " + motivo : "") + ". Autoriza: " + autoriza;
        if (existing[key]) {
          if (existing[key].codigo !== "DS") { pl.getRange(existing[key].row,7).setValue("DS"); pl.getRange(existing[key].row,8).setValue(capt); pl.getRange(existing[key].row,9).setValue(ts); }
        } else {
          pl.appendRow([fechaISO, q, sede, abrev, id, nombre, "DS", capt, ts, key]);
        }
        if (shInc) {
          var hdrInc = shInc.getRange(1,1,1,shInc.getLastColumn()).getValues()[0].map(normHdr_);
          var hIdx = {}; hdrInc.forEach(function(h,i){if(h)hIdx[h]=i+1;});
          var newRow=[]; for(var ci=0;ci<shInc.getLastColumn();ci++) newRow.push("");
          function setCI(n,v){var ix=hIdx[n];if(ix)newRow[ix-1]=v;}
          setCI("FECHA_ISO",fechaISO); setCI("SEDE",sede); setCI("ID",id); setCI("NOMBRE",nombre);
          setCI("TIPO","DS"); setCI("OBSERVACIONES",obsText); setCI("AUTORIZA",autoriza);
          setCI("CAPTURADO_POR",capt); setCI("TS_ISO",ts);
          shInc.appendRow(newRow);
        }
        registrados.push(fechaISO);
      }
      curD.setDate(curD.getDate()+1);
    }
  }
  return { ok:true, message: nombre + " — CDT del " + fechaInicio + " al " + fechaFin + ". Día temp: " + diaTemp + ".", fechasDS:registrados, total:registrados.length };
}

function apiCDTListar(payload) {
  var token    = String((payload && payload.token) || "");
  requireSession_(token);
  var sede     = String((payload && payload.sede)      || "").trim();
  var filterId = String((payload && payload.id)        || "").trim();
  var soloAct  = payload && payload.soloActivos !== false;
  var fechaRef = String((payload && payload.fechaRef)  || todayISO_()).trim();
  if (!sede) throw new Error("Falta sede.");
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CDT);
  if (!sh || sh.getLastRow() < 2) return { ok:true, items:[] };
  var v=sh.getDataRange().getValues(); var hdr=v[0].map(normHdr_); var idx={}; hdr.forEach(function(h,i){if(h)idx[h]=i;});
  function sg(row,col){return col!==undefined?String(row[col]||"").trim():"";}
  var out=[];
  for(var i=1;i<v.length;i++){
    var rowSede=sg(v[i],idx["SEDE"]); var rowId=formatCellValue_(v[i][idx["ID"]!==undefined?idx["ID"]:3]);
    var activo=sg(v[i],idx["ACTIVO"]).toUpperCase(); var fi=sg(v[i],idx["FECHA_INICIO"]); var ff=sg(v[i],idx["FECHA_FIN"]);
    if(rowSede!==sede) continue; if(filterId&&rowId!==filterId) continue; if(soloAct&&activo!=="SI") continue;
    out.push({row:i+1,fechaInicio:fi,fechaFin:ff,sede:rowSede,id:rowId,nombre:sg(v[i],idx["NOMBRE"]),
      diaOrig:sg(v[i],idx["DIA_DESCANSO_ORIG"]),diaTemp:sg(v[i],idx["DIA_DESCANSO_TEMP"]),
      motivo:sg(v[i],idx["MOTIVO"]),autoriza:sg(v[i],idx["AUTORIZA"]),captPor:sg(v[i],idx["CAPTURADO_POR"]),
      tsISO:sg(v[i],idx["TS_ISO"]),activo:activo==="SI",vigente:activo==="SI"&&fi<=fechaRef&&ff>=fechaRef});
  }
  out.sort(function(a,b){return a.fechaInicio<b.fechaInicio?1:-1;});
  return {ok:true,items:out};
}

function apiCDTCancelar(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN.");
  var row  = Number((payload && payload.row) || 0);
  var sede = String((payload && payload.sede) || "").trim();
  if (row < 2) throw new Error("Fila inválida.");
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CDT);
  if (!sh) throw new Error("No existe " + SH_CDT);
  var hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(normHdr_);
  var idx={}; hdr.forEach(function(h,i){if(h)idx[h]=i;});
  var rowSede=String(sh.getRange(row,(idx["SEDE"]||0)+1).getValue()||"").trim();
  if(rowSede!==sede) throw new Error("La fila no pertenece a la sede "+sede);
  sh.getRange(row,(idx["ACTIVO"]!==undefined?idx["ACTIVO"]:hdr.length-1)+1).setValue("NO");
  return {ok:true,message:"CDT cancelado (fila "+row+")."};
}

function apiCDTCheckFecha(payload) {
  var token = String((payload && payload.token) || "");
  requireSession_(token);
  var sede  = String((payload && payload.sede)  || "").trim();
  var id    = String((payload && payload.id)    || "").trim();
  var fecha = String((payload && payload.fecha) || todayISO_()).trim();
  if (!sede||!id||!isoValid_(fecha)) return {ok:true,tiene:false};
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CDT);
  if (!sh||sh.getLastRow()<2) return {ok:true,tiene:false};
  var v=sh.getDataRange().getValues(); var hdr=v[0].map(normHdr_); var idx={}; hdr.forEach(function(h,i){if(h)idx[h]=i;});
  for(var i=1;i<v.length;i++){
    var rs=String(v[i][idx["SEDE"]||2]||"").trim(); var ri=formatCellValue_(v[i][idx["ID"]||3]);
    var act=String(v[i][idx["ACTIVO"]!==undefined?idx["ACTIVO"]:11]||"").toUpperCase();
    var fi=String(v[i][idx["FECHA_INICIO"]||0]||"").trim(); var ff=String(v[i][idx["FECHA_FIN"]||1]||"").trim();
    var dT=String(v[i][idx["DIA_DESCANSO_TEMP"]||6]||"").trim().toUpperCase();
    var dO=String(v[i][idx["DIA_DESCANSO_ORIG"]||5]||"").trim().toUpperCase();
    if(rs!==sede||ri!==id) continue; if(act!=="SI") continue; if(fecha<fi||fecha>ff) continue;
    return {ok:true,tiene:true,diaTemp:dT,diaOrig:dO,fechaInicio:fi,fechaFin:ff};
  }
  return {ok:true,tiene:false};
}

function apiCDTGetPorSede(payload) {
  var token = String((payload && payload.token) || "");
  requireSession_(token);
  var sede  = String((payload && payload.sede)  || "").trim();
  var fecha = String((payload && payload.fecha) || todayISO_()).trim();
  if (!sede||!isoValid_(fecha)) return {ok:true,mapa:{}};
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CDT);
  if (!sh||sh.getLastRow()<2) return {ok:true,mapa:{}};
  var v=sh.getDataRange().getValues(); var hdr=v[0].map(normHdr_); var idx={}; hdr.forEach(function(h,i){if(h)idx[h]=i;});
  var mapa={};
  for(var i=1;i<v.length;i++){
    var rs=String(v[i][idx["SEDE"]!==undefined?idx["SEDE"]:2]||"").trim(); var ri=formatCellValue_(v[i][idx["ID"]!==undefined?idx["ID"]:3]);
    var act=String(v[i][idx["ACTIVO"]!==undefined?idx["ACTIVO"]:11]||"").toUpperCase();
    var fi=String(v[i][idx["FECHA_INICIO"]!==undefined?idx["FECHA_INICIO"]:0]||"").trim();
    var ff=String(v[i][idx["FECHA_FIN"]!==undefined?idx["FECHA_FIN"]:1]||"").trim();
    var dT=String(v[i][idx["DIA_DESCANSO_TEMP"]!==undefined?idx["DIA_DESCANSO_TEMP"]:6]||"").toUpperCase();
    var dO=String(v[i][idx["DIA_DESCANSO_ORIG"]!==undefined?idx["DIA_DESCANSO_ORIG"]:5]||"").toUpperCase();
    if(rs!==sede||!ri) continue; if(act!=="SI") continue; if(fecha<fi||fecha>ff) continue;
    mapa[ri]={diaTemp:dT,diaOrig:dO,fechaInicio:fi,fechaFin:ff};
  }
  return {ok:true,mapa:mapa};
}


// ══════════════════════════════════════════════════════════════
//  MÓDULO: REGISTROS DEL SUPERVISOR
// ══════════════════════════════════════════════════════════════
function apiGetRegistros(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  var capt  = String(ses.username || "").toLowerCase().trim();
  var fechaInicio = String((payload && payload.fechaInicio) || "").trim();
  var fechaFin    = String((payload && payload.fechaFin)    || todayISO_()).trim();
  var sede        = String((payload && payload.sede)        || "").trim();
  var soloMios    = !(String(ses.role || "").toUpperCase() === "SUPERADMIN");
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var out = { asistencias:[], incidencias:[], compensaciones:[], cdts:[], turnos:[], resumen:{} };
  var shPL = ss.getSheetByName(SH_PL);
  if (shPL && shPL.getLastRow() >= 2) {
    var plV = shPL.getDataRange().getValues();
    for (var i=1;i<plV.length;i++) {
      var fecha=formatCellValue_(plV[i][0]); var rowSd=String(plV[i][2]||"").trim();
      var codigo=String(plV[i][6]||"").trim().toUpperCase(); var cap=String(plV[i][7]||"").toLowerCase().trim();
      if(fechaInicio&&fecha<fechaInicio) continue; if(fechaFin&&fecha>fechaFin) continue;
      if(sede&&rowSd!==sede) continue; if(soloMios&&cap!==capt) continue;
      out.asistencias.push({fecha:fecha,sede:rowSd,id:formatCellValue_(plV[i][4]),nombre:String(plV[i][5]||"").trim(),codigo:codigo,ts:String(plV[i][8]||"").trim()});
    }
  }
  var shInc = ss.getSheetByName("INCIDENCIAS");
  if (shInc && shInc.getLastRow() >= 2) {
    var incV=shInc.getDataRange().getValues(); var incH=incV[0].map(normHdr_); var incI={};
    incH.forEach(function(h,i){if(h)incI[h]=i;});
    function sg(row,col){return col!==undefined?String(row[col]||"").trim():"";}
    for (var ii=1;ii<incV.length;ii++) {
      var iF=formatCellValue_(incV[ii][incI["FECHA_ISO"]!==undefined?incI["FECHA_ISO"]:0]);
      var iS=sg(incV[ii],incI["SEDE"]); var iCp=sg(incV[ii],incI["CAPTURADO_POR"]).toLowerCase();
      if(fechaInicio&&iF<fechaInicio) continue; if(fechaFin&&iF>fechaFin) continue;
      if(sede&&iS!==sede) continue; if(soloMios&&iCp!==capt) continue;
      out.incidencias.push({fecha:iF,sede:iS,id:formatCellValue_(incV[ii][incI["ID"]!==undefined?incI["ID"]:2]),
        nombre:sg(incV[ii],incI["NOMBRE"]),tipo:sg(incV[ii],incI["TIPO"]).toUpperCase(),
        obs:sg(incV[ii],incI["OBSERVACIONES"]),autoriza:sg(incV[ii],incI["AUTORIZA"]),
        captPor:sg(incV[ii],incI["CAPTURADO_POR"]),ts:sg(incV[ii],incI["TS_ISO"])});
    }
  }
  var codeCount={};
  out.asistencias.forEach(function(a){codeCount[a.codigo]=(codeCount[a.codigo]||0)+1;});
  out.resumen={totalAsistencias:out.asistencias.length,totalIncidencias:out.incidencias.length,codigosContados:codeCount};
  return {ok:true,data:out};
}


// ══════════════════════════════════════════════════════════════
//  MÓDULO: SOPORTE
// ══════════════════════════════════════════════════════════════
var SH_SOPORTE   = "SOPORTE_CHAT";
var SOPORTE_EMAIL= "edronemid@gmail.com";
var SOPORTE_EMAILS_EXTRA = ["contacto@mhsintegradora.mx"];
var HDR_SOPORTE  = ["TS_ISO","TOKEN_CHAT","USERNAME","FULL_NAME","SEDE","MENSAJE","TIPO","LEIDO"];

function _ensureSoporteSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_SOPORTE);
  if (!sh) {
    sh = ss.insertSheet(SH_SOPORTE);
    sh.getRange(1,1,1,HDR_SOPORTE.length).setValues([HDR_SOPORTE])
      .setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1); sh.autoResizeColumns(1,HDR_SOPORTE.length);
  }
  return sh;
}

function apiEnviarSoporte(payload) {
  var token   = String((payload && payload.token)   || "");
  var ses     = requireSession_(token);
  var mensaje = String((payload && payload.mensaje) || "").trim().substring(0,1000);
  var sede    = String((payload && payload.sede)    || "").trim();
  var chatId  = String((payload && payload.chatId)  || Utilities.getUuid().replace(/-/g,"").substring(0,12)).trim();
  var tipo    = String((payload && payload.tipo)    || "CONSULTA").trim().toUpperCase();
  if (!mensaje) throw new Error("Mensaje vacío.");
  var users = readUsers_(); var usuario = users.find(function(u){return u.username===ses.username;});
  var fullName = usuario ? (usuario.fullName||ses.username) : ses.username;
  var ts = isoNow_(); var sh = _ensureSoporteSheet_();
  sh.appendRow([ts, chatId, ses.username, fullName, sede, mensaje, tipo, "NO"]);
  try {
    var asunto = "[MHS Soporte] " + fullName + " · " + (sede||"Sin sede") + " · " + ts.substring(0,16);
    var cuerpo = "De: " + ses.username + "\nNombre: " + fullName + "\nSede: " + (sede||"N/A") + "\nFecha: " + ts + "\n\n" + mensaje;
    MailApp.sendEmail(SOPORTE_EMAIL, asunto, cuerpo);
  } catch(e){ Logger.log("Email soporte: "+e); }
  try {
    apiLogLiveEvent({token:token,tipo:"TICKET_NUEVO",sede:sede,jornada:"",mensaje:fullName+" envió una consulta de soporte · "+(sede||"Sin sede"),meta:{chatId:chatId,tipo:tipo}});
  } catch(e){ Logger.log("Live soporte: "+e); }
  return {ok:true,chatId:chatId,ts:ts,message:"Mensaje enviado."};
}

function apiGetSoporteHistorial(payload) {
  var token  = String((payload && payload.token)  || "");
  var ses    = requireSession_(token);
  var chatId = String((payload && payload.chatId) || "").trim();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SOPORTE);
  if (!sh||sh.getLastRow()<2) return {ok:true,items:[],chatId:chatId};
  var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={};
  h.forEach(function(hh,i){if(hh)idx[hh]=i;});
  var isAdmin=isAdminRole_(ses.role)||isSoporteRole_(ses.role)||String(ses.role).toUpperCase()==="SUPERADMIN";
  var out=[];
  for(var i=1;i<v.length;i++){
    var rowUser=String(v[i][idx["USERNAME"]!==undefined?idx["USERNAME"]:2]||"").toLowerCase().trim();
    var rowChat=String(v[i][idx["TOKEN_CHAT"]!==undefined?idx["TOKEN_CHAT"]:1]||"").trim();
    if(!isAdmin&&rowUser!==ses.username) continue;
    if(chatId&&rowChat!==chatId) continue;
    out.push({ts:String(v[i][idx["TS_ISO"]!==undefined?idx["TS_ISO"]:0]||"").trim(),chatId:rowChat,
      username:rowUser,fullName:String(v[i][idx["FULL_NAME"]!==undefined?idx["FULL_NAME"]:3]||"").trim(),
      sede:String(v[i][idx["SEDE"]!==undefined?idx["SEDE"]:4]||"").trim(),
      mensaje:String(v[i][idx["MENSAJE"]!==undefined?idx["MENSAJE"]:5]||"").trim(),
      tipo:String(v[i][idx["TIPO"]!==undefined?idx["TIPO"]:6]||"").trim().toUpperCase(),
      leido:String(v[i][idx["LEIDO"]!==undefined?idx["LEIDO"]:7]||"NO").toUpperCase()==="SI"});
  }
  out.sort(function(a,b){return a.ts<b.ts?-1:a.ts>b.ts?1:0;});
  return {ok:true,items:out,chatId:chatId};
}

function isSoporteRole_(role) {
  var r = String(role||"").toUpperCase();
  return r==="SOPORTE"||r==="SUPERADMIN"||r==="ADMIN";
}

function apiGetInboxSoporte(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN.");
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SOPORTE);
  if (!sh||sh.getLastRow()<2) return {ok:true,chats:[],totalUnread:0};
  var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={};
  h.forEach(function(hh,i){if(hh)idx[hh]=i;});
  function sg(row,col){return String(row[col!==undefined?col:99]||"").trim();}
  var chatsMap={};
  for(var i=1;i<v.length;i++){
    var chatId=sg(v[i],idx["TOKEN_CHAT"]); var ts=sg(v[i],idx["TS_ISO"]);
    var username=sg(v[i],idx["USERNAME"]); var fullName=sg(v[i],idx["FULL_NAME"]);
    var sede=sg(v[i],idx["SEDE"]); var mensaje=sg(v[i],idx["MENSAJE"]);
    var tipo=sg(v[i],idx["TIPO"]).toUpperCase(); var leido=sg(v[i],idx["LEIDO"]).toUpperCase()==="SI";
    if(!chatId) continue;
    if(!chatsMap[chatId]) chatsMap[chatId]={chatId:chatId,username:username,fullName:fullName,sede:sede,lastMsg:mensaje,lastTs:ts,lastTipo:tipo,unread:0,totalMsgs:0};
    if(ts>=chatsMap[chatId].lastTs){chatsMap[chatId].lastMsg=mensaje;chatsMap[chatId].lastTs=ts;chatsMap[chatId].lastTipo=tipo;if(username)chatsMap[chatId].username=username;if(fullName)chatsMap[chatId].fullName=fullName;}
    chatsMap[chatId].totalMsgs++;
    if(tipo==="CONSULTA"&&!leido) chatsMap[chatId].unread++;
  }
  var chats=Object.values(chatsMap).sort(function(a,b){return a.lastTs<b.lastTs?1:-1;});
  return {ok:true,chats:chats,totalUnread:chats.reduce(function(acc,c){return acc+c.unread;},0)};
}

function apiGetPendienteSoporte(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isSoporteRole_(ses.role)) return {ok:true,count:0,latestTs:""};
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SOPORTE);
  if (!sh||sh.getLastRow()<2) return {ok:true,count:0,latestTs:""};
  var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={}; h.forEach(function(hh,i){if(hh)idx[hh]=i;});
  var count=0,latestTs="";
  for(var i=1;i<v.length;i++){
    var tipo=String(v[i][idx["TIPO"]!==undefined?idx["TIPO"]:6]||"").toUpperCase();
    var leido=String(v[i][idx["LEIDO"]!==undefined?idx["LEIDO"]:7]||"NO").toUpperCase();
    var ts=String(v[i][idx["TS_ISO"]!==undefined?idx["TS_ISO"]:0]||"").trim();
    if(tipo==="CONSULTA"&&leido!=="SI"){count++;if(ts>latestTs)latestTs=ts;}
  }
  return {ok:true,count:count,latestTs:latestTs};
}

function apiResponderSoporte(payload) {
  var token   = String((payload && payload.token)   || "");
  var ses     = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN.");
  var chatId  = String((payload && payload.chatId)  || "").trim();
  var mensaje = String((payload && payload.mensaje) || "").trim().substring(0,1000);
  if (!chatId||!mensaje) throw new Error("chatId y mensaje requeridos.");
  var users=readUsers_(); var agente=users.find(function(u){return u.username===ses.username;});
  var agenteNombre=agente?(agente.fullName||ses.username):ses.username;
  var ts=isoNow_(); var sh=_ensureSoporteSheet_();
  sh.appendRow([ts,chatId,ses.username,agenteNombre,"SOPORTE",mensaje,"RESPUESTA","SI"]);
  var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={}; h.forEach(function(hh,i){if(hh)idx[hh]=i;});
  var userEmail="",userFullName="";
  for(var i=1;i<v.length;i++){
    var rowChat=String(v[i][idx["TOKEN_CHAT"]!==undefined?idx["TOKEN_CHAT"]:1]||"").trim();
    var rowTipo=String(v[i][idx["TIPO"]!==undefined?idx["TIPO"]:6]||"").toUpperCase();
    var rowUser=String(v[i][idx["USERNAME"]!==undefined?idx["USERNAME"]:2]||"").trim();
    if(rowChat!==chatId) continue;
    if(rowTipo==="CONSULTA"){
      sh.getRange(i+1,(idx["LEIDO"]!==undefined?idx["LEIDO"]:7)+1).setValue("SI");
      if(!userFullName) userFullName=String(v[i][idx["FULL_NAME"]!==undefined?idx["FULL_NAME"]:3]||"").trim();
      if(!userEmail&&rowUser){var fu=users.find(function(u){return u.username===rowUser;});if(fu&&fu.email)userEmail=fu.email;}
    }
  }
  try {
    if(userEmail) MailApp.sendEmail(userEmail,"[MHS Soporte] Respuesta · Chat "+chatId.substring(0,8),"Hola "+userFullName+",\n\nRespuesta del equipo de soporte:\n\n"+mensaje+"\n\n— Soporte MHS");
    MailApp.sendEmail(SOPORTE_EMAIL,"[MHS] ✅ Respuesta enviada · "+agenteNombre,"Chat: "+chatId+"\n\n"+mensaje);
  } catch(e){Logger.log("Email respuesta: "+e);}
  return {ok:true,ts:ts,chatId:chatId,message:"Respuesta enviada."};
}

function apiMarcarLeidoSoporte(payload) {
  var token  = String((payload && payload.token)  || "");
  var ses    = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Sin permisos.");
  var chatId = String((payload && payload.chatId) || "").trim();
  if (!chatId) throw new Error("chatId requerido.");
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SOPORTE);
  if (!sh||sh.getLastRow()<2) return {ok:true};
  var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={}; h.forEach(function(hh,i){if(hh)idx[hh]=i;});
  for(var i=1;i<v.length;i++){
    if(String(v[i][idx["TOKEN_CHAT"]!==undefined?idx["TOKEN_CHAT"]:1]||"").trim()===chatId)
      sh.getRange(i+1,(idx["LEIDO"]!==undefined?idx["LEIDO"]:7)+1).setValue("SI");
  }
  return {ok:true};
}


// ══════════════════════════════════════════════════════════════
//  RH PRO EXPORT + DASHBOARD
// ══════════════════════════════════════════════════════════════
var CODE_COLORS_RHPRO = {
  A:{bg:"#C6EFCE",fg:"#276221"},AF:{bg:"#D9EAD3",fg:"#276221"},F:{bg:"#F4CCCC",fg:"#990000"},
  PSG:{bg:"#FCE5CD",fg:"#783F04"},PCG:{bg:"#CFE2F3",fg:"#1155CC"},I:{bg:"#FFF2CC",fg:"#7F6000"},
  DT:{bg:"#D0E0E3",fg:"#134F5C"},"S/N":{bg:"#EFEFEF",fg:"#999999"},DS:{bg:"#EAD1DC",fg:"#4A1942"},
  FER:{bg:"#D9D2E9",fg:"#20124D"},INH:{bg:"#EFEFEF",fg:"#666666"}
};

function apiRHProExport(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN.");
  var fechas=[];
  if(Array.isArray(payload&&payload.fechas)&&payload.fechas.length) fechas=payload.fechas;
  else { var f=String((payload&&payload.fecha)||todayISO_()).trim(); fechas=[f]; }
  fechas=_uniqueKeepOrder_(fechas).filter(isoValid_);
  if(!fechas.length) throw new Error("No hay fechas válidas.");
  var sede=String((payload&&payload.sede)||"").trim();
  var label=String((payload&&payload.label)||fechas[0]).trim();
  var tipo=String((payload&&payload.tipo)||(fechas.length>1?"semana":"dia"));
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var matrix=buildAttendanceMatrix_({fechas:fechas,sede:sede});
  var tipoLabel=tipo==="dia"?"DIA":(tipo==="semana"?"SEM":(tipo==="quincena"?"Q":"MES"));
  var shName="RHPRO_"+tipoLabel+"_"+label.replace(/[^a-zA-Z0-9\-]/g,"").substring(0,18);
  var old=ss.getSheetByName(shName); if(old) ss.deleteSheet(old);
  var sh=ss.insertSheet(shName);
  sh.getRange(1,1).setValue("MHS RH PRO — Seguimiento de asistencias").setFontSize(13).setFontWeight("bold").setFontColor("#1a3a6c");
  sh.getRange(1,3).setValue("Período: "+label).setFontColor("#444");
  sh.getRange(1,4).setValue("Sede: "+(sede||"Todas")).setFontColor("#444");
  sh.getRange(1,5).setValue("Generado: "+isoNow_()).setFontColor("#888");
  var headers=["ID","NOMBRE","SEDE","JORNADA"].concat(fechas).concat(["TOTAL A","TOTAL F","TOTAL S/N","% CUMPL."]);
  sh.getRange(2,1,1,headers.length).setValues([headers]).setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sh.setFrozenRows(2);
  var SN_BG="#EFEFEF",SN_FG="#999999";
  var dataRows=[],bgMatrix=[],fgMatrix=[],textMatrix=[];
  for(var ri=0;ri<matrix.empleados.length;ri++){
    var emp=matrix.empleados[ri]; var codes=matrix.codeMatrix[ri]||[];
    var cntA=0,cntF=0,cntSN=0; var bgRow=[],fgRow=[],txtRow=[];
    for(var ci=0;ci<fechas.length;ci++){
      var cod=codes[ci]||"";
      if(cod===""){cntSN++;bgRow.push(SN_BG);fgRow.push(SN_FG);txtRow.push("S/N");}
      else{if(cod==="A"||cod==="AF")cntA++;else if(cod==="F")cntF++;var col=CODE_COLORS_RHPRO[cod]||{bg:"#fff",fg:"#333"};bgRow.push(col.bg);fgRow.push(col.fg);txtRow.push(cod);}
    }
    bgMatrix.push(bgRow);fgMatrix.push(fgRow);textMatrix.push(txtRow);
    var pct=fechas.length>0?Math.round((cntA/fechas.length)*100):0;
    dataRows.push([emp.id,emp.nombre,emp.sede,emp.turno||""].concat(txtRow).concat([cntA,cntF,cntSN,pct+"%"]));
  }
  if(dataRows.length){
    sh.getRange(3,1,dataRows.length,headers.length).setValues(dataRows);
    sh.getRange(3,5,dataRows.length,fechas.length).setBackgrounds(bgMatrix).setFontColors(fgMatrix).setHorizontalAlignment("center").setFontWeight("bold");
    sh.getRange(3,1,dataRows.length,4).setBackground("#ffffff").setFontColor("#1a1a1a");
  } else { sh.getRange(3,1).setValue("Sin empleados para este período/sede."); }
  sh.setColumnWidth(1,55);sh.setColumnWidth(2,240);sh.setColumnWidth(3,200);sh.setColumnWidth(4,100);
  for(var dc=5;dc<5+fechas.length;dc++) sh.setColumnWidth(dc,44);
  return {ok:true,message:"RH PRO: "+shName,url:ss.getUrl()+"#gid="+sh.getSheetId(),totalEmps:matrix.empleados.length};
}

function apiRHProDashboard(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN.");
  var fechas=[];
  if(Array.isArray(payload&&payload.fechas)&&payload.fechas.length) fechas=payload.fechas;
  else { var f=String((payload&&payload.fecha)||todayISO_()).trim(); fechas=[f]; }
  fechas=_uniqueKeepOrder_(fechas).filter(isoValid_);
  if(!fechas.length) return {ok:true,supervisores:[],resumen:{}};
  var sede=String((payload&&payload.sede)||"").trim();
  var matrix=buildAttendanceMatrix_({fechas:fechas,sede:sede});
  var totalSN=0; matrix.codeMatrix.forEach(function(row){row.forEach(function(c){if(!c)totalSN++;});});
  var totalCeldas=matrix.empleados.length*fechas.length;
  return {ok:true,supervisores:[],resumen:{totalEmpleados:matrix.empleados.length,totalFechas:fechas.length,totalCeldas:totalCeldas,totalSN:totalSN,pctCumplimiento:totalCeldas>0?Math.round(((totalCeldas-totalSN)/totalCeldas)*100):100}};
}


// ══════════════════════════════════════════════════════════════
//  GEOLOCALIZACIÓN + NÓMINA
// ══════════════════════════════════════════════════════════════
function reverseGeocode_(lat, lng) {
  try {
    var latN=parseFloat(lat),lngN=parseFloat(lng);
    if(isNaN(latN)||isNaN(lngN)) return "";
    var geocoder=Maps.newGeocoder().reverseGeocode(latN,lngN);
    var results=geocoder.results;
    if(!results||!results.length) return latN+","+lngN;
    var preferred=null,fallback=null;
    for(var i=0;i<results.length;i++){
      var types=results[i].types||[];
      if(!fallback) fallback=results[i];
      if(types.indexOf("route")>=0||types.indexOf("neighborhood")>=0||types.indexOf("sublocality")>=0){preferred=results[i];break;}
    }
    var res=preferred||fallback;
    if(!res) return latN+","+lngN;
    var comps=res.address_components||[];
    var route="",neighborhood="",locality="",subloc="";
    comps.forEach(function(comp){
      var t=comp.types||[];
      if(t.indexOf("route")>=0)route=comp.short_name;
      if(t.indexOf("neighborhood")>=0)neighborhood=comp.short_name;
      if(t.indexOf("sublocality_level_1")>=0)subloc=comp.short_name;
      if(t.indexOf("locality")>=0)locality=comp.short_name;
    });
    var parts=[];
    if(route)parts.push(route);
    if(neighborhood)parts.push(neighborhood); else if(subloc)parts.push(subloc);
    if(locality)parts.push(locality);
    if(!parts.length)parts.push(res.formatted_address||(latN+","+lngN));
    return parts.join(", ");
  } catch(e){ return String(lat)+","+String(lng); }
}

function apiReverseGeocode(payload) {
  var token = String((payload && payload.token) || "");
  requireSession_(token);
  var lat=String((payload&&payload.lat)||"").trim();
  var lng=String((payload&&payload.lng)||"").trim();
  return {ok:true,lat:lat,lng:lng,ubicacion:reverseGeocode_(lat,lng)};
}

var PAGO_DIA_=315.04,PRIMA_DOM_=78.76,DESC_FALTA_=393.80;

function apiExportarNomina(payload) {
  var token=String((payload&&payload.token)||"");
  var ses=requireSession_(token);
  if(!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN.");
  var fechas=[];
  if(Array.isArray(payload&&payload.fechas)&&payload.fechas.length) fechas=payload.fechas;
  else { var f=String((payload&&payload.fecha)||todayISO_()).trim(); fechas=[f]; }
  fechas=_uniqueKeepOrder_(fechas).filter(isoValid_);
  if(!fechas.length) throw new Error("No hay fechas válidas.");
  var tipo=String((payload&&payload.tipo)||(fechas.length>1?"semana":"dia"));
  var sede=String((payload&&payload.sede)||"").trim();
  var label=String((payload&&payload.label)||fechas[0]).trim();
  var domingoSet={};
  fechas.forEach(function(f){var p=f.split("-");var d=new Date(Number(p[0]),Number(p[1])-1,Number(p[2]));if(d.getDay()===0)domingoSet[f]=true;});
  var matrix=buildAttendanceMatrix_({fechas:fechas,sede:sede});
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var tipoLabel=tipo==="dia"?"DIA":(tipo==="semana"?"SEM":"MES");
  var shName="NOM_"+tipoLabel+"_"+label.replace(/[^a-zA-Z0-9\-]/g,"").substring(0,18);
  var old=ss.getSheetByName(shName); if(old) ss.deleteSheet(old);
  var sh=ss.insertSheet(shName);
  sh.getRange(1,1).setValue("MHS Integradora · Nómina estimada").setFontSize(13).setFontWeight("bold").setFontColor("#1a3a6c");
  sh.getRange(1,3).setValue("Período: "+label).setFontColor("#444");
  sh.getRange(1,4).setValue("Sede: "+(sede||"Todas")).setFontColor("#444");
  sh.getRange(1,5).setValue("Generado: "+isoNow_()).setFontColor("#888");
  var nomHeaders=["DÍAS LAB.","DÍAS FALTA","DOMS. TRAB.","PRIMA DOM.","DESC. FALTAS","PAGO ESTIM."];
  var headers=["ID","NOMBRE","SEDE","JORNADA"].concat(fechas).concat(nomHeaders);
  sh.getRange(2,1,1,headers.length).setValues([headers]).setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sh.setFrozenRows(2);
  var dataRows=[],dispCodes=[];
  for(var ri=0;ri<matrix.empleados.length;ri++){
    var emp=matrix.empleados[ri]; var codes=matrix.codeMatrix[ri]||[];
    var diasLab=0,diasFalta=0,diasDom=0,rowCodes=[];
    for(var ci=0;ci<fechas.length;ci++){
      var cod=codes[ci]||""; rowCodes.push(cod||"S/N");
      var esAsist=(cod==="A"||cod==="AF"||cod==="DT");
      if(esAsist){diasLab++;if(domingoSet[fechas[ci]])diasDom++;}
      if(cod==="F")diasFalta++;
    }
    function fmt(n){return "$"+n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g,",");}
    var primaDom=diasDom*PRIMA_DOM_; var descFaltas=diasFalta*DESC_FALTA_;
    var pagoEstim=(diasLab*PAGO_DIA_)+primaDom-descFaltas;
    dataRows.push([emp.id,emp.nombre,emp.sede,emp.turno||""].concat(rowCodes).concat([diasLab,diasFalta,diasDom,fmt(primaDom),fmt(descFaltas),fmt(pagoEstim)]));
    dispCodes.push(rowCodes);
  }
  if(dataRows.length){
    sh.getRange(3,1,dataRows.length,headers.length).setValues(dataRows);
    paintCodeMatrix_(sh,3,5,dispCodes);
    sh.getRange(3,1,dataRows.length,4).setBackground("#ffffff").setFontColor("#1a1a1a");
  } else { sh.getRange(3,1).setValue("Sin empleados para este período/sede."); }
  sh.setColumnWidth(1,55);sh.setColumnWidth(2,240);sh.setColumnWidth(3,190);sh.setColumnWidth(4,100);
  for(var dc=5;dc<5+fechas.length;dc++) sh.setColumnWidth(dc,44);
  return {ok:true,message:"Nómina '"+shName+"' creada.",url:ss.getUrl()+"#gid="+sh.getSheetId(),totalEmps:matrix.empleados.length,domingosPeriodo:Object.keys(domingoSet).length};
}




// ══════════════════════════════════════════════════════════════
//  LIVE CENTER V4 — Alertas internas en tiempo real ligero
//  Capa 1: funciona dentro de Apps Script/PWA abierta.
// ══════════════════════════════════════════════════════════════
function isLiveViewerRole_(role) {
  var r = String(role || "").toUpperCase().trim();
  return r === "SOPORTE" || r === "SUPERADMIN";
}

function _ensureLiveEventsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_LIVE_EVENTS);
  if (!sh) {
    sh = ss.insertSheet(SH_LIVE_EVENTS);
    sh.getRange(1, 1, 1, HDR_LIVE.length).setValues([HDR_LIVE]);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, HDR_LIVE.length);
    return sh;
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HDR_LIVE.length).setValues([HDR_LIVE]);
    sh.setFrozenRows(1);
    return sh;
  }
  var existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(normHdr_);
  HDR_LIVE.forEach(function(h){
    if (existing.indexOf(normHdr_(h)) < 0) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      existing.push(normHdr_(h));
    }
  });
  return sh;
}

function _liveUserName_(username) {
  try {
    var u = readUsers_().find(function(x){ return String(x.username||"").toLowerCase() === String(username||"").toLowerCase(); });
    return u ? (u.fullName || u.username || username) : username;
  } catch(e) {
    return username;
  }
}

function apiLogLiveEvent(payload) {
  payload = payload || {};
  var token = String(payload.token || "").trim();
  var ses = requireSession_(token);

  var tipo    = String(payload.tipo || payload.type || "INFO").toUpperCase().trim();
  var sede    = String(payload.sede || "").trim();
  var jornada = String(payload.jornada || "").trim();
  var meta    = payload.meta || {};
  var full    = _liveUserName_(ses.username);
  var role    = String(ses.role || "").toUpperCase().trim();
  var ts      = isoNow_();
  var fecha   = todayISO_();

  var mensaje = String(payload.mensaje || "").trim();
  if (!mensaje) {
    if (tipo === "PASE_INICIADO") {
      mensaje = full + " empezó a llenar " + (sede || "una sede") + (jornada ? " · " + jornada : "");
    } else if (tipo === "PASE_GUARDADO") {
      var total = meta && meta.total ? (" · " + meta.total + " registros") : "";
      mensaje = full + " guardó asistencia de " + (sede || "una sede") + (jornada ? " · " + jornada : "") + total;
    } else if (tipo === "TICKET_NUEVO") {
      mensaje = full + " envió un nuevo ticket" + (sede ? " · " + sede : "");
    } else {
      mensaje = full + " generó una actualización";
    }
  }

  var sh = _ensureLiveEventsSheet_();
  sh.appendRow([
    ts, fecha, ses.username, full, role, tipo, sede, jornada,
    mensaje, JSON.stringify(meta || {}), "SOPORTE,SUPERADMIN", ""
  ]);

  // Mantener la hoja ligera: conservar últimas 1000 filas.
  try {
    var maxRows = 1000;
    var lr = sh.getLastRow();
    if (lr > maxRows + 1) sh.deleteRows(2, lr - maxRows - 1);
  } catch(e) {}

  return { ok:true, ts:ts, mensaje:mensaje };
}

function apiGetLiveEvents(payload) {
  payload = payload || {};
  var token = String(payload.token || "").trim();
  var ses = requireSession_(token);
  if (!isLiveViewerRole_(ses.role)) return { ok:true, events:[], lastTs:String(payload.sinceTs||"") };

  var since = String(payload.sinceTs || payload.since || "").trim();
  var limit = Math.max(1, Math.min(Number(payload.limit || 50), 100));
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_LIVE_EVENTS);
  if (!sh || sh.getLastRow() < 2) return { ok:true, events:[], lastTs:since };

  var v = sh.getDataRange().getValues();
  var h = v[0].map(normHdr_);
  var idx = {};
  h.forEach(function(x,i){ if(x) idx[x]=i; });
  function g(row, name) {
    var i = idx[normHdr_(name)];
    return i === undefined ? "" : formatCellValue_(row[i]);
  }

  var out = [];
  var lastTs = since;
  for (var i = v.length - 1; i >= 1 && out.length < limit; i--) {
    var row = v[i];
    var ts = g(row, "TS_ISO");
    if (since && ts <= since) break;
    var username = g(row, "USERNAME");
    // No regresar eventos creados por el mismo usuario del visor.
    if (String(username).toLowerCase() === String(ses.username).toLowerCase()) {
      if (!lastTs || ts > lastTs) lastTs = ts;
      continue;
    }
    var ev = {
      ts: ts,
      fecha: g(row, "FECHA_ISO"),
      username: username,
      fullName: g(row, "FULL_NAME"),
      role: g(row, "ROLE"),
      tipo: g(row, "TIPO"),
      sede: g(row, "SEDE"),
      jornada: g(row, "JORNADA"),
      mensaje: g(row, "MENSAJE")
    };
    out.push(ev);
    if (!lastTs || ts > lastTs) lastTs = ts;
  }
  out.reverse();
  return { ok:true, events:out, lastTs:lastTs || since };
}


// ══════════════════════════════════════════════════════════════
//  TICKETS MEJORADOS
// ══════════════════════════════════════════════════════════════
function getSoporteEmailRecipients_() {
  var lista=[SOPORTE_EMAIL];
  if(Array.isArray(SOPORTE_EMAILS_EXTRA)) SOPORTE_EMAILS_EXTRA.forEach(function(e){if(e)lista.push(e);});
  try { readUsers_().forEach(function(u){ if(u.active&&(u.role==="SOPORTE"||u.role==="SUPERADMIN")&&u.email) lista.push(u.email); }); } catch(e){}
  var seen={}; return lista.filter(function(e){var k=String(e).toLowerCase().trim();if(!k||seen[k])return false;seen[k]=true;return true;});
}

function esc_(s) {
  return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function apiEnviarTicket(payload) {
  var token       = String((payload&&payload.token)         ||"");
  var ses         = requireSession_(token);
  var mensaje     = String((payload&&payload.mensaje)       ||"").trim().substring(0,2000);
  var sede        = String((payload&&payload.sede)          ||"").trim();
  var chatId      = String((payload&&payload.chatId)        ||Utilities.getUuid().replace(/-/g,"").substring(0,12)).trim();
  var tipo_ticket = String((payload&&payload.tipo_ticket)   ||"CONSULTA").trim().toUpperCase();
  var fecha_sol   = String((payload&&payload.fecha_solicitud)||"").trim();
  if(!mensaje) throw new Error("El mensaje no puede estar vacío.");
  var users=readUsers_(); var usuario=users.find(function(u){return u.username===ses.username;});
  var fullName=usuario?(usuario.fullName||ses.username):ses.username;
  var sedeEff=sede||(usuario&&usuario.sede)||"No especificada";
  var ts=isoNow_();
  var sh=_ensureSoporteSheet_();
  sh.appendRow([ts,chatId,ses.username,fullName,sedeEff,mensaje,tipo_ticket,"NO"]);
  var recipients=getSoporteEmailRecipients_();
  var iconos={CONSULTA:"💬",DESBLOQUEO_FECHA:"🔓",INCIDENCIA_URGENTE:"🚨",OTRO:"📋"};
  var icono=iconos[tipo_ticket]||"📋";
  var asunto=icono+" [MHS Ticket] "+tipo_ticket+" · "+fullName+" · "+sedeEff+" · "+ts.substring(0,16);
  var extra=tipo_ticket==="DESBLOQUEO_FECHA"&&fecha_sol?"Fecha a desbloquear: "+fecha_sol+"\n":"";
  var cuerpo="MHS RH HUB · Nuevo Ticket\n━━━━━━━━━━━━━━━━━━━\nTipo: "+tipo_ticket+"\nUsuario: "+ses.username+"\nNombre: "+fullName+"\nSede: "+sedeEff+(usuario&&usuario.jornada?"\nJornada: "+usuario.jornada:"")+"\nFecha: "+ts+"\nChat ID: "+chatId+(extra?"\n"+extra:"")+"\n━━━━━━━━━━━━━━━━━━━\n\n"+mensaje;
  try { recipients.forEach(function(rcpt){MailApp.sendEmail({to:rcpt,subject:asunto,body:cuerpo});}); } catch(e){Logger.log("Ticket email: "+e);}
  try {
    apiLogLiveEvent({
      token: token,
      tipo: "TICKET_NUEVO",
      sede: sedeEff,
      jornada: (usuario && usuario.jornada) || "",
      mensaje: fullName + " envió un ticket " + tipo_ticket + " · " + sedeEff,
      meta: { chatId: chatId, tipo_ticket: tipo_ticket, fecha_solicitud: fecha_sol }
    });
  } catch(e){ Logger.log("Live ticket: "+e); }
  return {ok:true,chatId:chatId,ts:ts,message:"Ticket enviado. El equipo de soporte ha sido notificado."};
}

function apiCerrarTicket(payload) {
  var token  = String((payload&&payload.token) ||"");
  var ses    = requireSession_(token);
  if(!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN.");
  var chatId = String((payload&&payload.chatId)||"").trim();
  if(!chatId) throw new Error("chatId requerido.");
  var ts=isoNow_(); var sh=_ensureSoporteSheet_();
  var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={}; h.forEach(function(hh,i){if(hh)idx[hh]=i;});
  var userUsername="",userFullName="";
  for(var i=1;i<v.length;i++){
    var rowChat=String(v[i][idx["TOKEN_CHAT"]!==undefined?idx["TOKEN_CHAT"]:1]||"").trim();
    if(rowChat!==chatId) continue;
    var rowTipo=String(v[i][idx["TIPO"]!==undefined?idx["TIPO"]:6]||"").toUpperCase();
    var rowUser=String(v[i][idx["USERNAME"]!==undefined?idx["USERNAME"]:2]||"").trim();
    if(rowTipo==="CONSULTA"||rowTipo.indexOf("TICKET")>=0||rowTipo.indexOf("DESBLOQUEO")>=0){
      sh.getRange(i+1,(idx["LEIDO"]!==undefined?idx["LEIDO"]:7)+1).setValue("SI");
      if(!userFullName)userFullName=String(v[i][idx["FULL_NAME"]!==undefined?idx["FULL_NAME"]:3]||"").trim();
      if(!userUsername)userUsername=rowUser;
    }
  }
  var agenteUsers=readUsers_(); var agente=agenteUsers.find(function(u){return u.username===ses.username;});
  var agenteNombre=agente?(agente.fullName||ses.username):ses.username;
  sh.appendRow([ts,chatId,ses.username,agenteNombre,"SOPORTE","Ticket cerrado como RESUELTO por "+agenteNombre,"RESUELTO","SI"]);
  try {
    var supUser=agenteUsers.find(function(u){return u.username===userUsername;});
    if(supUser&&supUser.email) MailApp.sendEmail({to:supUser.email,subject:"[MHS] ✅ Ticket resuelto · "+chatId.substring(0,8),body:"Hola "+(userFullName||"Supervisor")+",\n\nTu ticket fue marcado como RESUELTO por "+agenteNombre+".\n\n— Soporte MHS"});
  } catch(e){Logger.log("Email cierre ticket: "+e);}
  return {ok:true,ts:ts,message:"Ticket cerrado."};
}

function apiGetMisTickets(payload) {
  var token = String((payload&&payload.token)||"");
  var ses   = requireSession_(token);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SOPORTE);
  if(!sh||sh.getLastRow()<2) return {ok:true,tickets:[]};
  var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={}; h.forEach(function(hh,i){if(hh)idx[hh]=i;});
  function sg(row,col){return String(row[col!==undefined?col:99]||"").trim();}
  var chatsMap={},chatOrder=[];
  for(var i=1;i<v.length;i++){
    var rowUser=sg(v[i],idx["USERNAME"]).toLowerCase();
    if(rowUser!==ses.username) continue;
    var chatId=sg(v[i],idx["TOKEN_CHAT"]);
    var ts=sg(v[i],idx["TS_ISO"]);
    var tipo=sg(v[i],idx["TIPO"]).toUpperCase();
    var mensaje=sg(v[i],idx["MENSAJE"]);
    if(!chatId) continue;
    if(!chatsMap[chatId]){chatsMap[chatId]={chatId:chatId,tipo:tipo,firstTs:ts,lastMsg:mensaje,ts:ts,respondido:false,resuelto:false};chatOrder.push(chatId);}
    if(ts>=chatsMap[chatId].ts){chatsMap[chatId].lastMsg=mensaje;chatsMap[chatId].ts=ts;}
    if(tipo==="RESPUESTA")chatsMap[chatId].respondido=true;
    if(tipo==="RESUELTO")chatsMap[chatId].resuelto=true;
  }
  var tickets=chatOrder.map(function(id){return chatsMap[id];}).sort(function(a,b){return a.firstTs<b.firstTs?1:-1;});
  return {ok:true,tickets:tickets};
}


// ══════════════════════════════════════════════════════════════
//  RH PRO — ASIGNACIONES MULTI-SEDE (CORRECCIÓN PRINCIPAL)
//  Una fila por sede+jornada por usuario en ASIGNACIONES_SUP
// ══════════════════════════════════════════════════════════════
var SH_ASIGN  = "ASIGNACIONES_SUP";
var HDR_ASIGN = ["USERNAME","FULL_NAME","SEDE","JORNADA","ACTIVO","TS_ISO"];

function _ensureAsignSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_ASIGN);
  if (!sh) {
    sh = ss.insertSheet(SH_ASIGN);
    sh.getRange(1,1,1,HDR_ASIGN.length).setValues([HDR_ASIGN])
      .setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
    Logger.log("Hoja ASIGNACIONES_SUP creada.");
  }
  return sh;
}

/**
 * apiRHProGetAsignaciones — MULTI-SEDE
 * Cada supervisor puede tener N filas en ASIGNACIONES_SUP (una por sede+jornada).
 * Devuelve supervisores[].asignaciones:[{sede,jornada,activo}]
 */
function apiRHProGetAsignaciones(payload) {
  var token = String((payload && payload.token) || "");
  var ses   = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede ver asignaciones.");

  var sh = _ensureAsignSheet_();
  var usuarios = readUsers_().filter(function(u) {
    return u.active && (u.role==="USER"||u.role==="ADMIN"||u.role==="SOPORTE"||u.role==="SUPERADMIN");
  });

  // Leer TODAS las filas — múltiples por usuario
  var asignMap = {};
  if (sh.getLastRow() >= 2) {
    var v = sh.getDataRange().getValues();
    var h = v[0].map(normHdr_); var idx = {};
    h.forEach(function(hh,i){if(hh)idx[hh]=i;});
    for (var i=1;i<v.length;i++) {
      var uname = String(v[i][idx["USERNAME"]||0]||"").toLowerCase().trim();
      var sede_ = String(v[i][idx["SEDE"]    ||2]||"").trim();
      var jorn_ = String(v[i][idx["JORNADA"] ||3]||"").trim().toUpperCase();
      var act_  = String(v[i][idx["ACTIVO"]  ||4]||"SI").toUpperCase()==="SI";
      if (!uname||!sede_||!jorn_) continue;
      if (!asignMap[uname]) asignMap[uname]=[];
      asignMap[uname].push({sede:sede_,jornada:jorn_,activo:act_,row:i+1});
    }
  }

  var sedes    = getSedes_();
  var jornadas = ["MATUTINO","VESPERTINO","NOCTURNO"];

  var list = usuarios.map(function(u) {
    return {
      username:     u.username,
      fullName:     u.fullName || u.username,
      role:         u.role,
      email:        u.email || "",
      asignaciones: asignMap[u.username] || []
    };
  });

  return { ok:true, supervisores:list, sedes:sedes.map(function(s){return s.sede;}), jornadas:jornadas };
}

/**
 * apiRHProGuardarAsignacion — MULTI-SEDE con action add|remove
 * La clave única es (username + sede + jornada).
 * action="add"    → inserta o actualiza esa combinación
 * action="remove" → elimina esa fila específica
 */
function apiRHProGuardarAsignacion(payload) {
  var token    = String((payload&&payload.token)   ||"");
  var ses      = requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN puede asignar sedes.");

  var username = String((payload&&payload.username)||"").toLowerCase().trim();
  var sede     = String((payload&&payload.sede)    ||"").trim();
  var jornada  = String((payload&&payload.jornada) ||"").trim().toUpperCase();
  var activo   = (payload&&payload.activo)!==false ? "SI" : "NO";
  var action   = String((payload&&payload.action)  ||"add").toLowerCase(); // add | remove

  if (!username||!sede||!jornada) throw new Error("username, sede y jornada son requeridos.");

  var ts=isoNow_(); var sh=_ensureAsignSheet_();
  var users=readUsers_(); var u=users.find(function(x){return x.username===username;});
  var fullName=u?(u.fullName||username):username;

  // Buscar fila existente con la misma clave (username+sede+jornada)
  var existingRow=null;
  if (sh.getLastRow()>=2) {
    var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={};
    h.forEach(function(hh,i){if(hh)idx[hh]=i;});
    for (var i=1;i<v.length;i++) {
      var ru=String(v[i][idx["USERNAME"]||0]||"").toLowerCase().trim();
      var rs=String(v[i][idx["SEDE"]    ||2]||"").trim();
      var rj=String(v[i][idx["JORNADA"] ||3]||"").trim().toUpperCase();
      if (ru===username&&rs===sede&&rj===jornada){existingRow=i+1;break;}
    }
  }

  if (action==="remove") {
    if (existingRow) sh.deleteRow(existingRow);
    return {ok:true,message:"Eliminado: "+username+" → "+sede+" / "+jornada};
  }

  // add / update
  if (existingRow) {
    sh.getRange(existingRow,1,1,HDR_ASIGN.length).setValues([[username,fullName,sede,jornada,activo,ts]]);
  } else {
    sh.appendRow([username,fullName,sede,jornada,activo,ts]);
  }
  return {ok:true,message:"Guardado: "+username+" → "+sede+" / "+jornada};
}

/**
 * apiGetSedeJornadaUsuario — MULTI-SEDE
 * Devuelve TODAS las asignaciones del usuario autenticado desde ASIGNACIONES_SUP.
 * Fallback a columnas SEDE/JORNADA de USUARIOS si no hay filas en ASIGNACIONES_SUP.
 */
function apiGetSedeJornadaUsuario(payload) {
  var token = String((payload&&payload.token)||"");
  var ses   = requireSession_(token);
  var asignaciones=[];

  try {
    var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ASIGN);
    if (sh&&sh.getLastRow()>=2) {
      var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={};
      h.forEach(function(hh,i){if(hh)idx[hh]=i;});
      for (var i=1;i<v.length;i++) {
        var uname=String(v[i][idx["USERNAME"]||0]||"").toLowerCase().trim();
        if (uname!==ses.username) continue;
        var sed =String(v[i][idx["SEDE"]   ||2]||"").trim();
        var jorn=String(v[i][idx["JORNADA"]||3]||"").trim().toUpperCase();
        var act =String(v[i][idx["ACTIVO"] ||4]||"SI").toUpperCase()==="SI";
        if (sed&&jorn&&act) asignaciones.push({sede:sed,jornada:jorn});
      }
    }
  } catch(e){}

  // Fallback: columnas SEDE/JORNADA en hoja USUARIOS
  if (!asignaciones.length) {
    var users=readUsers_();
    var u=users.find(function(x){return x.username===ses.username;});
    if (u&&u.sede&&u.jornada) asignaciones.push({sede:u.sede,jornada:u.jornada});
  }

  return {
    ok:           true,
    asignaciones: asignaciones,
    // Campos legacy para compatibilidad con código anterior
    sede:    asignaciones.length>0?asignaciones[0].sede   :"",
    jornada: asignaciones.length>0?asignaciones[0].jornada:"",
    email:   (function(){var users=readUsers_();var u=users.find(function(x){return x.username===ses.username;});return u?u.email:"";})()
  };
}

/**
 * apiGetPendientesAsistencia — MULTI-SEDE
 * Suma los empleados sin marcar de TODAS las sedes+jornadas asignadas al supervisor.
 * Devuelve resumen:[{sede,jornada,pendientes,total}] además del total global.
 */
function apiGetPendientesAsistencia(payload) {
  var token=String((payload&&payload.token)||"");
  var ses  =requireSession_(token);
  var fechaISO=String((payload&&payload.fecha)||todayISO_()).trim();

  // Obtener todas las asignaciones del usuario
  var asignRes=apiGetSedeJornadaUsuario({token:token});
  var asignaciones=asignRes.asignaciones||[];

  // Fallback: payload.sede (llamada legacy)
  if (!asignaciones.length&&payload&&payload.sede) {
    asignaciones=[{sede:String(payload.sede).trim(),jornada:String(payload.jornada||"").trim().toUpperCase()}];
  }
  if (!asignaciones.length) return {ok:true,pendientes:0,total:0,sedes:[],resumen:[]};

  var pl=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  var marcados={};
  if (pl&&pl.getLastRow()>=2) {
    var v=pl.getDataRange().getValues();
    for (var i=1;i<v.length;i++) {
      var f=formatCellValue_(v[i][0]);
      if (f!==fechaISO) continue;
      var rowSede=String(v[i][2]||"").trim();
      var id=formatCellValue_(v[i][4]);
      var cod=String(v[i][6]||"").trim();
      if (id&&cod) marcados[rowSede+"|"+id]=cod;
    }
  }

  var totalPend=0,totalEmps=0,resumen=[];
  asignaciones.forEach(function(asign){
    var todos=apiGetEmpleadosPorSede(asign.sede);
    if (asign.jornada) {
      todos=todos.filter(function(e){return String(e.turno||"").toUpperCase().indexOf(asign.jornada)>=0;});
    }
    var pend=todos.filter(function(e){return !marcados[asign.sede+"|"+e.id];}).length;
    totalPend+=pend; totalEmps+=todos.length;
    resumen.push({sede:asign.sede,jornada:asign.jornada,pendientes:pend,total:todos.length});
  });

  return {
    ok:true,pendientes:totalPend,total:totalEmps,
    sede:asignaciones[0].sede,jornada:asignaciones[0].jornada,
    fechaISO:fechaISO,
    sedes:asignaciones.map(function(a){return a.sede;}),
    resumen:resumen
  };
}

/**
 * apiRHProExportarAsignaciones — MULTI-SEDE
 * Exporta la tabla completa de asignaciones: una fila por sede+jornada por supervisor.
 */
function apiRHProExportarAsignaciones(payload) {
  var token=String((payload&&payload.token)||"");
  var ses  =requireSession_(token);
  if (!isAdminRole_(ses.role)) throw new Error("Solo ADMIN/SUPERADMIN.");

  var res=apiRHProGetAsignaciones(payload);
  var supervisores=res.supervisores||[];
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var shName="ASIGN_SUP_"+todayISO_().replace(/-/g,"");
  var old=ss.getSheetByName(shName); if(old) ss.deleteSheet(old);
  var sh=ss.insertSheet(shName);

  sh.getRange(1,1).setValue("MHS RH PRO · Asignaciones de Supervisores").setFontSize(13).setFontWeight("bold").setFontColor("#1a3a6c");
  sh.getRange(1,3).setValue("Generado: "+isoNow_()).setFontColor("#888");

  var headers=["SUPERVISOR","ROL","SEDE","JORNADA","EMAIL"];
  sh.getRange(2,1,1,headers.length).setValues([headers]).setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold");
  sh.setFrozenRows(2);

  var jornadaColors={MATUTINO:{bg:"#E8F5E9",fg:"#1B5E20"},VESPERTINO:{bg:"#E3F2FD",fg:"#0D47A1"},NOCTURNO:{bg:"#F3E5F5",fg:"#4A148C"}};
  var rows=[];
  supervisores.forEach(function(sup){
    if(!sup.asignaciones||!sup.asignaciones.length){
      rows.push([sup.fullName,sup.role,"—","—",sup.email||""]);
    } else {
      sup.asignaciones.forEach(function(a){
        rows.push([sup.fullName,sup.role,a.sede,a.jornada,sup.email||""]);
      });
    }
  });

  if (rows.length) {
    sh.getRange(3,1,rows.length,headers.length).setValues(rows);
    for(var r=0;r<rows.length;r++){
      var j=rows[r][3]; var col=jornadaColors[j];
      if(col) sh.getRange(3+r,4).setBackground(col.bg).setFontColor(col.fg).setFontWeight("bold");
    }
    sh.setColumnWidth(1,220);sh.setColumnWidth(2,90);sh.setColumnWidth(3,220);sh.setColumnWidth(4,110);sh.setColumnWidth(5,200);
  }

  var assigned=supervisores.filter(function(s){return s.asignaciones&&s.asignaciones.length>0;}).length;
  return {ok:true,message:"Hoja '"+shName+"' creada. "+assigned+" supervisores con asignaciones.",url:ss.getUrl()+"#gid="+sh.getSheetId()};
}

/**
 * apiGetEmpleadosPorSedeJornada
 * Versión filtrada por jornada de apiGetEmpleadosPorSede.
 */
function apiGetEmpleadosPorSedeJornada(payload) {
  var sede, jornada;
  if (typeof payload==="string"){sede=String(payload||"").trim();jornada="";}
  else {
    sede   =String((payload&&payload.sede)   ||"").trim();
    jornada=String((payload&&payload.jornada)||"").trim().toUpperCase();
    var token=String((payload&&payload.token)||"");
    if (token) {
      try {
        var ses=requireSession_(token);
        if (!jornada||!sede) {
          var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ASIGN);
          if (sh&&sh.getLastRow()>=2) {
            var v=sh.getDataRange().getValues(); var h=v[0].map(normHdr_); var idx={};
            h.forEach(function(hh,i){if(hh)idx[hh]=i;});
            for(var i=1;i<v.length;i++){
              if(String(v[i][idx["USERNAME"]||0]||"").toLowerCase().trim()===ses.username){
                if(!jornada) jornada=String(v[i][idx["JORNADA"]||3]||"").trim().toUpperCase();
                if(!sede)    sede   =String(v[i][idx["SEDE"]   ||2]||"").trim();
                break;
              }
            }
          }
        }
      } catch(e){}
    }
  }
  var todos=apiGetEmpleadosPorSede(sede);
  if(!jornada) return todos;
  return todos.filter(function(emp){return String(emp.turno||"").toUpperCase().indexOf(jornada)>=0;});
}

// ══════════════════════════════════════════════════════
//  OPTIMIZACIÓN V1 — CARGA RÁPIDA MOBILE / SHEETS
//  Agregado al final para no romper funciones existentes.
//  Estas funciones sobreescriben las anteriores con el mismo nombre.
// ══════════════════════════════════════════════════════
var __MHS_MEM__ = (typeof __MHS_MEM__ !== 'undefined') ? __MHS_MEM__ : {};
var MHS_CACHE_SEDES_TTL = 60;
var MHS_CACHE_USERS_TTL = 120;
function cacheGetJSON_(key){try{var raw=CacheService.getScriptCache().get(key);return raw?JSON.parse(raw):null;}catch(e){return null;}}
function cachePutJSON_(key,obj,ttl){try{CacheService.getScriptCache().put(key,JSON.stringify(obj),ttl||60);}catch(e){}}
function mhsGetSS_(){return SpreadsheetApp.getActiveSpreadsheet();}
function getContratosSnapshot_(){
  if(__MHS_MEM__.contratosSnap)return __MHS_MEM__.contratosSnap;
  var ss=mhsGetSS_(), sh=ss.getSheetByName(SH_CONTRATOS);
  if(!sh||sh.getLastRow()<2){__MHS_MEM__.contratosSnap={rows:[],idx:{},estadoPase:readEstadoPaseMap_()};return __MHS_MEM__.contratosSnap;}
  var raw=sh.getDataRange().getValues(); var hdr=raw[0].map(normHdr_); var idx={};
  hdr.forEach(function(h,i){if(h)idx[h]=i;});
  __MHS_MEM__.contratosSnap={rows:raw.slice(1),idx:idx,estadoPase:readEstadoPaseMap_()};
  return __MHS_MEM__.contratosSnap;
}
function readUsers_(){
  var cached=cacheGetJSON_('MHS:USERS:v1'); if(cached)return cached;
  var sh=mhsGetSS_().getSheetByName(SH_USERS); if(!sh||sh.getLastRow()<2)return [];
  var v=sh.getDataRange().getValues(); var idx={}; v[0].forEach(function(h,i){idx[normHdr_(h)]=i;});
  var out=[];
  for(var i=1;i<v.length;i++){var r=v[i];var activeRaw=String(r[idx.ACTIVE]||'').toLowerCase().trim();out.push({
    username:String(r[idx.USERNAME]||'').toLowerCase().trim(),password:String(r[idx.PASSWORD]||''),
    role:String(r[idx.ROLE]||'USER').toUpperCase().trim(),active:activeRaw==='true'||activeRaw==='si'||activeRaw==='sí'||activeRaw==='1',
    fullName:String(r[idx.FULL_NAME]||''),email:String((idx.EMAIL!==undefined?r[idx.EMAIL]:'')||''),
    sede:String((idx.SEDE!==undefined?r[idx.SEDE]:'')||'').trim(),jornada:String((idx.JORNADA!==undefined?r[idx.JORNADA]:'')||'').trim().toUpperCase()
  });}
  cachePutJSON_('MHS:USERS:v1',out,MHS_CACHE_USERS_TTL); return out;
}
function getSedes_(){
  var cached=cacheGetJSON_('MHS:SEDES:v2'); if(cached)return cached;
  var ss=mhsGetSS_(); var catMap={}; var shCat=ss.getSheetByName(SH_CAT_SEDES);
  if(shCat&&shCat.getLastRow()>=2){var cat=shCat.getDataRange().getValues();for(var i=1;i<cat.length;i++){var sedeCat=String(cat[i][0]||'').trim();var abrevCat=String(cat[i][1]||'').trim();if(sedeCat&&abrevCat)catMap[sedeCat]=abrevCat;}}
  var sedesEnContratos={}; var snap=getContratosSnapshot_(); var cS=snap.idx.SEDE; var cSt=snap.idx[COL_STATUS_TRAB];
  if(cS!==undefined){snap.rows.forEach(function(r){var st=cSt!==undefined?String(r[cSt]||'ACTIVO').toUpperCase().trim():'ACTIVO';if(st==='BAJA')return;var s=String(r[cS]||'').trim();if(s)sedesEnContratos[s]=true;});}
  var all={}; Object.keys(catMap).forEach(function(s){all[s]=true;}); Object.keys(sedesEnContratos).forEach(function(s){all[s]=true;});
  var out=Object.keys(all).map(function(sede){var abrev=catMap[sede];if(!abrev){abrev=sede.split(/s+/).filter(function(w){return w.length>2&&['DE','LA','EL','LOS','LAS','DEL'].indexOf(w.toUpperCase())<0;}).map(function(w){return w[0];}).join('').toUpperCase().substring(0,4);if(!abrev)abrev=sede.substring(0,3).toUpperCase();}return{sede:sede,abrev:abrev};});
  out.sort(function(a,b){return a.sede.localeCompare(b.sede,'es');}); cachePutJSON_('MHS:SEDES:v2',out,MHS_CACHE_SEDES_TTL); return out;
}
function abrevBySede_(sede){sede=String(sede||'').trim();if(!sede)return'';var sedes=getSedes_();for(var i=0;i<sedes.length;i++)if(sedes[i].sede===sede)return sedes[i].abrev;return'';}
function apiLogin(payload){var r=apiLogin_(payload||{});return r.ok?{ok:true,token:r.token,role:r.role,username:r.username,fullName:r.fullName||''}:{ok:false,message:r.message||r.error||'Login fallido'};}
function apiLogin_(body){var u=String((body&&body.username)||'').toLowerCase().trim();var p=String((body&&body.password)||'');if(!u||!p)return{ok:false,message:'Falta usuario/contrasena'};var users=readUsers_();var found=null;for(var i=0;i<users.length;i++)if(users[i].username===u&&users[i].active){found=users[i];break;}if(!found)return{ok:false,message:'Usuario no existe o inactivo'};if(found.password!==p)return{ok:false,message:'Contrasena incorrecta'};return{ok:true,token:issueToken_(u,found.role),role:found.role,username:u,fullName:found.fullName||''};}
function apiGetEmpleadosPorSede(sede){
  var sedeN=String(sede||'').trim(); if(!sedeN)throw new Error('Falta sede');
  var snap=getContratosSnapshot_(); var idx=snap.idx; var cID=idx.ID,cSEDE=idx.SEDE,cNOM=idx.NOMBRE_TRABAJADOR,cSTAT=idx[COL_STATUS_TRAB];
  var cTURN=(idx.JORNADA!==undefined)?idx.JORNADA:((idx.TURNO!==undefined)?idx.TURNO:-1); var cDIA=(idx.DIA_DESCANSO!==undefined)?idx.DIA_DESCANSO:-1;
  if(cID===undefined||cSEDE===undefined||cNOM===undefined)throw new Error('CONTRATOS_2026 requiere columnas: ID, SEDE, NOMBRE_TRABAJADOR');
  var estadoPase=snap.estadoPase||{}; var out=[];
  for(var i=0;i<snap.rows.length;i++){var r=snap.rows[i]; if(String(r[cSEDE]||'').trim()!==sedeN)continue; var id=formatCellValue_(r[cID]); if(!id)continue; var st=cSTAT!==undefined?String(r[cSTAT]||'ACTIVO').toUpperCase().trim():'ACTIVO'; if(st==='BAJA')continue; if(estadoPase[id]&&estadoPase[id].estado==='BAJA')continue; out.push({id:id,nombre:String(r[cNOM]||'').trim(),turno:cTURN>=0?String(r[cTURN]||'').trim():'',diaDescanso:cDIA>=0?String(r[cDIA]||'').trim():''});}
  out.sort(function(a,b){return a.nombre.localeCompare(b.nombre,'es');}); return out;
}
function apiGetEmpleadosPorSedeJornada(payload){var sede='',jornada='';if(typeof payload==='string')sede=String(payload||'').trim();else{sede=String((payload&&payload.sede)||'').trim();jornada=String((payload&&payload.jornada)||'').trim().toUpperCase();var token=String((payload&&payload.token)||'');if(token&&(!sede||!jornada)){try{var asign=apiGetSedeJornadaUsuario({token:token}).asignaciones||[];if(asign.length){if(!sede)sede=asign[0].sede;if(!jornada)jornada=asign[0].jornada;}}catch(e){}}}var todos=apiGetEmpleadosPorSede(sede);if(!jornada)return todos;return todos.filter(function(e){return String(e.turno||'').toUpperCase().indexOf(jornada)>=0;});}
function apiGetPendientesAsistencia(payload){
  var token=String((payload&&payload.token)||''); requireSession_(token); var fechaISO=String((payload&&payload.fecha)||todayISO_()).trim();
  var asignaciones=[]; try{asignaciones=apiGetSedeJornadaUsuario({token:token}).asignaciones||[];}catch(e){}
  if(!asignaciones.length&&payload&&payload.sede)asignaciones=[{sede:String(payload.sede).trim(),jornada:String(payload.jornada||'').trim().toUpperCase()}];
  if(!asignaciones.length)return{ok:true,pendientes:0,total:0,sedes:[],resumen:[]};
  var wanted={}; asignaciones.forEach(function(a){wanted[a.sede]=true;}); var marcados={}; var pl=mhsGetSS_().getSheetByName(SH_PL);
  if(pl&&pl.getLastRow()>=2){var v=pl.getDataRange().getValues();for(var i=1;i<v.length;i++){var f=formatCellValue_(v[i][0]);if(f!==fechaISO)continue;var rowSede=String(v[i][2]||'').trim();if(!wanted[rowSede])continue;var id=formatCellValue_(v[i][4]);var cod=String(v[i][6]||'').trim();if(id&&cod)marcados[rowSede+'|'+id]=cod;}}
  var totalPend=0,totalEmps=0,resumen=[]; asignaciones.forEach(function(a){var todos=apiGetEmpleadosPorSede(a.sede);var j=String(a.jornada||'').toUpperCase();if(j)todos=todos.filter(function(e){return String(e.turno||'').toUpperCase().indexOf(j)>=0;});var pend=0;todos.forEach(function(e){if(!marcados[a.sede+'|'+e.id])pend++;});totalPend+=pend;totalEmps+=todos.length;resumen.push({sede:a.sede,jornada:a.jornada,pendientes:pend,total:todos.length});});
  return{ok:true,pendientes:totalPend,total:totalEmps,sede:asignaciones[0].sede,jornada:asignaciones[0].jornada,fechaISO:fechaISO,sedes:asignaciones.map(function(a){return a.sede;}),resumen:resumen};
}
function apiMobileBootstrap(payload){var login=apiLogin(payload||{});if(!login.ok)return login;var token=login.token;var fecha=todayISO_();var sedes=getSedes_();var asign=apiGetSedeJornadaUsuario({token:token});var pendientes=apiGetPendientesAsistencia({token:token,fecha:fecha});return{ok:true,token:token,role:login.role,username:login.username,fullName:login.fullName||'',today:fecha,sedes:sedes,asignaciones:asign.asignaciones||[],sede:asign.sede||'',jornada:asign.jornada||'',pendientes:pendientes};}
function apiMobileSessionBootstrap(payload){var token=String((payload&&payload.token)||'');var ses=requireSession_(token);var fecha=String((payload&&payload.fecha)||todayISO_()).trim();var users=readUsers_();var u=users.find(function(x){return x.username===ses.username;})||{};var sedes=getSedes_();var asign=apiGetSedeJornadaUsuario({token:token});var pendientes=apiGetPendientesAsistencia({token:token,fecha:fecha});var perm=canEdit_(ses.role,fecha);var lock=isLockedByNomina_(ses.role,fecha);return{ok:true,token:token,role:ses.role,username:ses.username,fullName:u.fullName||'',today:fecha,sedes:sedes,asignaciones:asign.asignaciones||[],sede:asign.sede||'',jornada:asign.jornada||'',pendientes:pendientes,canEdit:!!(perm.ok&&!lock.locked),isAdmin:isAdminRole_(ses.role),reason:(!perm.ok?perm.reason:(lock.locked?('Nomina cerrada: '+lock.reason):''))};}
function apiMobileCargarPase(payload){var token=String((payload&&payload.token)||'');var fecha=String((payload&&payload.fecha)||'').trim();var sede=String((payload&&payload.sede)||'').trim();var jornada=String((payload&&payload.jornada)||'').trim().toUpperCase();var ses=requireSession_(token);if(!isoValid_(fecha))throw new Error('Fecha invalida');if(!sede)throw new Error('Falta sede');var perm=canEdit_(ses.role,fecha);var lock=isLockedByNomina_(ses.role,fecha);var can=perm.ok&&!lock.locked;var reason=!perm.ok?perm.reason:(lock.locked?'Nomina cerrada: '+lock.reason:'');if(!can&&!isAdminRole_(ses.role))return{ok:false,canEdit:false,isAdmin:false,reason:reason||'Sin permiso para esta fecha'};var empleados=apiGetEmpleadosPorSedeJornada({token:token,sede:sede,jornada:jornada});var marks=apiGetPaseListaDia(fecha,sede);return{ok:true,canEdit:!!can,isAdmin:isAdminRole_(ses.role),reason:reason,graceExpiry:perm.graceExpiry||null,liberada:perm.liberada||false,empleados:empleados,marks:marks,sedes:getSedes_()};}


// ══════════════════════════════════════════════════════
//  OPTIMIZACIÓN V2 ULTRA — CACHE CONTROL + BOOTSTRAP SEGURO
// ══════════════════════════════════════════════════════
var MHS_V2_CACHE_VERSION = 'v2u_20260430';
var MHS_CACHE_SEDES_TTL_V2 = 180;
var MHS_CACHE_USERS_TTL_V2 = 180;
function cacheKeyV2_(name){ return 'MHS:' + name + ':' + MHS_V2_CACHE_VERSION; }
function cacheGetJSON_(key){ try{ var raw=CacheService.getScriptCache().get(key); return raw?JSON.parse(raw):null; }catch(e){ return null; } }
function cachePutJSON_(key,obj,ttl){ try{ CacheService.getScriptCache().put(key,JSON.stringify(obj),ttl||60); }catch(e){} }
function mhsGetSS_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function apiClearSystemCache(payload){
  var token=String((payload&&payload.token)||''); var ses=requireSession_(token);
  if(!isAdminRole_(ses.role)) throw new Error('Solo ADMIN / SUPERADMIN / CEO puede limpiar caché.');
  try{CacheService.getScriptCache().removeAll(['MHS:USERS:'+MHS_V2_CACHE_VERSION,'MHS:SEDES:'+MHS_V2_CACHE_VERSION,'MHS:USERS:v1','MHS:SEDES:v2','SEDES','USERS']);}catch(e){}
  __MHS_MEM__={}; return {ok:true,message:'Caché limpiado correctamente',ts:isoNow_(),version:MHS_V2_CACHE_VERSION};
}
function apiWarmupSystem(payload){
  var token=String((payload&&payload.token)||''); var ses=requireSession_(token);
  if(!isAdminRole_(ses.role)) throw new Error('Solo ADMIN / SUPERADMIN / CEO puede precargar el sistema.');
  var t0=Date.now(); var users=readUsers_(); var sedes=getSedes_();
  return {ok:true,message:'Sistema precargado',ms:Date.now()-t0,users:users.length,sedes:sedes.length,version:MHS_V2_CACHE_VERSION};
}
function getContratosSnapshot_(){
  if(__MHS_MEM__&&__MHS_MEM__.contratosSnap) return __MHS_MEM__.contratosSnap;
  var sh=mhsGetSS_().getSheetByName(SH_CONTRATOS);
  if(!sh||sh.getLastRow()<2){__MHS_MEM__.contratosSnap={rows:[],idx:{},estadoPase:readEstadoPaseMap_()};return __MHS_MEM__.contratosSnap;}
  var raw=sh.getDataRange().getValues(); var hdr=raw[0].map(normHdr_); var idx={}; hdr.forEach(function(h,i){if(h)idx[h]=i;});
  __MHS_MEM__.contratosSnap={rows:raw.slice(1),idx:idx,estadoPase:readEstadoPaseMap_()}; return __MHS_MEM__.contratosSnap;
}
function readUsers_(){
  var key=cacheKeyV2_('USERS'); var cached=cacheGetJSON_(key); if(cached)return cached;
  var sh=mhsGetSS_().getSheetByName(SH_USERS); if(!sh||sh.getLastRow()<2)return [];
  var v=sh.getDataRange().getValues(); var idx={}; v[0].forEach(function(h,i){idx[normHdr_(h)]=i;}); var out=[];
  for(var i=1;i<v.length;i++){var r=v[i];var activeRaw=String(r[idx.ACTIVE]||'').toLowerCase().trim();out.push({username:String(r[idx.USERNAME]||'').toLowerCase().trim(),password:String(r[idx.PASSWORD]||''),role:String(r[idx.ROLE]||'USER').toUpperCase().trim(),active:activeRaw==='true'||activeRaw==='si'||activeRaw==='sí'||activeRaw==='1'||activeRaw==='activo',fullName:String(r[idx.FULL_NAME]||''),email:String((idx.EMAIL!==undefined?r[idx.EMAIL]:'')||''),sede:String((idx.SEDE!==undefined?r[idx.SEDE]:'')||'').trim(),jornada:String((idx.JORNADA!==undefined?r[idx.JORNADA]:'')||'').trim().toUpperCase()});}
  cachePutJSON_(key,out,MHS_CACHE_USERS_TTL_V2); return out;
}
function getSedes_(){
  var key=cacheKeyV2_('SEDES'); var cached=cacheGetJSON_(key); if(cached)return cached;
  var ss=mhsGetSS_(); var catMap={}; var shCat=ss.getSheetByName(SH_CAT_SEDES);
  if(shCat&&shCat.getLastRow()>=2){var cat=shCat.getDataRange().getValues();for(var i=1;i<cat.length;i++){var s1=String(cat[i][0]||'').trim();var a1=String(cat[i][1]||'').trim();if(s1&&a1)catMap[s1]=a1;}}
  var sedesEnContratos={}; var snap=getContratosSnapshot_(); var cS=snap.idx.SEDE; var cSt=snap.idx[COL_STATUS_TRAB];
  if(cS!==undefined){snap.rows.forEach(function(r){var st=cSt!==undefined?String(r[cSt]||'ACTIVO').toUpperCase().trim():'ACTIVO';if(st==='BAJA')return;var s=String(r[cS]||'').trim();if(s)sedesEnContratos[s]=true;});}
  var all={}; Object.keys(catMap).forEach(function(s){all[s]=true;}); Object.keys(sedesEnContratos).forEach(function(s){all[s]=true;});
  var out=Object.keys(all).map(function(sede){var abrev=catMap[sede]; if(!abrev){abrev=sede.split(/\s+/).filter(function(w){return w.length>2&&['DE','LA','EL','LOS','LAS','DEL'].indexOf(w.toUpperCase())<0;}).map(function(w){return w[0];}).join('').toUpperCase().substring(0,4); if(!abrev)abrev=sede.substring(0,3).toUpperCase();} return {sede:sede,abrev:abrev};});
  out.sort(function(a,b){return a.sede.localeCompare(b.sede,'es');}); cachePutJSON_(key,out,MHS_CACHE_SEDES_TTL_V2); return out;
}
function apiLogin(payload){var r=apiLogin_(payload||{});return r.ok?{ok:true,token:r.token,role:r.role,username:r.username,fullName:r.fullName||'',cacheVersion:MHS_V2_CACHE_VERSION}:{ok:false,message:r.message||r.error||'Login fallido'};}
function apiMobileBootstrap(payload){var t0=Date.now();var login=apiLogin(payload||{});if(!login.ok)return login;var token=login.token;var fecha=todayISO_();var sedes=getSedes_();var asign=apiGetSedeJornadaUsuario({token:token});var pendientes=apiGetPendientesAsistencia({token:token,fecha:fecha});return{ok:true,token:token,role:login.role,username:login.username,fullName:login.fullName||'',today:fecha,sedes:sedes,asignaciones:asign.asignaciones||[],sede:asign.sede||'',jornada:asign.jornada||'',pendientes:pendientes,perf:{ms:Date.now()-t0,cacheVersion:MHS_V2_CACHE_VERSION}};}
function apiMobileSessionBootstrap(payload){var t0=Date.now();var token=String((payload&&payload.token)||'');var ses=requireSession_(token);var fecha=String((payload&&payload.fecha)||todayISO_()).trim();var users=readUsers_();var u=users.find(function(x){return x.username===ses.username;})||{};var sedes=getSedes_();var asign=apiGetSedeJornadaUsuario({token:token});var pendientes=apiGetPendientesAsistencia({token:token,fecha:fecha});var perm=canEdit_(ses.role,fecha);var lock=isLockedByNomina_(ses.role,fecha);return{ok:true,token:token,role:ses.role,username:ses.username,fullName:u.fullName||'',today:fecha,sedes:sedes,asignaciones:asign.asignaciones||[],sede:asign.sede||'',jornada:asign.jornada||'',pendientes:pendientes,canEdit:!!(perm.ok&&!lock.locked),isAdmin:isAdminRole_(ses.role),reason:(!perm.ok?perm.reason:(lock.locked?('Nomina cerrada: '+lock.reason):'')),jornadasPorSede:getJornadasPorSede_(ses.username),perf:{ms:Date.now()-t0,cacheVersion:MHS_V2_CACHE_VERSION}};}
function apiMobileCargarPase(payload){var t0=Date.now();var token=String((payload&&payload.token)||'');var fecha=String((payload&&payload.fecha)||'').trim();var sede=String((payload&&payload.sede)||'').trim();var jornada=String((payload&&payload.jornada)||'').trim().toUpperCase();var ses=requireSession_(token);if(!isoValid_(fecha))throw new Error('Fecha invalida');if(!sede)throw new Error('Falta sede');var perm=canEdit_(ses.role,fecha);var lock=isLockedByNomina_(ses.role,fecha);var can=perm.ok&&!lock.locked;var reason=!perm.ok?perm.reason:(lock.locked?'Nomina cerrada: '+lock.reason:'');if(!can&&!isAdminRole_(ses.role))return{ok:false,canEdit:false,isAdmin:false,reason:reason||'Sin permiso para esta fecha'};var empleados=apiGetEmpleadosPorSedeJornada({token:token,sede:sede,jornada:jornada});var ids={};empleados.forEach(function(e){ids[String(e.id)]=true;});var rawMarks=apiGetPaseListaDia(fecha,sede)||{};var marks={};Object.keys(rawMarks).forEach(function(id){if(ids[String(id)])marks[id]=rawMarks[id];});return{ok:true,canEdit:!!can,isAdmin:isAdminRole_(ses.role),reason:reason,graceExpiry:perm.graceExpiry||null,liberada:perm.liberada||false,empleados:empleados,marks:marks,sedes:getSedes_(),perf:{ms:Date.now()-t0,cacheVersion:MHS_V2_CACHE_VERSION}};}
function apiGuardarPaseListaMobile(payload){var res=apiGuardarPaseLista(payload||{});try{var token=String((payload&&payload.token)||'');var fecha=String((payload&&payload.fecha)||todayISO_()).trim();res.pendientes=apiGetPendientesAsistencia({token:token,fecha:fecha});}catch(e){res.pendientesError=String(e&&e.message||e);}return res;}


// ══════════════════════════════════════════════════════════════
//  SOPORTE V3 PRO — Tickets en vivo seguro + prioridad + polling delta
//  Añadido sin romper el flujo anterior.
// ══════════════════════════════════════════════════════════════
var HDR_SOPORTE_V3 = [
  "TS_ISO","TOKEN_CHAT","USERNAME","FULL_NAME","SEDE","MENSAJE","TIPO","LEIDO",
  "PRIORIDAD","ESTADO","FECHA_SOLICITADA","JORNADA","DEVICE_ID","RESUELTO_TS","ASIGNADO_A"
];

function _ensureSoporteSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_SOPORTE);
  if (!sh) {
    sh = ss.insertSheet(SH_SOPORTE);
    sh.getRange(1, 1, 1, HDR_SOPORTE_V3.length).setValues([HDR_SOPORTE_V3])
      .setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, HDR_SOPORTE_V3.length);
    return sh;
  }
  var lastCol = Math.max(1, sh.getLastColumn());
  var current = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(normHdr_);
  var missing = [];
  HDR_SOPORTE_V3.forEach(function(h){ if (current.indexOf(normHdr_(h)) < 0) missing.push(h); });
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing])
      .setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold");
    sh.autoResizeColumns(1, sh.getLastColumn());
  }
  return sh;
}

function soporteIdx_(values) {
  var h = (values && values.length ? values[0] : HDR_SOPORTE_V3).map(normHdr_);
  var idx = {};
  h.forEach(function(k, i){ if (k) idx[k] = i; });
  return idx;
}
function soporteCell_(row, idx, key, fallbackIndex) {
  var i = idx[normHdr_(key)];
  if (i === undefined || i === null) i = fallbackIndex;
  return String(row[i] || "").trim();
}
function soportePrioridad_(tipo) {
  tipo = String(tipo || "CONSULTA").toUpperCase().trim();
  if (tipo === "INCIDENCIA_URGENTE") return "URGENTE";
  if (tipo === "DESBLOQUEO_FECHA") return "ALTA";
  if (tipo === "CONSULTA") return "MEDIA";
  return "NORMAL";
}
function soporteEstadoInicial_(tipo) {
  tipo = String(tipo || "CONSULTA").toUpperCase().trim();
  if (tipo === "INCIDENCIA_URGENTE") return "URGENTE";
  if (tipo === "DESBLOQUEO_FECHA") return "PENDIENTE_FECHA";
  return "PENDIENTE";
}
function soporteTipoIcon_(tipo) {
  tipo = String(tipo || "CONSULTA").toUpperCase().trim();
  return ({ CONSULTA:"💬", DESBLOQUEO_FECHA:"🔓", INCIDENCIA_URGENTE:"🚨", OTRO:"📋", RESPUESTA:"✅", RESUELTO:"✅" })[tipo] || "📋";
}

function apiEnviarTicket(payload) {
  var token       = String((payload && payload.token) || "");
  var ses         = requireSession_(token);
  var mensaje     = String((payload && payload.mensaje) || "").trim().substring(0, 2000);
  var sede        = String((payload && payload.sede) || "").trim();
  var chatId      = String((payload && payload.chatId) || Utilities.getUuid().replace(/-/g, "").substring(0, 12)).trim();
  var tipoTicket  = String((payload && payload.tipo_ticket) || "CONSULTA").trim().toUpperCase();
  var fechaSol    = String((payload && payload.fecha_solicitud) || "").trim();
  var deviceId    = String((payload && payload.device_id) || "").trim();
  if (!mensaje) throw new Error("El mensaje no puede estar vacío.");
  var users = readUsers_();
  var usuario = users.find(function(u){ return u.username === ses.username; });
  var fullName = usuario ? (usuario.fullName || ses.username) : ses.username;
  var sedeEff = sede || (usuario && usuario.sede) || "No especificada";
  var jornada = (usuario && usuario.jornada) ? usuario.jornada : String((payload && payload.jornada) || "").trim();
  var prioridad = soportePrioridad_(tipoTicket);
  var estado = soporteEstadoInicial_(tipoTicket);
  var ts = isoNow_();
  var sh = _ensureSoporteSheet_();
  var values = sh.getDataRange().getValues();
  var idx = soporteIdx_(values);
  var row = new Array(sh.getLastColumn()).fill("");
  function set(k, v){ var i = idx[normHdr_(k)]; if (i !== undefined) row[i] = v; }
  set("TS_ISO", ts); set("TOKEN_CHAT", chatId); set("USERNAME", ses.username); set("FULL_NAME", fullName);
  set("SEDE", sedeEff); set("MENSAJE", mensaje); set("TIPO", tipoTicket); set("LEIDO", "NO");
  set("PRIORIDAD", prioridad); set("ESTADO", estado); set("FECHA_SOLICITADA", fechaSol); set("JORNADA", jornada);
  set("DEVICE_ID", deviceId); set("ASIGNADO_A", "");
  sh.appendRow(row);
  var recipients = getSoporteEmailRecipients_();
  var icono = soporteTipoIcon_(tipoTicket);
  var asunto = icono + " [MHS Ticket " + prioridad + "] " + tipoTicket + " · " + fullName + " · " + sedeEff + " · " + ts.substring(0,16);
  var extra = tipoTicket === "DESBLOQUEO_FECHA" && fechaSol ? "Fecha a desbloquear: " + fechaSol + "\n" : "";
  var cuerpo = "MHS RH HUB · Nuevo Ticket\n━━━━━━━━━━━━━━━━━━━\n" +
    "Tipo: " + tipoTicket + "\nPrioridad: " + prioridad + "\nEstado: " + estado + "\n" +
    "Usuario: " + ses.username + "\nNombre: " + fullName + "\nSede: " + sedeEff +
    (jornada ? "\nJornada: " + jornada : "") + "\nFecha: " + ts + "\nChat ID: " + chatId +
    (extra ? "\n" + extra : "") + "━━━━━━━━━━━━━━━━━━━\n\n" + mensaje;
  try { recipients.forEach(function(rcpt){ MailApp.sendEmail({to:rcpt, subject:asunto, body:cuerpo}); }); } catch(e){ Logger.log("Ticket email V3: " + e); }
  return { ok:true, chatId:chatId, ts:ts, tipo:tipoTicket, prioridad:prioridad, estado:estado, message:"Ticket enviado. Soporte fue notificado." };
}

function apiGetInboxSoporte(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN.");
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SOPORTE);
  if (!sh || sh.getLastRow() < 2) return { ok:true, chats:[], totalUnread:0, latestTs:"" };
  var v = sh.getDataRange().getValues();
  var idx = soporteIdx_(v);
  var map = {}, order = [], totalUnread = 0, latestTs = "";
  for (var i=1; i<v.length; i++) {
    var r = v[i], chatId = soporteCell_(r, idx, "TOKEN_CHAT", 1);
    if (!chatId) continue;
    var ts = soporteCell_(r, idx, "TS_ISO", 0);
    var tipo = soporteCell_(r, idx, "TIPO", 6).toUpperCase();
    var leido = soporteCell_(r, idx, "LEIDO", 7).toUpperCase();
    if (ts > latestTs) latestTs = ts;
    if (!map[chatId]) {
      map[chatId] = {
        chatId: chatId, username: soporteCell_(r, idx, "USERNAME", 2),
        fullName: soporteCell_(r, idx, "FULL_NAME", 3),
        sede: soporteCell_(r, idx, "SEDE", 4),
        jornada: soporteCell_(r, idx, "JORNADA", 11),
        prioridad: soporteCell_(r, idx, "PRIORIDAD", 8) || soportePrioridad_(tipo),
        estado: soporteCell_(r, idx, "ESTADO", 9) || soporteEstadoInicial_(tipo),
        fechaSolicitada: soporteCell_(r, idx, "FECHA_SOLICITADA", 10),
        firstTs: ts, lastTs: ts, lastMsg: soporteCell_(r, idx, "MENSAJE", 5),
        lastTipo: tipo, unread: 0, resuelto: false
      };
      order.push(chatId);
    }
    var c = map[chatId];
    if (ts >= c.lastTs) { c.lastTs = ts; c.lastMsg = soporteCell_(r, idx, "MENSAJE", 5); c.lastTipo = tipo; }
    var pr = soporteCell_(r, idx, "PRIORIDAD", 8); if (pr && (c.prioridad !== "URGENTE")) c.prioridad = pr;
    var est = soporteCell_(r, idx, "ESTADO", 9); if (est) c.estado = est;
    if (tipo === "RESUELTO") { c.resuelto = true; c.estado = "RESUELTO"; }
    if (tipo !== "RESPUESTA" && tipo !== "RESUELTO" && leido !== "SI") { c.unread++; totalUnread++; }
  }
  var rank = { URGENTE:0, ALTA:1, MEDIA:2, NORMAL:3 };
  var chats = order.map(function(id){ return map[id]; }).sort(function(a,b){
    var ar = rank[a.prioridad] !== undefined ? rank[a.prioridad] : 9;
    var br = rank[b.prioridad] !== undefined ? rank[b.prioridad] : 9;
    if (a.unread && !b.unread) return -1;
    if (!a.unread && b.unread) return 1;
    if (ar !== br) return ar - br;
    return a.lastTs < b.lastTs ? 1 : -1;
  });
  return { ok:true, chats:chats, totalUnread:totalUnread, latestTs:latestTs };
}

function apiGetPendienteSoporte(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN.");
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SOPORTE);
  if (!sh || sh.getLastRow() < 2) return { ok:true, count:0, urgent:0, latestTs:"" };
  var v = sh.getDataRange().getValues(), idx = soporteIdx_(v);
  var count = 0, urgent = 0, latestTs = "";
  for (var i=1; i<v.length; i++) {
    var r = v[i], ts = soporteCell_(r, idx, "TS_ISO", 0);
    var tipo = soporteCell_(r, idx, "TIPO", 6).toUpperCase();
    var leido = soporteCell_(r, idx, "LEIDO", 7).toUpperCase();
    var pr = soporteCell_(r, idx, "PRIORIDAD", 8) || soportePrioridad_(tipo);
    if (ts > latestTs) latestTs = ts;
    if (tipo !== "RESPUESTA" && tipo !== "RESUELTO" && leido !== "SI") {
      count++;
      if (pr === "URGENTE" || tipo === "INCIDENCIA_URGENTE") urgent++;
    }
  }
  return { ok:true, count:count, urgent:urgent, latestTs:latestTs };
}

function apiSoportePollV3(payload) {
  var token = String((payload && payload.token) || "");
  var since = String((payload && payload.since) || "").trim();
  var ses = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN.");
  var inbox = apiGetInboxSoporte({token: token});
  var nuevos = [];
  (inbox.chats || []).forEach(function(c){ if (c.unread > 0 && (!since || c.lastTs > since)) nuevos.push(c); });
  return { ok:true, count:inbox.totalUnread || 0, latestTs:inbox.latestTs || "", chats:nuevos.slice(0,8) };
}

function apiResponderSoporte(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN.");
  var chatId = String((payload && payload.chatId) || "").trim();
  var mensaje = String((payload && payload.mensaje) || "").trim().substring(0,1000);
  if (!chatId || !mensaje) throw new Error("chatId y mensaje requeridos.");
  var users = readUsers_();
  var agente = users.find(function(u){ return u.username === ses.username; });
  var agenteNombre = agente ? (agente.fullName || ses.username) : ses.username;
  var ts = isoNow_(), sh = _ensureSoporteSheet_();
  var values = sh.getDataRange().getValues(), idx = soporteIdx_(values);
  var userEmail = "", userFullName = "", userUsername = "", sede = "", jornada = "";
  for (var i=1; i<values.length; i++) {
    var r = values[i];
    if (soporteCell_(r, idx, "TOKEN_CHAT", 1) !== chatId) continue;
    var tipo = soporteCell_(r, idx, "TIPO", 6).toUpperCase();
    var rowUser = soporteCell_(r, idx, "USERNAME", 2);
    if (tipo !== "RESPUESTA" && tipo !== "RESUELTO") {
      sh.getRange(i+1, idx["LEIDO"]+1).setValue("SI");
      if (!userFullName) userFullName = soporteCell_(r, idx, "FULL_NAME", 3);
      if (!userUsername) userUsername = rowUser;
      if (!sede) sede = soporteCell_(r, idx, "SEDE", 4);
      if (!jornada) jornada = soporteCell_(r, idx, "JORNADA", 11);
    }
  }
  if (userUsername) { var fu = users.find(function(u){ return u.username === userUsername; }); if (fu && fu.email) userEmail = fu.email; }
  var row = new Array(sh.getLastColumn()).fill("");
  function set(k, v){ var ci = idx[normHdr_(k)]; if (ci !== undefined) row[ci] = v; }
  set("TS_ISO", ts); set("TOKEN_CHAT", chatId); set("USERNAME", ses.username); set("FULL_NAME", agenteNombre);
  set("SEDE", sede || "SOPORTE"); set("MENSAJE", mensaje); set("TIPO", "RESPUESTA"); set("LEIDO", "SI");
  set("PRIORIDAD", "NORMAL"); set("ESTADO", "RESPONDIDO"); set("JORNADA", jornada); set("ASIGNADO_A", agenteNombre);
  sh.appendRow(row);
  try { if (userEmail) MailApp.sendEmail(userEmail, "[MHS Soporte] Respuesta · Chat " + chatId.substring(0,8), "Hola " + (userFullName || "Supervisor") + ",\n\nRespuesta del equipo de soporte:\n\n" + mensaje + "\n\n— Soporte MHS"); } catch(e){ Logger.log("Email respuesta V3: " + e); }
  return { ok:true, ts:ts, chatId:chatId, message:"Respuesta enviada." };
}

function apiCerrarTicket(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN.");
  var chatId = String((payload && payload.chatId) || "").trim();
  if (!chatId) throw new Error("chatId requerido.");
  var ts = isoNow_(), sh = _ensureSoporteSheet_();
  var values = sh.getDataRange().getValues(), idx = soporteIdx_(values);
  var users = readUsers_();
  var agente = users.find(function(u){ return u.username === ses.username; });
  var agenteNombre = agente ? (agente.fullName || ses.username) : ses.username;
  var userUsername = "", userFullName = "", sede = "", jornada = "";
  for (var i=1; i<values.length; i++) {
    var r = values[i];
    if (soporteCell_(r, idx, "TOKEN_CHAT", 1) !== chatId) continue;
    var tipo = soporteCell_(r, idx, "TIPO", 6).toUpperCase();
    if (tipo !== "RESPUESTA" && tipo !== "RESUELTO") {
      sh.getRange(i+1, idx["LEIDO"]+1).setValue("SI");
      if (!userUsername) userUsername = soporteCell_(r, idx, "USERNAME", 2);
      if (!userFullName) userFullName = soporteCell_(r, idx, "FULL_NAME", 3);
      if (!sede) sede = soporteCell_(r, idx, "SEDE", 4);
      if (!jornada) jornada = soporteCell_(r, idx, "JORNADA", 11);
    }
  }
  var row = new Array(sh.getLastColumn()).fill("");
  function set(k, v){ var ci = idx[normHdr_(k)]; if (ci !== undefined) row[ci] = v; }
  set("TS_ISO", ts); set("TOKEN_CHAT", chatId); set("USERNAME", ses.username); set("FULL_NAME", agenteNombre);
  set("SEDE", sede || "SOPORTE"); set("MENSAJE", "Ticket cerrado como RESUELTO por " + agenteNombre); set("TIPO", "RESUELTO"); set("LEIDO", "SI");
  set("PRIORIDAD", "NORMAL"); set("ESTADO", "RESUELTO"); set("JORNADA", jornada); set("RESUELTO_TS", ts); set("ASIGNADO_A", agenteNombre);
  sh.appendRow(row);
  try { var supUser = users.find(function(u){ return u.username === userUsername; }); if (supUser && supUser.email) MailApp.sendEmail({to:supUser.email, subject:"[MHS] ✅ Ticket resuelto · " + chatId.substring(0,8), body:"Hola " + (userFullName || "Supervisor") + ",\n\nTu ticket fue marcado como RESUELTO por " + agenteNombre + ".\n\n— Soporte MHS"}); } catch(e){ Logger.log("Email cierre ticket V3: " + e); }
  return { ok:true, ts:ts, message:"Ticket cerrado." };
}


/* ══════════════════════════════════════════════════════
   V5.2 HOTFIX — SOPORTE CHAT EN TIEMPO CASI REAL
   Corrige que el supervisor vea el preview del ticket, pero no vea
   las respuestas de soporte al abrir el chat.

   Causa: apiGetSoporteHistorial filtraba por USERNAME en cada fila.
   Las respuestas se guardan con USERNAME del agente SOPORTE/ADMIN,
   por eso el supervisor no podía leerlas.

   Nuevo comportamiento:
   - SOPORTE / ADMIN / SUPERADMIN: pueden leer cualquier chat.
   - Supervisor: puede leer TODO el hilo completo si el chatId le pertenece.
   - Devuelve latestTs para polling tipo WhatsApp.
══════════════════════════════════════════════════════ */
function apiGetSoporteHistorial(payload) {
  var token  = String((payload && payload.token)  || "");
  var ses    = requireSession_(token);
  var chatId = String((payload && payload.chatId) || "").trim();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SOPORTE);
  if (!sh || sh.getLastRow() < 2) return { ok:true, items:[], chatId:chatId, latestTs:"" };

  var v = sh.getDataRange().getValues();
  var idx = soporteIdx_(v);
  var isStaff = isSoporteRole_(ses.role) || isAdminRole_(ses.role) || String(ses.role).toUpperCase() === "SUPERADMIN";
  var allowedChats = {};

  if (!isStaff) {
    for (var a = 1; a < v.length; a++) {
      var rowUserA = soporteCell_(v[a], idx, "USERNAME", 2).toLowerCase();
      var rowChatA = soporteCell_(v[a], idx, "TOKEN_CHAT", 1);
      var tipoA    = soporteCell_(v[a], idx, "TIPO", 6).toUpperCase();
      // El chat pertenece al supervisor cuando existe al menos un mensaje inicial suyo.
      if (rowUserA === String(ses.username).toLowerCase() && tipoA !== "RESPUESTA" && tipoA !== "RESUELTO" && rowChatA) {
        allowedChats[rowChatA] = true;
      }
    }
    if (chatId && !allowedChats[chatId]) {
      return { ok:true, items:[], chatId:chatId, latestTs:"", error:"Este chat no pertenece a tu usuario." };
    }
  }

  var out = [], latestTs = "", estado = "", prioridad = "", resuelto = false;
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    var rowChat = soporteCell_(r, idx, "TOKEN_CHAT", 1);
    if (!rowChat) continue;
    if (chatId && rowChat !== chatId) continue;
    if (!isStaff && !allowedChats[rowChat]) continue;

    var ts   = soporteCell_(r, idx, "TS_ISO", 0);
    var tipo = soporteCell_(r, idx, "TIPO", 6).toUpperCase();
    if (ts > latestTs) latestTs = ts;
    var est = soporteCell_(r, idx, "ESTADO", 9);
    var pr  = soporteCell_(r, idx, "PRIORIDAD", 8);
    if (est) estado = est;
    if (pr) prioridad = pr;
    if (tipo === "RESUELTO") { resuelto = true; estado = "RESUELTO"; }

    out.push({
      ts: ts,
      chatId: rowChat,
      username: soporteCell_(r, idx, "USERNAME", 2).toLowerCase(),
      fullName: soporteCell_(r, idx, "FULL_NAME", 3),
      sede: soporteCell_(r, idx, "SEDE", 4),
      jornada: soporteCell_(r, idx, "JORNADA", 11),
      mensaje: soporteCell_(r, idx, "MENSAJE", 5),
      tipo: tipo,
      prioridad: pr,
      estado: est,
      leido: soporteCell_(r, idx, "LEIDO", 7).toUpperCase() === "SI"
    });
  }
  out.sort(function(a,b){ return a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0; });
  return { ok:true, items:out, chatId:chatId, latestTs:latestTs, estado:estado, prioridad:prioridad, resuelto:resuelto };
}

function apiGetSoporteHistorialLive(payload) {
  return apiGetSoporteHistorial(payload);
}

/* ══════════════════════════════════════════════════════
   V6 — MESA DE AYUDA HISTÓRICA / TICKETS CON FOLIO
   - Crea SOPORTE_TICKETS y SOPORTE_MENSAJES.
   - Cada ticket tiene folio TCK-0001, TCK-0002...
   - Los tickets cerrados quedan históricos y consultables.
   - Si el supervisor escribe sobre un chat cerrado, se abre ticket nuevo.
   - Mantiene compatibilidad con las APIs usadas por mobile/index.
══════════════════════════════════════════════════════ */
var SH_SOPORTE_TICKETS_V6 = "SOPORTE_TICKETS";
var SH_SOPORTE_MSGS_V6    = "SOPORTE_MENSAJES";
var HDR_SOPORTE_TICKETS_V6 = [
  "FOLIO","TOKEN_CHAT","ESTADO","TIPO","PRIORIDAD","USERNAME","FULL_NAME","SEDE","JORNADA",
  "FECHA_SOLICITADA","APERTURA_TS","ULTIMO_TS","CIERRE_TS","CERRADO_POR","ASIGNADO_A",
  "LAST_MSG","UNREAD_SOPORTE","UNREAD_USER","DEVICE_ID"
];
var HDR_SOPORTE_MSGS_V6 = [
  "TS_ISO","FOLIO","TOKEN_CHAT","USERNAME","FULL_NAME","ROL","SEDE","MENSAJE","TIPO","LEIDO_SOPORTE","LEIDO_USER"
];

function _ensureSheetHeadersV6_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers])
      .setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
    return sh;
  }
  var lastCol = Math.max(1, sh.getLastColumn());
  var cur = sh.getRange(1,1,1,lastCol).getValues()[0].map(normHdr_);
  var missing = [];
  headers.forEach(function(h){ if (cur.indexOf(normHdr_(h)) < 0) missing.push(h); });
  if (missing.length) {
    sh.getRange(1,lastCol+1,1,missing.length).setValues([missing])
      .setBackground("#1a3a6c").setFontColor("#ffffff").setFontWeight("bold");
    sh.autoResizeColumns(1, sh.getLastColumn());
  }
  return sh;
}
function _ticketsSheetV6_(){ return _ensureSheetHeadersV6_(SH_SOPORTE_TICKETS_V6, HDR_SOPORTE_TICKETS_V6); }
function _msgsSheetV6_(){ return _ensureSheetHeadersV6_(SH_SOPORTE_MSGS_V6, HDR_SOPORTE_MSGS_V6); }
function _idxV6_(values){ var h=(values&&values.length?values[0]:[]).map(normHdr_); var idx={}; h.forEach(function(k,i){ if(k) idx[k]=i; }); return idx; }
function _cellV6_(row, idx, key){ var i=idx[normHdr_(key)]; return i===undefined?"":String(row[i]||"").trim(); }
function _setV6_(row, idx, key, val){ var i=idx[normHdr_(key)]; if(i!==undefined) row[i]=val; }
function _rowObjV6_(row, idx){ var o={}; Object.keys(idx).forEach(function(k){ o[k]=String(row[idx[k]]||"").trim(); }); return o; }

function _nextTicketFolioV6_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty("MHS_TICKET_FOLIO_SEQ") || 0);
    if (!n) {
      var sh = _ticketsSheetV6_();
      if (sh.getLastRow() >= 2) {
        var vals = sh.getDataRange().getValues();
        var idx = _idxV6_(vals);
        for (var i=1;i<vals.length;i++) {
          var f = _cellV6_(vals[i], idx, "FOLIO");
          var m = /TCK-(\d+)/i.exec(f);
          if (m) n = Math.max(n, Number(m[1]));
        }
      }
    }
    n++;
    props.setProperty("MHS_TICKET_FOLIO_SEQ", String(n));
    return "TCK-" + String(n).padStart(4,"0");
  } finally { lock.releaseLock(); }
}

function _findTicketByChatV6_(chatId) {
  var sh = _ticketsSheetV6_();
  if (!chatId || sh.getLastRow()<2) return null;
  var vals = sh.getDataRange().getValues();
  var idx = _idxV6_(vals);
  for (var i=1;i<vals.length;i++) {
    if (_cellV6_(vals[i], idx, "TOKEN_CHAT") === chatId) {
      var obj = _rowObjV6_(vals[i], idx); obj._row = i+1; obj._idx = idx; obj._sheet = sh;
      return obj;
    }
  }
  return null;
}
function _ticketObjToClientV6_(t) {
  return {
    folio: t.FOLIO || t.folio || "",
    chatId: t.TOKEN_CHAT || t.chatId || "",
    username: t.USERNAME || t.username || "",
    fullName: t.FULL_NAME || t.fullName || "",
    sede: t.SEDE || t.sede || "",
    jornada: t.JORNADA || t.jornada || "",
    tipo: t.TIPO || t.tipo || "CONSULTA",
    prioridad: t.PRIORIDAD || t.prioridad || "MEDIA",
    estado: t.ESTADO || t.estado || "ABIERTO",
    firstTs: t.APERTURA_TS || t.firstTs || "",
    ts: t.ULTIMO_TS || t.ts || "",
    lastTs: t.ULTIMO_TS || t.lastTs || "",
    cierreTs: t.CIERRE_TS || t.cierreTs || "",
    cerradoPor: t.CERRADO_POR || t.cerradoPor || "",
    asignadoA: t.ASIGNADO_A || t.asignadoA || "",
    lastMsg: t.LAST_MSG || t.lastMsg || "",
    unread: Number(t.UNREAD_SOPORTE || t.unread || 0),
    unreadUser: Number(t.UNREAD_USER || t.unreadUser || 0),
    totalMsgs: Number(t.totalMsgs || 0),
    respondido: String(t.ESTADO||"").toUpperCase()==="RESPONDIDO",
    resuelto: String(t.ESTADO||"").toUpperCase()==="CERRADO" || String(t.ESTADO||"").toUpperCase()==="RESUELTO"
  };
}
function _updateTicketRowV6_(ticket, patch) {
  var sh = ticket._sheet || _ticketsSheetV6_();
  var idx = ticket._idx;
  if (!idx) idx = _idxV6_(sh.getDataRange().getValues());
  Object.keys(patch).forEach(function(k){ var ci = idx[normHdr_(k)]; if(ci!==undefined) sh.getRange(ticket._row, ci+1).setValue(patch[k]); });
}
function _appendMsgV6_(data) {
  var sh = _msgsSheetV6_();
  var vals = sh.getDataRange().getValues();
  var idx = _idxV6_(vals);
  var row = new Array(sh.getLastColumn()).fill("");
  Object.keys(data).forEach(function(k){ _setV6_(row, idx, k, data[k]); });
  sh.appendRow(row);
}
function _createTicketV6_(data) {
  var sh = _ticketsSheetV6_();
  var vals = sh.getDataRange().getValues();
  var idx = _idxV6_(vals);
  var row = new Array(sh.getLastColumn()).fill("");
  Object.keys(data).forEach(function(k){ _setV6_(row, idx, k, data[k]); });
  sh.appendRow(row);
  return _findTicketByChatV6_(data.TOKEN_CHAT);
}
function _estadoCerradoV6_(estado) {
  estado = String(estado||"").toUpperCase();
  return estado === "CERRADO" || estado === "RESUELTO";
}

function apiEnviarTicket(payload) {
  var token       = String((payload && payload.token) || "");
  var ses         = requireSession_(token);
  var mensaje     = String((payload && payload.mensaje) || "").trim().substring(0, 2000);
  var sede        = String((payload && payload.sede) || "").trim();
  var chatIdIn    = String((payload && payload.chatId) || "").trim();
  var forceNew    = !!(payload && payload.forceNew);
  var tipoTicket  = String((payload && payload.tipo_ticket) || "CONSULTA").trim().toUpperCase();
  var fechaSol    = String((payload && payload.fecha_solicitud) || "").trim();
  var deviceId    = String((payload && payload.device_id) || "").trim();
  if (!mensaje) throw new Error("El mensaje no puede estar vacío.");
  var users = readUsers_();
  var usuario = users.find(function(u){ return u.username === ses.username; }) || {};
  var fullName = usuario.fullName || ses.username;
  var sedeEff = sede || usuario.sede || "No especificada";
  var jornada = usuario.jornada || String((payload && payload.jornada) || "").trim();
  var ts = isoNow_();
  var prioridad = soportePrioridad_(tipoTicket);
  var estadoInicial = soporteEstadoInicial_(tipoTicket);
  if (estadoInicial === "PENDIENTE") estadoInicial = "ABIERTO";

  var ticket = chatIdIn && !forceNew ? _findTicketByChatV6_(chatIdIn) : null;
  if (ticket && _estadoCerradoV6_(ticket.ESTADO)) ticket = null;
  var chatId = ticket ? ticket.TOKEN_CHAT : (Utilities.getUuid().replace(/-/g, "").substring(0, 12));
  var folio = ticket ? ticket.FOLIO : _nextTicketFolioV6_();

  if (!ticket) {
    ticket = _createTicketV6_({
      FOLIO: folio, TOKEN_CHAT: chatId, ESTADO: estadoInicial, TIPO: tipoTicket, PRIORIDAD: prioridad,
      USERNAME: ses.username, FULL_NAME: fullName, SEDE: sedeEff, JORNADA: jornada,
      FECHA_SOLICITADA: fechaSol, APERTURA_TS: ts, ULTIMO_TS: ts, CIERRE_TS: "", CERRADO_POR: "",
      ASIGNADO_A: "", LAST_MSG: mensaje, UNREAD_SOPORTE: 1, UNREAD_USER: 0, DEVICE_ID: deviceId
    });
  } else {
    var unreadS = Number(ticket.UNREAD_SOPORTE || 0) + 1;
    _updateTicketRowV6_(ticket, { ESTADO:"ABIERTO", ULTIMO_TS:ts, LAST_MSG:mensaje, UNREAD_SOPORTE:unreadS });
  }

  _appendMsgV6_({
    TS_ISO: ts, FOLIO: folio, TOKEN_CHAT: chatId, USERNAME: ses.username, FULL_NAME: fullName,
    ROL: ses.role, SEDE: sedeEff, MENSAJE: mensaje, TIPO: tipoTicket, LEIDO_SOPORTE:"NO", LEIDO_USER:"SI"
  });

  try {
    var recipients = getSoporteEmailRecipients_();
    var icono = soporteTipoIcon_(tipoTicket);
    var asunto = icono + " [MHS " + folio + "] " + prioridad + " · " + tipoTicket + " · " + fullName;
    var extra = tipoTicket === "DESBLOQUEO_FECHA" && fechaSol ? "Fecha a desbloquear: " + fechaSol + "\n" : "";
    var cuerpo = "MHS RH HUB · Nuevo ticket histórico\n━━━━━━━━━━━━━━━━━━━\n"+
      "Folio: "+folio+"\nTipo: "+tipoTicket+"\nPrioridad: "+prioridad+"\nEstado: "+estadoInicial+"\n"+
      "Usuario: "+ses.username+"\nNombre: "+fullName+"\nSede: "+sedeEff+(jornada?"\nJornada: "+jornada:"")+"\nFecha: "+ts+"\nChat ID: "+chatId+"\n"+
      extra+"━━━━━━━━━━━━━━━━━━━\n\n"+mensaje;
    recipients.forEach(function(rcpt){ MailApp.sendEmail({to:rcpt, subject:asunto, body:cuerpo}); });
  } catch(e) { Logger.log("Ticket email V6: "+e); }

  return { ok:true, folio:folio, chatId:chatId, ts:ts, tipo:tipoTicket, prioridad:prioridad, estado:(ticket&&ticket.ESTADO)||estadoInicial, message:"Ticket "+folio+" enviado. Soporte fue notificado." };
}

function apiGetInboxSoporte(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN/ADMIN.");
  var sh = _ticketsSheetV6_();
  if (sh.getLastRow()<2) return { ok:true, chats:[], totalUnread:0, latestTs:"" };
  var vals = sh.getDataRange().getValues();
  var idx = _idxV6_(vals);
  var out = [], totalUnread = 0, latestTs = "";
  for (var i=1;i<vals.length;i++) {
    var o = _rowObjV6_(vals[i], idx);
    var c = _ticketObjToClientV6_(o);
    if (c.lastTs > latestTs) latestTs = c.lastTs;
    totalUnread += c.unread || 0;
    out.push(c);
  }
  var rank = { URGENTE:0, ALTA:1, MEDIA:2, NORMAL:3 };
  out.sort(function(a,b){
    var ac = a.resuelto ? 1 : 0, bc = b.resuelto ? 1 : 0;
    if (ac !== bc) return ac-bc;
    if ((a.unread>0) !== (b.unread>0)) return a.unread>0 ? -1 : 1;
    var ar = rank[a.prioridad] !== undefined ? rank[a.prioridad] : 9;
    var br = rank[b.prioridad] !== undefined ? rank[b.prioridad] : 9;
    if (ar !== br && !a.resuelto && !b.resuelto) return ar-br;
    return a.lastTs < b.lastTs ? 1 : -1;
  });
  return { ok:true, chats:out, totalUnread:totalUnread, latestTs:latestTs };
}

function apiGetPendienteSoporte(payload) {
  var inbox = apiGetInboxSoporte(payload || {});
  var urgent = 0, count = 0;
  (inbox.chats||[]).forEach(function(c){ if(!c.resuelto && c.unread>0){ count += c.unread; if(c.prioridad==="URGENTE") urgent += c.unread; } });
  return { ok:true, count:count, urgent:urgent, latestTs:inbox.latestTs||"" };
}
function apiSoportePollV3(payload) {
  var since = String((payload && payload.since) || "").trim();
  var inbox = apiGetInboxSoporte(payload || {});
  var nuevos = (inbox.chats||[]).filter(function(c){ return !c.resuelto && c.unread>0 && (!since || c.lastTs > since); });
  return { ok:true, count:inbox.totalUnread||0, latestTs:inbox.latestTs||"", chats:nuevos.slice(0,8) };
}

function apiMarcarLeidoSoporte(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN/ADMIN.");
  var chatId = String((payload && payload.chatId) || "").trim();
  var ticket = _findTicketByChatV6_(chatId);
  if (ticket) _updateTicketRowV6_(ticket, { UNREAD_SOPORTE:0 });
  return { ok:true };
}

function apiGetMisTickets(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  var sh = _ticketsSheetV6_();
  if (sh.getLastRow()<2) return { ok:true, tickets:[] };
  var vals = sh.getDataRange().getValues();
  var idx = _idxV6_(vals);
  var tickets = [];
  for (var i=1;i<vals.length;i++) {
    var o = _rowObjV6_(vals[i], idx);
    if (String(o.USERNAME||"").toLowerCase() !== String(ses.username).toLowerCase()) continue;
    tickets.push(_ticketObjToClientV6_(o));
  }
  tickets.sort(function(a,b){ return a.firstTs < b.firstTs ? 1 : -1; });
  return { ok:true, tickets:tickets };
}

function apiGetSoporteHistorial(payload) {
  var token  = String((payload && payload.token)  || "");
  var ses    = requireSession_(token);
  var chatId = String((payload && payload.chatId) || "").trim();
  var ticket = _findTicketByChatV6_(chatId);
  if (!ticket) return { ok:true, items:[], chatId:chatId, latestTs:"", error:"Ticket no encontrado." };
  var isStaff = isSoporteRole_(ses.role) || isAdminRole_(ses.role) || String(ses.role).toUpperCase()==="SUPERADMIN";
  if (!isStaff && String(ticket.USERNAME||"").toLowerCase() !== String(ses.username).toLowerCase()) {
    return { ok:true, items:[], chatId:chatId, latestTs:"", error:"Este ticket no pertenece a tu usuario." };
  }
  var sh = _msgsSheetV6_();
  var out = [], latestTs = "";
  if (sh.getLastRow() >= 2) {
    var vals = sh.getDataRange().getValues();
    var idx = _idxV6_(vals);
    for (var i=1;i<vals.length;i++) {
      var r = vals[i];
      if (_cellV6_(r, idx, "TOKEN_CHAT") !== chatId) continue;
      var ts = _cellV6_(r, idx, "TS_ISO");
      if (ts > latestTs) latestTs = ts;
      out.push({
        ts: ts, folio: _cellV6_(r, idx, "FOLIO"), chatId: chatId,
        username: _cellV6_(r, idx, "USERNAME").toLowerCase(), fullName: _cellV6_(r, idx, "FULL_NAME"), rol: _cellV6_(r, idx, "ROL"),
        sede: _cellV6_(r, idx, "SEDE"), mensaje: _cellV6_(r, idx, "MENSAJE"), tipo: _cellV6_(r, idx, "TIPO"),
        prioridad: ticket.PRIORIDAD, estado: ticket.ESTADO, leido: true
      });
    }
  }
  out.sort(function(a,b){ return a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0; });
  if (!isStaff && Number(ticket.UNREAD_USER||0)>0) _updateTicketRowV6_(ticket, { UNREAD_USER:0 });
  return { ok:true, items:out, chatId:chatId, latestTs:latestTs, folio:ticket.FOLIO, estado:ticket.ESTADO, prioridad:ticket.PRIORIDAD, resuelto:_estadoCerradoV6_(ticket.ESTADO), ticket:_ticketObjToClientV6_(ticket) };
}
function apiGetSoporteHistorialLive(payload){ return apiGetSoporteHistorial(payload); }

function apiResponderSoporte(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN/ADMIN.");
  var chatId = String((payload && payload.chatId) || "").trim();
  var mensaje = String((payload && payload.mensaje) || "").trim().substring(0,1000);
  if (!chatId || !mensaje) throw new Error("chatId y mensaje requeridos.");
  var ticket = _findTicketByChatV6_(chatId);
  if (!ticket) throw new Error("Ticket no encontrado.");
  if (_estadoCerradoV6_(ticket.ESTADO)) throw new Error("Este ticket ya está cerrado. Puedes verlo como histórico, pero no responderlo.");
  var users = readUsers_();
  var agente = users.find(function(u){ return u.username === ses.username; }) || {};
  var agenteNombre = agente.fullName || ses.username;
  var ts = isoNow_();
  _appendMsgV6_({ TS_ISO:ts, FOLIO:ticket.FOLIO, TOKEN_CHAT:chatId, USERNAME:ses.username, FULL_NAME:agenteNombre, ROL:ses.role, SEDE:ticket.SEDE||"SOPORTE", MENSAJE:mensaje, TIPO:"RESPUESTA", LEIDO_SOPORTE:"SI", LEIDO_USER:"NO" });
  _updateTicketRowV6_(ticket, { ESTADO:"RESPONDIDO", ULTIMO_TS:ts, LAST_MSG:mensaje, UNREAD_SOPORTE:0, UNREAD_USER:Number(ticket.UNREAD_USER||0)+1, ASIGNADO_A:agenteNombre });
  try {
    var fu = users.find(function(u){ return u.username === ticket.USERNAME; });
    if (fu && fu.email) MailApp.sendEmail(fu.email, "[MHS " + ticket.FOLIO + "] Respuesta de soporte", "Hola " + (ticket.FULL_NAME||"Supervisor") + ",\n\nRespuesta del equipo de soporte:\n\n" + mensaje + "\n\n— Soporte MHS");
  } catch(e){ Logger.log("Email respuesta V6: "+e); }
  return { ok:true, ts:ts, chatId:chatId, folio:ticket.FOLIO, message:"Respuesta enviada." };
}

function apiCerrarTicket(payload) {
  var token = String((payload && payload.token) || "");
  var ses = requireSession_(token);
  if (!isSoporteRole_(ses.role)) throw new Error("Solo SOPORTE/SUPERADMIN/ADMIN.");
  var chatId = String((payload && payload.chatId) || "").trim();
  if (!chatId) throw new Error("chatId requerido.");
  var ticket = _findTicketByChatV6_(chatId);
  if (!ticket) throw new Error("Ticket no encontrado.");
  if (_estadoCerradoV6_(ticket.ESTADO)) return { ok:true, ts:ticket.CIERRE_TS||"", message:"El ticket ya estaba cerrado.", folio:ticket.FOLIO };
  var users = readUsers_();
  var agente = users.find(function(u){ return u.username === ses.username; }) || {};
  var agenteNombre = agente.fullName || ses.username;
  var ts = isoNow_();
  var msg = "Ticket " + ticket.FOLIO + " cerrado por " + agenteNombre;
  _appendMsgV6_({ TS_ISO:ts, FOLIO:ticket.FOLIO, TOKEN_CHAT:chatId, USERNAME:ses.username, FULL_NAME:agenteNombre, ROL:ses.role, SEDE:ticket.SEDE||"SOPORTE", MENSAJE:msg, TIPO:"RESUELTO", LEIDO_SOPORTE:"SI", LEIDO_USER:"NO" });
  _updateTicketRowV6_(ticket, { ESTADO:"CERRADO", ULTIMO_TS:ts, CIERRE_TS:ts, CERRADO_POR:agenteNombre, LAST_MSG:msg, UNREAD_SOPORTE:0, UNREAD_USER:Number(ticket.UNREAD_USER||0)+1 });
  try {
    var fu = users.find(function(u){ return u.username === ticket.USERNAME; });
    if (fu && fu.email) MailApp.sendEmail({to:fu.email, subject:"[MHS " + ticket.FOLIO + "] Ticket cerrado", body:"Hola " + (ticket.FULL_NAME||"Supervisor") + ",\n\nTu ticket " + ticket.FOLIO + " fue marcado como CERRADO por " + agenteNombre + ".\n\n— Soporte MHS"});
  } catch(e){ Logger.log("Email cierre V6: "+e); }
  return { ok:true, ts:ts, folio:ticket.FOLIO, message:"Ticket cerrado." };
}

function apiCrearNuevoTicketSoporte(payload) {
  // No crea fila vacía; solo entrega un chatId temporal para el frontend.
  requireSession_(String((payload && payload.token) || ""));
  return { ok:true, chatId:"mob_" + Date.now().toString(36) + "_" + Utilities.getUuid().replace(/-/g,"").substring(0,5) };
}

/* ══════════════════════════════════════════════════════
   V6.1 — FIX CHAT HISTÓRICO + SYNC MÓVIL OPTIMIZADO
   - Si el móvil conserva un chatId temporal/antiguo, abre el último ticket real del usuario.
   - El supervisor recibe respuestas de soporte en el mismo hilo.
   - apiSoporteMobileSync reduce llamadas: tickets + historial en una sola petición.
══════════════════════════════════════════════════════ */
function _findLatestTicketForUserV61_(username) {
  username = String(username || "").toLowerCase().trim();
  var sh = _ticketsSheetV6_();
  if (!username || sh.getLastRow() < 2) return null;
  var vals = sh.getDataRange().getValues();
  var idx = _idxV6_(vals);
  var best = null;
  for (var i = 1; i < vals.length; i++) {
    var o = _rowObjV6_(vals[i], idx);
    if (String(o.USERNAME || "").toLowerCase().trim() !== username) continue;
    o._row = i + 1; o._idx = idx; o._sheet = sh;
    if (!best) { best = o; continue; }
    var oClosed = _estadoCerradoV6_(o.ESTADO), bClosed = _estadoCerradoV6_(best.ESTADO);
    if (bClosed && !oClosed) { best = o; continue; }
    if (Number(o.UNREAD_USER || 0) > 0 && Number(best.UNREAD_USER || 0) <= 0) { best = o; continue; }
    if (String(o.ULTIMO_TS || o.APERTURA_TS || "") > String(best.ULTIMO_TS || best.APERTURA_TS || "")) best = o;
  }
  return best;
}

function apiGetSoporteHistorial(payload) {
  var token  = String((payload && payload.token)  || "");
  var ses    = requireSession_(token);
  var chatId = String((payload && payload.chatId) || "").trim();
  var isStaff = isSoporteRole_(ses.role) || isAdminRole_(ses.role) || String(ses.role).toUpperCase() === "SUPERADMIN";
  var ticket = _findTicketByChatV6_(chatId);
  if (!ticket && !isStaff) {
    ticket = _findLatestTicketForUserV61_(ses.username);
    if (ticket) chatId = ticket.TOKEN_CHAT;
  }
  if (!ticket) return { ok:true, items:[], chatId:chatId, latestTs:"", error:"Ticket no encontrado.", notFound:true };
  if (!isStaff && String(ticket.USERNAME || "").toLowerCase() !== String(ses.username).toLowerCase()) {
    return { ok:true, items:[], chatId:chatId, latestTs:"", error:"Este ticket no pertenece a tu usuario.", forbidden:true };
  }
  var sh = _msgsSheetV6_();
  var out = [], latestTs = "";
  if (sh.getLastRow() >= 2) {
    var vals = sh.getDataRange().getValues();
    var idx = _idxV6_(vals);
    for (var i = 1; i < vals.length; i++) {
      var r = vals[i];
      if (_cellV6_(r, idx, "TOKEN_CHAT") !== chatId) continue;
      var ts = _cellV6_(r, idx, "TS_ISO");
      if (ts > latestTs) latestTs = ts;
      out.push({
        ts: ts,
        folio: _cellV6_(r, idx, "FOLIO"),
        chatId: chatId,
        username: _cellV6_(r, idx, "USERNAME").toLowerCase(),
        fullName: _cellV6_(r, idx, "FULL_NAME"),
        rol: _cellV6_(r, idx, "ROL"),
        sede: _cellV6_(r, idx, "SEDE"),
        mensaje: _cellV6_(r, idx, "MENSAJE"),
        tipo: _cellV6_(r, idx, "TIPO"),
        prioridad: ticket.PRIORIDAD,
        estado: ticket.ESTADO,
        leido: true
      });
    }
  }
  out.sort(function(a,b){ return a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0; });
  if (!isStaff && Number(ticket.UNREAD_USER || 0) > 0) _updateTicketRowV6_(ticket, { UNREAD_USER:0 });
  return {
    ok:true,
    items:out,
    chatId:chatId,
    latestTs:latestTs,
    folio:ticket.FOLIO,
    estado:ticket.ESTADO,
    prioridad:ticket.PRIORIDAD,
    resuelto:_estadoCerradoV6_(ticket.ESTADO),
    ticket:_ticketObjToClientV6_(ticket)
  };
}
function apiGetSoporteHistorialLive(payload){ return apiGetSoporteHistorial(payload); }

function apiSoporteMobileSync(payload) {
  var token = String((payload && payload.token) || "");
  requireSession_(token);
  var requested = String((payload && payload.chatId) || "").trim();
  var ticketsRes = apiGetMisTickets({ token: token });
  var tickets = (ticketsRes && ticketsRes.tickets) ? ticketsRes.tickets : [];
  var chosen = "";
  if (requested && tickets.some(function(t){ return t.chatId === requested; })) chosen = requested;
  if (!chosen) {
    var unread = tickets.find(function(t){ return !t.resuelto && Number(t.unreadUser || 0) > 0; });
    var open = tickets.find(function(t){ return !t.resuelto; });
    var any = tickets[0];
    chosen = (unread || open || any || {}).chatId || requested || "";
  }
  var hist = chosen ? apiGetSoporteHistorial({ token: token, chatId: chosen }) : { ok:true, items:[], chatId:"", latestTs:"" };
  return {
    ok:true,
    tickets:tickets,
    chatId:(hist && hist.chatId) || chosen,
    historial:hist,
    latestTs:(hist && hist.latestTs) || ""
  };
}

/* ══════════════════════════════════════════════════════
   V6.2 — FIX SYNC MÓVIL + SELECCIÓN INTELIGENTE DE TICKET
   - Si el móvil se queda apuntando a un ticket cerrado/viejo, cambia al ticket activo.
   - Si soporte responde desde PC, el móvil prioriza tickets con UNREAD_USER.
   - Mantiene una sola llamada para tickets + mensajes.
══════════════════════════════════════════════════════ */
function _ticketTsV62_(t){ return String((t && (t.lastTs || t.ts || t.firstTs)) || ""); }
function _pickBestUserTicketV62_(tickets, requested){
  tickets = Array.isArray(tickets) ? tickets : [];
  var req = requested ? tickets.find(function(t){ return t.chatId === requested; }) : null;
  var sorted = tickets.slice().sort(function(a,b){
    var au = Number(a.unreadUser || 0), bu = Number(b.unreadUser || 0);
    if ((au > 0) !== (bu > 0)) return au > 0 ? -1 : 1;
    var ao = !a.resuelto, bo = !b.resuelto;
    if (ao !== bo) return ao ? -1 : 1;
    return _ticketTsV62_(a) < _ticketTsV62_(b) ? 1 : -1;
  });
  var unread = sorted.find(function(t){ return Number(t.unreadUser || 0) > 0; });
  var open   = sorted.find(function(t){ return !t.resuelto; });
  var best   = unread || open || sorted[0] || null;
  if (!req) return best;
  // Si el solicitado está cerrado y existe uno abierto/no leído, cambiar automáticamente.
  if (req.resuelto && best && best.chatId !== req.chatId) return best;
  // Si hay una respuesta nueva de soporte en otro ticket, priorizarla.
  if (unread && unread.chatId !== req.chatId) return unread;
  return req;
}

function apiSoporteMobileSync(payload) {
  var token = String((payload && payload.token) || "");
  requireSession_(token);
  var requested = String((payload && payload.chatId) || "").trim();
  var ticketsRes = apiGetMisTickets({ token: token });
  var tickets = (ticketsRes && ticketsRes.tickets) ? ticketsRes.tickets : [];
  var picked = _pickBestUserTicketV62_(tickets, requested);
  var chosen = picked ? picked.chatId : (requested || "");
  var hist = chosen ? apiGetSoporteHistorial({ token: token, chatId: chosen }) : { ok:true, items:[], chatId:"", latestTs:"" };
  return {
    ok:true,
    tickets:tickets,
    activeTicket:picked || (hist && hist.ticket) || null,
    chatId:(hist && hist.chatId) || chosen,
    historial:hist,
    latestTs:(hist && hist.latestTs) || ""
  };
}

/* V8 — REPORTES PDF PARA SUPERVISORES */
function _isSupportOrAdminV8_(role){ var r=String(role||'').toUpperCase(); return isAdminRole_(r)||r==='SOPORTE'; }
function _supAllowedSedesV8_(token, role){
  if(_isSupportOrAdminV8_(role)) return getSedes_().map(function(s){return s.sede;});
  try{ var a=apiGetSedeJornadaUsuario({token:token}).asignaciones||[]; var m={}; a.forEach(function(x){if(x&&x.sede)m[String(x.sede).trim()]=true;}); return Object.keys(m).sort(); }catch(e){return [];}
}
function _supEmpleadoMapV8_(sedesAllowed){
  var allowed={}; (sedesAllowed||[]).forEach(function(s){allowed[String(s||'').trim()]=true;});
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CONTRATOS), byId={}, bySede={};
  if(!sh||sh.getLastRow()<2) return {byId:byId,bySede:bySede};
  var v=sh.getDataRange().getValues(), h=v[0].map(normHdr_), idx={}; h.forEach(function(x,i){idx[x]=i;});
  var cID=idx['ID'], cS=idx['SEDE'], cN=idx['NOMBRE_TRABAJADOR'];
  var cT=idx['JORNADA']!==undefined?idx['JORNADA']:(idx['TURNO']!==undefined?idx['TURNO']:-1);
  var cST=idx[COL_STATUS_TRAB];
  for(var i=1;i<v.length;i++){
    var sede=String(v[i][cS]||'').trim(); if(!sede||!allowed[sede]) continue;
    var id=formatCellValue_(v[i][cID]); if(!id) continue;
    var st=cST!==undefined?String(v[i][cST]||'ACTIVO').toUpperCase().trim():'ACTIVO'; if(st==='BAJA') continue;
    var obj={id:id,nombre:String(v[i][cN]||'').trim(),sede:sede,jornada:cT>=0?String(v[i][cT]||'').trim().toUpperCase():''};
    byId[id]=obj; if(!bySede[sede]) bySede[sede]=[]; bySede[sede].push(obj);
  }
  Object.keys(bySede).forEach(function(s){bySede[s].sort(function(a,b){return a.nombre.localeCompare(b.nombre,'es');});});
  return {byId:byId,bySede:bySede};
}
function apiSupervisorReporteOpciones(payload){
  var token=String((payload&&payload.token)||''), ses=requireSession_(token);
  var sedes=_supAllowedSedesV8_(token,ses.role); var emp=_supEmpleadoMapV8_(sedes);
  var sede=String((payload&&payload.sede)||sedes[0]||'').trim();
  return {ok:true,role:ses.role,sedes:sedes,empleados:emp.bySede[sede]||[],sede:sede};
}
function _supReportRangeV8_(payload){
  var tipo=String((payload&&payload.periodo)||'custom').toLowerCase(), today=todayISO_();
  var y=Number(today.substring(0,4)), m=Number(today.substring(5,7)), d=Number(today.substring(8,10));
  function iso(dt){return Utilities.formatDate(dt,TZ,'yyyy-MM-dd');}
  if(tipo==='q_actual') return {ini:iso(d<=15?new Date(y,m-1,1):new Date(y,m-1,16)),fin:iso(d<=15?new Date(y,m-1,15):new Date(y,m,0)),label:'Quincena actual'};
  if(tipo==='q_pasada') return d<=15?{ini:iso(new Date(y,m-2,16)),fin:iso(new Date(y,m-1,0)),label:'Quincena pasada'}:{ini:iso(new Date(y,m-1,1)),fin:iso(new Date(y,m-1,15)),label:'Quincena pasada'};
  var ini=String((payload&&payload.fechaInicio)||'').trim(), fin=String((payload&&payload.fechaFin)||'').trim();
  if(!isoValid_(ini)||!isoValid_(fin)) throw new Error('Selecciona un rango válido.');
  if(ini>fin) throw new Error('La fecha inicial no puede ser mayor a la final.');
  return {ini:ini,fin:fin,label:'Rango personalizado'};
}
function _supBuildReportHtmlV8_(data){
  function e(s){return esc_(s);} var rows=data.rows||[], counts=data.counts||{};
  var codeHtml=Object.keys(counts).sort().map(function(k){return '<div class="chip"><b>'+e(k)+'</b><span>'+counts[k]+'</span></div>';}).join('')||'<div class="muted">Sin códigos registrados.</div>';
  var trs=rows.map(function(r){
    var statusClass = r.codigo === 'SIN REGISTRO' ? 'no' : (r.codigo === 'A' || r.codigo === 'AF' ? 'ok' : (r.codigo === 'F' || r.codigo === 'PSG' ? 'bad' : 'mid'));
    return '<tr><td>'+e(r.fecha)+'</td><td>'+e(r.dia||'')+'</td><td>'+e(r.sede)+'</td><td>'+e(r.jornada||'')+'</td><td>'+e(r.id)+'</td><td>'+e(r.nombre)+'</td><td><span class="code '+statusClass+'">'+e(r.codigo)+'</span></td><td>'+e(r.capturadoPor||'—')+'</td><td>'+e(r.hora||'—')+'</td></tr>';
  }).join('')||'<tr><td colspan="9" class="empty">No se encontraron trabajadores o registros en el período seleccionado.</td></tr>';
  var gen=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd HH:mm:ss');
  return '<!doctype html><html><head><meta charset="utf-8"><style>'+ 
    '@page{size:letter landscape;margin:18px}body{font-family:Arial,sans-serif;color:#111827;margin:0}.head{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #0f172a;padding-bottom:12px;margin-bottom:12px}h1{font-size:20px;margin:0 0 4px}.sub{font-size:11px;color:#475569;line-height:1.45}.badge{background:#dbeafe;color:#1d4ed8;border-radius:999px;padding:5px 10px;font-size:11px;font-weight:bold}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:12px 0}.box{border:1px solid #cbd5e1;border-radius:10px;padding:9px}.box .k{font-size:9px;color:#64748b;text-transform:uppercase}.box .v{font-size:15px;font-weight:bold;margin-top:4px}.chips{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}.chip{border:1px solid #cbd5e1;border-radius:999px;padding:5px 9px;font-size:11px}.chip b{margin-right:6px}table{width:100%;border-collapse:collapse;margin-top:10px;font-size:9.5px}th{background:#0f172a;color:white;text-align:left;padding:7px 5px}td{border-bottom:1px solid #e2e8f0;padding:6px 5px;vertical-align:top}tr:nth-child(even) td{background:#f8fafc}.code{display:inline-block;border-radius:999px;padding:3px 7px;font-weight:bold;border:1px solid #cbd5e1;background:#f8fafc;color:#334155}.code.ok{background:#dcfce7;color:#166534;border-color:#86efac}.code.bad{background:#fee2e2;color:#991b1b;border-color:#fca5a5}.code.mid{background:#fef3c7;color:#92400e;border-color:#fcd34d}.code.no{background:#e2e8f0;color:#475569;border-color:#cbd5e1}.empty{text-align:center;color:#64748b;padding:20px}.muted{color:#64748b;font-size:11px}.note{margin-top:8px;font-size:10px;color:#64748b;line-height:1.4}'+
    '</style></head><body><div class="head"><div><h1>MHS Integradora · Reporte de asistencia</h1><div class="sub">'+e(data.subtitle)+'<br>Periodo: '+e(data.ini)+' → '+e(data.fin)+' · Generado: '+e(gen)+'</div></div><div class="badge">'+e(data.reportNo)+'</div></div><div class="grid"><div class="box"><div class="k">Tipo</div><div class="v">'+e(data.tipoLabel)+'</div></div><div class="box"><div class="k">Sede</div><div class="v">'+e(data.sede||'Todas')+'</div></div><div class="box"><div class="k">Trabajador</div><div class="v">'+e(data.trabajadorLabel||'Todos')+'</div></div><div class="box"><div class="k">Filas</div><div class="v">'+rows.length+'</div></div></div><div class="sub"><b>Resumen por asistencia</b></div><div class="chips">'+codeHtml+'</div><div class="note">Este PDF muestra cada día del período y cada trabajador consultado. La columna <b>Quién marcó</b> indica el usuario que capturó la asistencia. Si no existe registro para ese trabajador en ese día, se muestra <b>SIN REGISTRO</b>.</div><table><thead><tr><th>Fecha</th><th>Día</th><th>Sede</th><th>Jornada</th><th>ID</th><th>Trabajador</th><th>Asistencia</th><th>Quién marcó</th><th>Hora</th></tr></thead><tbody>'+trs+'</tbody></table></body></html>';
}

function apiSupervisorReportePDF(payload){
  var token=String((payload&&payload.token)||''), ses=requireSession_(token), allowed=_supAllowedSedesV8_(token,ses.role);
  if(!allowed.length) throw new Error('No tienes sedes asignadas para consultar.');
  var allowedMap={}; allowed.forEach(function(s){allowedMap[s]=true;});
  var tipo=String((payload&&payload.tipo)||'sede').toLowerCase(), sede=String((payload&&payload.sede)||allowed[0]||'').trim();
  if(!allowedMap[sede]) throw new Error('No tienes permiso para consultar esta sede.');
  var jornada=String((payload&&payload.jornada)||'').trim().toUpperCase();
  var trabajador=String((payload&&payload.trabajador)||'').trim().toLowerCase();
  var range=_supReportRangeV8_(payload), emp=_supEmpleadoMapV8_(allowed), empById=emp.byId;
  var empleados=(emp.bySede[sede]||[]).slice();
  if(jornada){
    empleados=empleados.filter(function(x){return String(x.jornada||'').toUpperCase().indexOf(jornada)>=0;});
  }
  if(tipo==='trabajador'){
    if(!trabajador) throw new Error('Escribe el trabajador o ID.');
    empleados=empleados.filter(function(x){return (String(x.id||'')+' '+String(x.nombre||'')).toLowerCase().indexOf(trabajador)>=0;});
  }
  if(!empleados.length) throw new Error('No se encontraron trabajadores para ese filtro.');

  var regMap={}, sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  if(sh&&sh.getLastRow()>=2){
    var v=sh.getDataRange().getValues();
    for(var i=1;i<v.length;i++){
      var f=formatCellValue_(v[i][0]); if(f<range.ini||f>range.fin) continue;
      var rs=String(v[i][2]||'').trim(); if(rs!==sede) continue;
      var id=formatCellValue_(v[i][4]); if(!id) continue;
      var info=empById[id]||null; if(!info) continue;
      var j=String(info.jornada||'').toUpperCase(); if(jornada&&j.indexOf(jornada)<0) continue;
      var key=f+'|'+id;
      regMap[key]={
        codigo:String(v[i][6]||'').trim().toUpperCase()||'-',
        capturadoPor:String(v[i][7]||'').trim(),
        ts:formatCellValue_(v[i][8])
      };
    }
  }

  function daysBetween_(ini,fin){
    var out=[], p1=ini.split('-'), p2=fin.split('-');
    var d=new Date(Number(p1[0]),Number(p1[1])-1,Number(p1[2]));
    var endD=new Date(Number(p2[0]),Number(p2[1])-1,Number(p2[2]));
    while(d<=endD){out.push(Utilities.formatDate(d,TZ,'yyyy-MM-dd')); d.setDate(d.getDate()+1);}
    return out;
  }
  var dias=daysBetween_(range.ini,range.fin);
  var diaNames=['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
  var counts={}, rows=[];
  dias.forEach(function(f){
    var pp=f.split('-'), dd=new Date(Number(pp[0]),Number(pp[1])-1,Number(pp[2]));
    empleados.forEach(function(empObj){
      var r=regMap[f+'|'+empObj.id];
      var cod=r?r.codigo:'SIN REGISTRO'; counts[cod]=(counts[cod]||0)+1;
      var hora='';
      if(r&&r.ts){ var m=String(r.ts).match(/(\d{2}:\d{2})/); hora=m?m[1]:String(r.ts); }
      rows.push({fecha:f,dia:diaNames[dd.getDay()],sede:sede,jornada:empObj.jornada||'',id:empObj.id,nombre:empObj.nombre,codigo:cod,capturadoPor:r?r.capturadoPor:'',hora:hora});
    });
  });
  // Mantener PDF razonable y evitar tiempos excesivos en rangos muy amplios.
  if(rows.length>5000) throw new Error('El reporte es muy grande ('+rows.length+' filas). Reduce el rango, sede, jornada o elige un trabajador.');
  var reportNo='RPT-'+Utilities.formatDate(new Date(),TZ,'yyyyMMdd-HHmmss');
  var trabajadorLabel=tipo==='trabajador'?(empleados.length===1?(empleados[0].id+' · '+empleados[0].nombre):('Filtro: '+trabajador)):'Todos';
  var html=_supBuildReportHtmlV8_({reportNo:reportNo,tipoLabel:tipo==='trabajador'?'Trabajador':'Sede general',subtitle:(tipo==='trabajador'?'Asistencia histórica individual':'Asistencia histórica por sede'),ini:range.ini,fin:range.fin,sede:sede,trabajadorLabel:trabajadorLabel,rows:rows,counts:counts});
  var blob=HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF);
  var safe=(sede+'_'+(tipo==='trabajador'?trabajador:'sede')+'_'+range.ini+'_'+range.fin).replace(/[^\w\-]+/g,'_');
  blob.setName('Reporte_Asistencia_'+safe+'.pdf');
  var file=DriveApp.getFolderById(DRIVE_PDF_FOLDER_ID).createFile(blob);
  return {ok:true,url:file.getUrl(),message:'PDF generado correctamente.',total:rows.length,reportNo:reportNo};
}



/*******************************************************
 * V9 · RH MÓVIL + PAGOS PENDIENTES
 * Capa ligera: no toca guardado de asistencia ni tickets.
 *******************************************************/
const SH_PAGOS_PEND = "PAGOS_PENDIENTES";
const HDR_PAGOS_PEND = ["FOLIO","ESTADO","FECHA_REGISTRO","QUINCENA","SEDE","ID","NOMBRE","MONTO","MOTIVO","OBSERVACIONES","CREADO_POR","TS_CREACION","PAGADO_POR","TS_PAGO"];

function _ensurePagosPendientes_(){
  return _ensureSheetIfMissing(SpreadsheetApp.getActiveSpreadsheet(), SH_PAGOS_PEND, [HDR_PAGOS_PEND]);
}
function _isRHRoleV9_(role){
  var r=String(role||'').toUpperCase();
  return r==='SUPERADMIN'||r==='ADMIN'||r==='CEO';
}
function _nextPagoFolioV9_(sh){
  var n=Math.max(0, sh.getLastRow()-1)+1;
  return 'PP-'+String(n).padStart(4,'0');
}
function apiRHMobilePanel(payload){
  var token=String((payload&&payload.token)||'');
  var ses=requireSession_(token);
  if(!_isRHRoleV9_(ses.role) && String(ses.role).toUpperCase()!=='SOPORTE') throw new Error('Sin permiso para RH móvil.');
  var sedes=getSedes_();
  var users=readUsers_().map(function(u){return {username:u.username,role:u.role,fullName:u.fullName,active:u.active};});
  var pagos={total:0,pendientes:0,pagados:0,montoPendiente:0};
  try{
    var sh=_ensurePagosPendientes_();
    if(sh.getLastRow()>=2){
      var v=sh.getDataRange().getValues(), h=v[0].map(normHdr_), idx={}; h.forEach(function(x,i){idx[x]=i;});
      for(var i=1;i<v.length;i++){
        pagos.total++;
        var est=String(v[i][idx.ESTADO]||'PENDIENTE').toUpperCase();
        var monto=Number(v[i][idx.MONTO]||0)||0;
        if(est==='PAGADO') pagos.pagados++; else {pagos.pendientes++; pagos.montoPendiente+=monto;}
      }
    }
  }catch(e){}
  return {ok:true,role:ses.role,username:ses.username,sedes:sedes,users:users,pagos:pagos,today:todayISO_()};
}
function apiRHMobileEmpleados(payload){
  var token=String((payload&&payload.token)||''); var ses=requireSession_(token);
  if(!_isRHRoleV9_(ses.role)) throw new Error('Solo ADMIN/SUPERADMIN/CEO.');
  var sede=String((payload&&payload.sede)||'').trim();
  var q=String((payload&&payload.q)||'').toUpperCase().trim();
  var out=[];
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CONTRATOS);
  if(!sh||sh.getLastRow()<2) return {ok:true,empleados:[]};
  var v=sh.getDataRange().getValues(), h=v[0].map(normHdr_), idx={}; h.forEach(function(x,i){idx[x]=i;});
  for(var i=1;i<v.length;i++){
    var st=String(v[i][idx[COL_STATUS_TRAB]]||'ACTIVO').toUpperCase();
    var s=String(v[i][idx.SEDE]||'').trim(); if(sede&&s!==sede) continue;
    var id=formatCellValue_(v[i][idx.ID]); var nom=String(v[i][idx.NOMBRE_TRABAJADOR]||'').trim();
    if(q && (id+' '+nom+' '+s).toUpperCase().indexOf(q)<0) continue;
    out.push({id:id,nombre:nom,sede:s,jornada:String(v[i][idx.JORNADA]||v[i][idx.TURNO]||'').trim(),status:st});
    if(out.length>=250) break;
  }
  return {ok:true,empleados:out};
}
function apiPagoPendienteCrear(payload){
  var token=String((payload&&payload.token)||''); var ses=requireSession_(token);
  if(!_isRHRoleV9_(ses.role)) throw new Error('Solo ADMIN/SUPERADMIN puede registrar pagos pendientes.');
  var sede=String((payload&&payload.sede)||'').trim();
  var id=String((payload&&payload.id)||'').trim();
  var nombre=String((payload&&payload.nombre)||'').trim();
  var monto=Number((payload&&payload.monto)||0)||0;
  var motivo=String((payload&&payload.motivo)||'').trim();
  var obs=String((payload&&payload.observaciones)||'').trim();
  var quincena=String((payload&&payload.quincena)||'').trim();
  if(!sede||!id||!nombre||!motivo) throw new Error('Faltan sede, ID, nombre o motivo.');
  if(monto<=0) throw new Error('El monto debe ser mayor a 0.');
  var sh=_ensurePagosPendientes_(); var folio=_nextPagoFolioV9_(sh); var ts=isoNow_();
  sh.appendRow([folio,'PENDIENTE',todayISO_(),quincena,sede,id,nombre,monto,motivo,obs,ses.username,ts,'','']);
  return {ok:true,folio:folio,message:'Pago pendiente creado: '+folio};
}
function apiPagosPendientesListar(payload){
  var token=String((payload&&payload.token)||''); var ses=requireSession_(token);
  if(!_isRHRoleV9_(ses.role)) throw new Error('Solo ADMIN/SUPERADMIN/CEO.');
  var estado=String((payload&&payload.estado)||'PENDIENTE').toUpperCase();
  var sedeF=String((payload&&payload.sede)||'').trim();
  var sh=_ensurePagosPendientes_(); var out=[];
  if(sh.getLastRow()<2) return {ok:true,items:[]};
  var v=sh.getDataRange().getValues(), h=v[0].map(normHdr_), idx={}; h.forEach(function(x,i){idx[x]=i;});
  for(var i=v.length-1;i>=1;i--){
    var est=String(v[i][idx.ESTADO]||'PENDIENTE').toUpperCase();
    var sede=String(v[i][idx.SEDE]||'').trim();
    if(estado && estado!=='TODOS' && est!==estado) continue;
    if(sedeF && sede!==sedeF) continue;
    out.push({folio:String(v[i][idx.FOLIO]||''),estado:est,fecha:formatCellValue_(v[i][idx.FECHA_REGISTRO]),quincena:String(v[i][idx.QUINCENA]||''),sede:sede,id:formatCellValue_(v[i][idx.ID]),nombre:String(v[i][idx.NOMBRE]||''),monto:Number(v[i][idx.MONTO]||0)||0,motivo:String(v[i][idx.MOTIVO]||''),observaciones:String(v[i][idx.OBSERVACIONES]||''),creadoPor:String(v[i][idx.CREADO_POR]||''),ts:String(v[i][idx.TS_CREACION]||''),pagadoPor:String(v[i][idx.PAGADO_POR]||''),tsPago:String(v[i][idx.TS_PAGO]||'')});
    if(out.length>=300) break;
  }
  return {ok:true,items:out};
}
function apiPagoPendienteMarcarPagado(payload){
  var token=String((payload&&payload.token)||''); var ses=requireSession_(token);
  if(!_isRHRoleV9_(ses.role)) throw new Error('Solo ADMIN/SUPERADMIN.');
  var folio=String((payload&&payload.folio)||'').trim(); if(!folio) throw new Error('Falta folio.');
  var sh=_ensurePagosPendientes_(); if(sh.getLastRow()<2) throw new Error('No hay pagos.');
  var v=sh.getDataRange().getValues(), h=v[0].map(normHdr_), idx={}; h.forEach(function(x,i){idx[x]=i;});
  for(var i=1;i<v.length;i++) if(String(v[i][idx.FOLIO]||'')===folio){
    sh.getRange(i+1, idx.ESTADO+1).setValue('PAGADO');
    sh.getRange(i+1, idx.PAGADO_POR+1).setValue(ses.username);
    sh.getRange(i+1, idx.TS_PAGO+1).setValue(isoNow_());
    return {ok:true,message:'Marcado como pagado: '+folio};
  }
  throw new Error('Folio no encontrado.');
}
function apiPagosPendientesPDF(payload){
  var token=String((payload&&payload.token)||''); var ses=requireSession_(token);
  if(!_isRHRoleV9_(ses.role)) throw new Error('Solo ADMIN/SUPERADMIN/CEO.');
  var data=apiPagosPendientesListar(payload).items||[];
  var total=data.reduce(function(a,x){return a+(Number(x.monto)||0);},0);
  var html='<!doctype html><html><head><meta charset="UTF-8"><style>body{font-family:Arial,sans-serif;color:#111}h1{font-size:18px;margin:0 0 4px}.sub{font-size:11px;color:#555;margin-bottom:14px}table{width:100%;border-collapse:collapse;font-size:10px}th,td{border:1px solid #ddd;padding:5px;text-align:left}th{background:#0f1d34;color:white}.right{text-align:right}.badge{font-weight:bold}</style></head><body>'+
    '<h1>MHS · Reporte de pagos pendientes</h1><div class="sub">Generado por '+ses.username+' · '+isoNow_()+' · Total: $'+total.toFixed(2)+'</div><table><thead><tr><th>Folio</th><th>Estado</th><th>Fecha</th><th>Quincena</th><th>Sede</th><th>ID</th><th>Trabajador</th><th>Motivo</th><th>Monto</th><th>Observaciones</th></tr></thead><tbody>'+
    data.map(function(x){return '<tr><td>'+x.folio+'</td><td class="badge">'+x.estado+'</td><td>'+x.fecha+'</td><td>'+x.quincena+'</td><td>'+x.sede+'</td><td>'+x.id+'</td><td>'+x.nombre+'</td><td>'+x.motivo+'</td><td class="right">$'+Number(x.monto||0).toFixed(2)+'</td><td>'+x.observaciones+'</td></tr>';}).join('')+
    '</tbody></table></body></html>';
  var blob=Utilities.newBlob(html,'text/html','pagos_pendientes.html').getAs('application/pdf').setName('Pagos_pendientes_'+todayISO_()+'.pdf');
  var folder=DriveApp.getFolderById(DRIVE_PDF_FOLDER_ID); var file=folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {ok:true,url:file.getUrl(),id:file.getId(),count:data.length,total:total};
}

// ══════════════════════════════════════════════════════
//  V16 SUPERVISORES PRO · Operación rápida + contexto RH
//  Apps Script puro. No toca login ni arquitectura base.
// ══════════════════════════════════════════════════════
var SH_V16_GEO_LOG = 'GEO_REGISTROS_ASISTENCIA';

function _v16EnsureGeoSheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_V16_GEO_LOG);
  var hdr = ['TS_ISO','FECHA_ISO','SEDE','JORNADA','USERNAME','LAT','LNG','UBICACION','TOTAL_REGISTROS'];
  if(!sh){
    sh = ss.insertSheet(SH_V16_GEO_LOG);
    sh.getRange(1,1,1,hdr.length).setValues([hdr]);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1,hdr.length);
  }
  return sh;
}

function _v16CanSeeSede_(ses, sede, jornada){
  var role = String((ses && ses.role) || '').toUpperCase();
  if(isAdminRole_(role) || role === 'SOPORTE') return true;
  try{
    var asign = apiGetSedeJornadaUsuario({token:'__none__'}); // intentionally fails into catch if signature changes
  }catch(e){}
  // Lectura directa usando el token no es posible aquí; se valida en callers por token.
  return true;
}

function _v16RequireSupervisorAccess_(token, sede, jornada){
  var ses = requireSession_(token);
  var role = String(ses.role || '').toUpperCase();
  if(isAdminRole_(role) || role === 'SOPORTE') return ses;
  var asign = [];
  try{ asign = apiGetSedeJornadaUsuario({token:token}).asignaciones || []; }catch(e){ asign = []; }
  if(!asign.length) return ses;
  var sedeN = String(sede || '').trim();
  var jornN = String(jornada || '').trim().toUpperCase();
  var ok = asign.some(function(a){
    var sOk = String(a.sede || '').trim() === sedeN;
    var j = String(a.jornada || '').trim().toUpperCase();
    return sOk && (!jornN || !j || j === jornN);
  });
  if(!ok) throw new Error('No tienes asignada esta sede/jornada.');
  return ses;
}

function _v16AddDaysISO_(iso, delta){
  var p = String(iso || todayISO_()).split('-').map(Number);
  var d = new Date(p[0], p[1]-1, p[2] + delta);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function _v16CodeGroup_(cod){
  cod = String(cod || '').toUpperCase();
  if(cod === 'A' || cod === 'AF') return 'ASISTENCIA';
  if(cod === 'F') return 'FALTA';
  if(cod === 'DS') return 'DESCANSO';
  if(['PCG','PSG','I','DT','FER','INH'].indexOf(cod) >= 0) return 'INCIDENCIA';
  return 'OTRO';
}

function apiSupervisorTrabajadorFicha(payload){
  payload = payload || {};
  var token   = String(payload.token || '');
  var sede    = String(payload.sede || '').trim();
  var jornada = String(payload.jornada || '').trim().toUpperCase();
  var id      = String(payload.id || '').trim();
  var dias    = Math.max(15, Math.min(180, Number(payload.dias || 90)));
  if(!id) throw new Error('Falta ID de trabajador.');
  _v16RequireSupervisorAccess_(token, sede, jornada);

  var empleados = sede ? apiGetEmpleadosPorSedeJornada({token:token, sede:sede, jornada:jornada}) : [];
  var emp = empleados.filter(function(e){ return String(e.id) === id; })[0] || {id:id, nombre:String(payload.nombre||''), turno:jornada};
  var fechaFin = String(payload.fechaFin || todayISO_()).trim();
  var fechaIni = _v16AddDaysISO_(fechaFin, -dias);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  var registros = [];
  var resumen = {asistencias:0, faltas:0, incidencias:0, descansos:0, total:0};
  if(sh && sh.getLastRow() >= 2){
    var v = sh.getDataRange().getValues();
    for(var i=1;i<v.length;i++){
      var f = formatCellValue_(v[i][0]);
      if(f < fechaIni || f > fechaFin) continue;
      var s = String(v[i][2] || '').trim();
      var rowId = formatCellValue_(v[i][4]);
      if(rowId !== id) continue;
      if(sede && s !== sede) continue;
      var cod = String(v[i][6] || '').trim().toUpperCase();
      var cap = String(v[i][7] || '').trim();
      var ts  = formatCellValue_(v[i][8]);
      var grp = _v16CodeGroup_(cod);
      if(grp === 'ASISTENCIA') resumen.asistencias++;
      else if(grp === 'FALTA') resumen.faltas++;
      else if(grp === 'INCIDENCIA') resumen.incidencias++;
      else if(grp === 'DESCANSO') resumen.descansos++;
      resumen.total++;
      registros.push({fecha:f, sede:s, codigo:cod, grupo:grp, capturadoPor:cap, ts:ts});
    }
  }
  registros.sort(function(a,b){ return String(b.fecha).localeCompare(String(a.fecha)); });
  return {ok:true, trabajador:emp, resumen:resumen, registros:registros.slice(0,80), rango:{inicio:fechaIni, fin:fechaFin}};
}

function apiSupervisorCopiarJornada(payload){
  payload = payload || {};
  var token   = String(payload.token || '');
  var sede    = String(payload.sede || '').trim();
  var jornada = String(payload.jornada || '').trim().toUpperCase();
  var fecha   = String(payload.fecha || todayISO_()).trim();
  if(!sede) throw new Error('Falta sede.');
  if(!isoValid_(fecha)) throw new Error('Fecha inválida.');
  _v16RequireSupervisorAccess_(token, sede, jornada);

  var empleados = apiGetEmpleadosPorSedeJornada({token:token, sede:sede, jornada:jornada});
  var ids = {};
  empleados.forEach(function(e){ ids[String(e.id)] = true; });
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PL);
  if(!sh || sh.getLastRow() < 2) return {ok:true, fechaOrigen:'', marcas:{}, total:0};
  var v = sh.getDataRange().getValues();
  var porFecha = {};
  for(var i=1;i<v.length;i++){
    var f = formatCellValue_(v[i][0]);
    if(!f || f >= fecha) continue;
    var s = String(v[i][2] || '').trim();
    if(s !== sede) continue;
    var id = formatCellValue_(v[i][4]);
    if(!ids[id]) continue;
    var cod = String(v[i][6] || '').trim().toUpperCase();
    if(!cod) continue;
    if(!porFecha[f]) porFecha[f] = {};
    porFecha[f][id] = cod;
  }
  var fechas = Object.keys(porFecha).sort().reverse();
  if(!fechas.length) return {ok:true, fechaOrigen:'', marcas:{}, total:0};
  var origen = fechas[0];
  var marcas = porFecha[origen] || {};
  return {ok:true, fechaOrigen:origen, marcas:marcas, total:Object.keys(marcas).length};
}

function apiSupervisorResumenJornada(payload){
  payload = payload || {};
  var token   = String(payload.token || '');
  var sede    = String(payload.sede || '').trim();
  var jornada = String(payload.jornada || '').trim().toUpperCase();
  var fecha   = String(payload.fecha || todayISO_()).trim();
  _v16RequireSupervisorAccess_(token, sede, jornada);
  var empleados = apiGetEmpleadosPorSedeJornada({token:token, sede:sede, jornada:jornada});
  var marks = apiGetPaseListaDia(fecha, sede) || {};
  var out = {total:empleados.length, marcados:0, pendientes:0, asistencias:0, faltas:0, incidencias:0, descansos:0};
  empleados.forEach(function(e){
    var cod = String(marks[e.id] || '').toUpperCase();
    if(!cod){ out.pendientes++; return; }
    out.marcados++;
    var grp = _v16CodeGroup_(cod);
    if(grp === 'ASISTENCIA') out.asistencias++;
    else if(grp === 'FALTA') out.faltas++;
    else if(grp === 'INCIDENCIA') out.incidencias++;
    else if(grp === 'DESCANSO') out.descansos++;
  });
  return {ok:true, resumen:out};
}

function _v16LogGeoFromPayload_(payload, ses){
  try{
    payload = payload || {};
    var lista = Array.isArray(payload.lista) ? payload.lista : [];
    if(!lista.length) return;
    var geo = null;
    for(var i=0;i<lista.length;i++){
      if(lista[i] && (lista[i].lat || lista[i].lng || lista[i].ubicacion)){ geo = lista[i]; break; }
    }
    if(!geo) return;
    var sh = _v16EnsureGeoSheet_();
    sh.appendRow([
      isoNow_(),
      String(payload.fecha || '').trim(),
      String(payload.sede || '').trim(),
      String(payload.jornada || payload.turno || '').trim(),
      String((ses && ses.username) || ''),
      String(geo.lat || ''),
      String(geo.lng || ''),
      String(geo.ubicacion || ''),
      lista.length
    ]);
  }catch(e){ Logger.log('V16 GEO LOG: ' + e); }
}

// Override seguro: conserva guardado original y añade registro de ubicación no bloqueante.
function apiGuardarPaseListaMobile(payload){
  payload = payload || {};
  var token = String(payload.token || '');
  var ses = requireSession_(token);
  var res = apiGuardarPaseLista(payload);
  _v16LogGeoFromPayload_(payload, ses);
  try{
    var fecha = String(payload.fecha || todayISO_()).trim();
    res.pendientes = apiGetPendientesAsistencia({token:token, fecha:fecha});
    res.resumenJornada = apiSupervisorResumenJornada({token:token, fecha:fecha, sede:String(payload.sede||''), jornada:String(payload.jornada||'')}).resumen;
  }catch(e){ res.pendientesError = String(e && e.message || e); }
  return res;
}

// ══════════════════════════════════════════════════════════════
//  getJornadasPorSede_ — Filtro de jornadas por supervisor
//  Lee hoja "Asignaciones": col A=username, B=sede, C=jornada
//  Retorna { "OHORAN": ["MATUTINO","VESPERTINO"], ... }
// ══════════════════════════════════════════════════════════════
function getJornadasPorSede_(username) {
  try {
    var ss    = getMHSSpreadsheet_();
    var sheet = ss.getSheetByName('Asignaciones');
    if (!sheet) return {};
    var rows = sheet.getDataRange().getValues();
    var map  = {};
    var user = String(username || '').toLowerCase().trim();
    for (var i = 1; i < rows.length; i++) {
      var u = String(rows[i][0] || '').toLowerCase().trim();
      var s = String(rows[i][1] || '').toUpperCase().trim();
      var j = String(rows[i][2] || '').toUpperCase().trim();
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
