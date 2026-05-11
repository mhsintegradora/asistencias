/**
 * MHS Integradora — turno-filter.js  V1
 * ──────────────────────────────────────────────────────────────────────────
 * Filtra las jornadas disponibles en la pantalla de Setup según lo que
 * tenga asignado el supervisor actual por sede.
 *
 * Ejemplo:  Fer  → OHORAN: [MATUTINO, VESPERTINO]
 *           Ivan → OHORAN: [VESPERTINO, NOCTURNO]
 *           Alex → OHORAN: [NOCTURNO]
 *
 * El backend (apiMobileSessionBootstrap) debe devolver:
 *   res.jornadasPorSede = { "OHORAN": ["MATUTINO","VESPERTINO"], ... }
 *
 * Si el backend todavía no envía ese campo se usa el mapa de fallback
 * definido al final de este archivo (SUPERVISOR_JORNADAS_FALLBACK).
 * ──────────────────────────────────────────────────────────────────────────
 */
(function () {
  'use strict';

  /* ── Estado del módulo ── */
  var _jornadasPorSede  = {};   // { "SEDE": ["MATUTINO", ...] }
  var _sedeSelectReady  = false;
  var _initDone         = false;

  /* ── Fallback manual mientras GAS no devuelva jornadasPorSede ──────────
   * Clave: username en MINÚSCULAS tal como lo guarda getUser().
   * Modifica esto según tus supervisores reales.
   * ─────────────────────────────────────────────────────────────────────*/
  var SUPERVISOR_JORNADAS_FALLBACK = {
    'fer':  {
      'OHORAN':   ['MATUTINO',  'VESPERTINO'],
      'DEFAULT':  ['MATUTINO',  'VESPERTINO']
    },
    'ivan': {
      'OHORAN':   ['VESPERTINO','NOCTURNO'],
      'DEFAULT':  ['VESPERTINO','NOCTURNO']
    },
    'alex': {
      'OHORAN':   ['NOCTURNO'],
      'DEFAULT':  ['NOCTURNO']
    }
  };

  var TODAS_LAS_JORNADAS = ['MATUTINO', 'VESPERTINO', 'NOCTURNO'];

  /* ──────────────────────────────────────────────────────────────────────
   * Helpers
   * ────────────────────────────────────────────────────────────────────*/
  function getJornadasParaSede(sedeValue) {
    /* 1) Lo que vino del bootstrap */
    if (_jornadasPorSede && Object.keys(_jornadasPorSede).length) {
      return _jornadasPorSede[sedeValue] || TODAS_LAS_JORNADAS;
    }
    /* 2) Fallback por usuario */
    var user = (typeof getUser === 'function' ? getUser() : '').toLowerCase().trim();
    var supMap = SUPERVISOR_JORNADAS_FALLBACK[user];
    if (supMap) {
      return supMap[sedeValue] || supMap['DEFAULT'] || TODAS_LAS_JORNADAS;
    }
    /* 3) Sin restricción */
    return TODAS_LAS_JORNADAS;
  }

  function labelJornada(j) {
    var map = {
      MATUTINO:   'Matutino',
      VESPERTINO: 'Vespertino',
      NOCTURNO:   'Nocturno'
    };
    return map[j] || (j.charAt(0).toUpperCase() + j.slice(1).toLowerCase());
  }

  /* ──────────────────────────────────────────────────────────────────────
   * Inyectar el <select> de jornada en la pantalla de Setup
   * ────────────────────────────────────────────────────────────────────*/
  function ensureJornadaSelect() {
    if (document.getElementById('mJornadaFieldGroup')) return true;

    var sedeSelect = document.getElementById('mSedeSelect');
    if (!sedeSelect) return false;

    var sedeField = sedeSelect.closest
      ? sedeSelect.closest('.field-group')
      : (function () {
          var n = sedeSelect;
          while (n && !(n.classList && n.classList.contains('field-group'))) n = n.parentNode;
          return n;
        })();
    if (!sedeField) return false;

    var fieldGroup = document.createElement('div');
    fieldGroup.id        = 'mJornadaFieldGroup';
    fieldGroup.className = 'field-group';
    fieldGroup.innerHTML =
      '<div class="field-label">Jornada</div>' +
      '<div class="select-wrap">' +
        '<select id="mJornadaSelect" onchange="mhsTurnoFilterOnJornadaChange()"></select>' +
      '</div>';

    /* Insertar después del hint de jornada (si existe) o del campo de sede */
    var hint     = document.getElementById('setupJornadaHint');
    var ancla    = hint || sedeField;
    var siguiente = ancla.nextSibling;
    if (siguiente) {
      ancla.parentNode.insertBefore(fieldGroup, siguiente);
    } else {
      ancla.parentNode.appendChild(fieldGroup);
    }

    return true;
  }

  /* ──────────────────────────────────────────────────────────────────────
   * Rellenar el select de jornadas según la sede elegida
   * ────────────────────────────────────────────────────────────────────*/
  function updateJornadaSelect(sedeValue) {
    var sel = document.getElementById('mJornadaSelect');
    if (!sel) return;

    var jornadas = sedeValue ? getJornadasParaSede(sedeValue) : [];

    sel.innerHTML = '';

    if (!jornadas.length || !sedeValue) {
      sel.innerHTML = '<option value="">— Selecciona una sede primero —</option>';
      var fg = document.getElementById('mJornadaFieldGroup');
      if (fg) fg.style.display = 'none';
      return;
    }

    var fg = document.getElementById('mJornadaFieldGroup');
    if (fg) fg.style.display = '';

    /* Si tiene más de una jornada, mostrar opción vacía primero */
    if (jornadas.length > 1) {
      var blank = document.createElement('option');
      blank.value = '';
      blank.textContent = 'Selecciona jornada…';
      sel.appendChild(blank);
    }

    jornadas.forEach(function (j) {
      var opt = document.createElement('option');
      opt.value       = j;
      opt.textContent = labelJornada(j);
      sel.appendChild(opt);
    });

    /* Con una sola jornada asignada, seleccionar automáticamente */
    if (jornadas.length === 1) {
      sel.value = jornadas[0];
      window.mJornadaAsignada = jornadas[0];
      /* Mostrar hint */
      var hint = document.getElementById('setupJornadaHint');
      if (hint) {
        hint.textContent = '🕐 Jornada asignada: ' + labelJornada(jornadas[0]);
        hint.style.display = '';
      }
    }
  }

  /* ──────────────────────────────────────────────────────────────────────
   * Exponer callback público para el onchange del select
   * ────────────────────────────────────────────────────────────────────*/
  window.mhsTurnoFilterOnJornadaChange = function () {
    var sel = document.getElementById('mJornadaSelect');
    if (!sel) return;
    window.mJornadaAsignada = sel.value;
  };

  /* ──────────────────────────────────────────────────────────────────────
   * Hook: mOnSedeChange  (cuando el usuario cambia la sede en Setup)
   * ────────────────────────────────────────────────────────────────────*/
  function hookSedeChange() {
    var orig = window.mOnSedeChange;
    window.mOnSedeChange = function () {
      if (typeof orig === 'function') orig.apply(this, arguments);
      var sedeVal = (document.getElementById('mSedeSelect') || {}).value || '';
      ensureJornadaSelect();
      updateJornadaSelect(sedeVal);
    };
  }

  /* ──────────────────────────────────────────────────────────────────────
   * Hook: mCargar  (botón "Cargar pase →")
   * Inyecta la jornada seleccionada antes de llamar a la API.
   * ────────────────────────────────────────────────────────────────────*/
  function hookMCargar() {
    var orig = window.mCargar;
    if (typeof orig !== 'function') return false;
    window.mCargar = function () {
      var sel = document.getElementById('mJornadaSelect');
      if (sel && sel.value) {
        window.mJornadaAsignada = sel.value;
      }
      /* Validar que se haya elegido jornada si hay varias disponibles */
      var sede = (document.getElementById('mSedeSelect') || {}).value || '';
      if (sede) {
        var jorns = getJornadasParaSede(sede);
        if (jorns.length > 1 && !window.mJornadaAsignada) {
          if (typeof showToast === 'function') showToast('Selecciona una jornada antes de continuar', 'warn');
          return;
        }
      }
      return orig.apply(this, arguments);
    };
    return true;
  }

  /* ──────────────────────────────────────────────────────────────────────
   * Hook: mApplyBootstrap_  (recibe datos del bootstrap de sesión)
   * Captura jornadasPorSede si el backend lo devuelve.
   * ────────────────────────────────────────────────────────────────────*/
  function hookBootstrap() {
    var orig = window.mApplyBootstrap_;
    if (typeof orig !== 'function') return false;
    window.mApplyBootstrap_ = function (res) {
      if (res && res.jornadasPorSede && typeof res.jornadasPorSede === 'object') {
        _jornadasPorSede = res.jornadasPorSede;
      }
      return orig.apply(this, arguments);
    };
    return true;
  }

  /* ──────────────────────────────────────────────────────────────────────
   * Inicialización con reintentos
   * ────────────────────────────────────────────────────────────────────*/
  function init() {
    if (_initDone) return;

    var bootstrapOk = hookBootstrap();
    var cargarOk    = hookMCargar();
    hookSedeChange();
    ensureJornadaSelect();

    if (bootstrapOk && cargarOk) {
      _initDone = true;
    }
  }

  /* Arrancar cuando el DOM esté listo */
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { setTimeout(init, 100); });
  } else {
    setTimeout(init, 100);
  }
  setTimeout(init, 500);
  setTimeout(init, 1500);
  setTimeout(init, 3000);

})();
