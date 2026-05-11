/**
 * MHS Integradora — unlock-dates.js  V1
 * ──────────────────────────────────────────────────────────────────────────
 * Módulo "Desbloqueo de Fechas" — flujo de 3 pasos:
 *
 *   Paso 1 →  Selecciona la fecha a desbloquear
 *   Paso 2 →  Confirma sede y turno (auto-rellenados)
 *   Paso 3 →  Solicitud enviada → notificación a SOPORTE/SUPERADMIN
 *             → SOPORTE libera la fecha
 *             → Supervisor recibe confirmación
 *
 * Agrega el botón al home y la pantalla modal completamente independiente.
 * ──────────────────────────────────────────────────────────────────────────
 */
(function () {
  'use strict';

  /* ────────────────────────────────────────────────────────────────────
   * ESTILOS del módulo
   * ──────────────────────────────────────────────────────────────────*/
  var CSS = `
  /* ── Módulo Desbloqueo ── */
  #screenDesbloqueo {
    display: none;
    position: fixed;
    inset: 0;
    z-index: 9000;
    background: var(--bg, #030712);
    flex-direction: column;
    overflow: hidden;
  }
  #screenDesbloqueo.active { display: flex; }

  .desb-topbar {
    flex-shrink: 0;
    display: flex;
    align-items: center;
    gap: 10px;
    padding: calc(env(safe-area-inset-top, 0px) + 12px) 16px 12px;
    background: rgba(3,7,18,.94);
    border-bottom: 1px solid rgba(255,255,255,.08);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
  }
  .desb-back {
    width: 36px; height: 36px; border-radius: 10px;
    background: rgba(255,255,255,.07); border: 1px solid rgba(255,255,255,.12);
    color: #fff; font-size: 18px; cursor: pointer;
    display: flex; align-items: center; justify-content: center;
    -webkit-tap-highlight-color: transparent;
    flex-shrink: 0;
  }
  .desb-topbar-title { font-size: 15px; font-weight: 800; flex: 1; }
  .desb-topbar-sub   { font-size: 10px; color: rgba(147,175,206,.8); margin-top: 1px; }

  .desb-body {
    flex: 1;
    overflow-y: auto;
    -webkit-overflow-scrolling: touch;
    padding: 16px;
    display: flex;
    flex-direction: column;
    gap: 14px;
  }

  /* Stepper */
  .desb-stepper {
    display: flex;
    align-items: center;
    gap: 0;
    margin-bottom: 6px;
  }
  .desb-step-dot {
    width: 28px; height: 28px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 11px; font-weight: 800;
    background: rgba(255,255,255,.06);
    border: 2px solid rgba(255,255,255,.15);
    color: rgba(147,175,206,.7);
    transition: all .25s; flex-shrink: 0;
  }
  .desb-step-dot.active {
    background: #3b82f6;
    border-color: #3b82f6;
    color: #fff;
  }
  .desb-step-dot.done {
    background: rgba(16,185,129,.15);
    border-color: #10b981;
    color: #6ee7b7;
  }
  .desb-step-line {
    flex: 1; height: 2px;
    background: rgba(255,255,255,.1);
    margin: 0 4px;
    transition: background .3s;
  }
  .desb-step-line.done { background: #10b981; }
  .desb-step-labels {
    display: flex; justify-content: space-between;
    font-size: 9px; color: rgba(147,175,206,.6);
    font-weight: 700; text-transform: uppercase; letter-spacing: .35px;
    margin-bottom: 16px;
    padding: 0 2px;
  }

  /* Card contenedor de cada paso */
  .desb-card {
    background: rgba(11,21,38,.95);
    border: 1px solid rgba(59,130,246,.2);
    border-radius: 18px;
    padding: 18px;
    display: flex; flex-direction: column; gap: 14px;
  }
  .desb-card-title {
    font-size: 15px; font-weight: 800; letter-spacing: -.2px;
  }
  .desb-card-sub {
    font-size: 12px; color: rgba(147,175,206,.75);
    line-height: 1.5; margin-top: -8px;
  }
  .desb-field-label {
    font-size: 11px; font-weight: 700;
    color: rgba(147,175,206,.65);
    text-transform: uppercase; letter-spacing: .4px;
    margin-bottom: 6px;
  }
  .desb-field input,
  .desb-field select {
    width: 100%;
    background: rgba(255,255,255,.05);
    border: 1px solid rgba(255,255,255,.12);
    border-radius: 12px;
    color: #f0f6ff;
    font-family: inherit;
    font-size: 15px;
    padding: 13px 14px;
    outline: none;
    -webkit-appearance: none;
    transition: border-color .15s;
  }
  .desb-field input:focus,
  .desb-field select:focus {
    border-color: #3b82f6;
  }
  .desb-field select {
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2393AFCE'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: right 14px center;
    padding-right: 36px;
  }

  /* Info box (sede/turno confirmados) */
  .desb-info-box {
    background: rgba(59,130,246,.08);
    border: 1px solid rgba(59,130,246,.25);
    border-radius: 12px;
    padding: 13px 14px;
    display: flex; flex-direction: column; gap: 6px;
    font-size: 12px;
  }
  .desb-info-row { display: flex; gap: 6px; align-items: center; }
  .desb-info-label { color: rgba(147,175,206,.65); min-width: 60px; }
  .desb-info-val { font-weight: 700; color: #f0f6ff; }

  /* Botón primario del módulo */
  .desb-btn {
    width: 100%;
    padding: 15px;
    border-radius: 14px;
    border: none;
    font-family: inherit;
    font-size: 15px;
    font-weight: 800;
    cursor: pointer;
    -webkit-tap-highlight-color: transparent;
    transition: transform .12s, filter .12s;
  }
  .desb-btn:active { transform: scale(.97); filter: brightness(1.1); }
  .desb-btn.primary {
    background: #3b82f6;
    color: #fff;
  }
  .desb-btn.success {
    background: rgba(16,185,129,.15);
    border: 1px solid rgba(16,185,129,.45);
    color: #6ee7b7;
  }
  .desb-btn.ghost {
    background: rgba(255,255,255,.06);
    border: 1px solid rgba(255,255,255,.12);
    color: rgba(147,175,206,.8);
  }
  .desb-btn:disabled { opacity: .45; cursor: not-allowed; }

  /* Paso 3 — confirmación de envío */
  .desb-sent-icon {
    font-size: 48px; text-align: center; margin: 4px 0;
  }
  .desb-sent-title {
    font-size: 18px; font-weight: 800; text-align: center;
  }
  .desb-sent-sub {
    font-size: 12px; color: rgba(147,175,206,.75);
    text-align: center; line-height: 1.6;
  }
  .desb-timeline {
    display: flex; flex-direction: column; gap: 0;
    border-left: 2px solid rgba(59,130,246,.3);
    margin-left: 12px;
    padding-left: 14px;
  }
  .desb-tl-item {
    position: relative;
    padding: 8px 0;
    font-size: 12px;
    line-height: 1.45;
  }
  .desb-tl-item::before {
    content: '';
    position: absolute;
    left: -19px; top: 11px;
    width: 8px; height: 8px;
    border-radius: 50%;
    background: rgba(59,130,246,.6);
    border: 2px solid rgba(3,7,18,1);
  }
  .desb-tl-item.done::before { background: #10b981; }
  .desb-tl-title { font-weight: 700; color: #f0f6ff; }
  .desb-tl-sub   { color: rgba(147,175,206,.7); margin-top: 2px; }

  /* Error inline */
  .desb-error {
    font-size: 11px; color: #fca5a5;
    background: rgba(239,68,68,.1);
    border: 1px solid rgba(239,68,68,.3);
    border-radius: 9px;
    padding: 8px 12px;
    display: none;
  }
  .desb-error.show { display: block; }

  /* Botón en el home */
  #btnModDesbloqueo {
    display: none;
  }
  `;

  /* ────────────────────────────────────────────────────────────────────
   * HTML del módulo
   * ──────────────────────────────────────────────────────────────────*/
  var HTML = `
  <div id="screenDesbloqueo" role="dialog" aria-label="Solicitar desbloqueo de fecha">

    <!-- Topbar -->
    <div class="desb-topbar">
      <button class="desb-back" onclick="deskClose()" aria-label="Volver">←</button>
      <div>
        <div class="desb-topbar-title">🔓 Desbloqueo de fecha</div>
        <div class="desb-topbar-sub" id="deskTopbarSub">Paso 1 de 3</div>
      </div>
    </div>

    <!-- Cuerpo -->
    <div class="desb-body" id="deskBody">

      <!-- Stepper visual -->
      <div>
        <div class="desb-stepper">
          <div class="desb-step-dot active" id="deskDot1">1</div>
          <div class="desb-step-line" id="deskLine1"></div>
          <div class="desb-step-dot" id="deskDot2">2</div>
          <div class="desb-step-line" id="deskLine2"></div>
          <div class="desb-step-dot" id="deskDot3">3</div>
        </div>
        <div class="desb-step-labels">
          <span>Fecha</span><span>Confirmar</span><span>Enviado</span>
        </div>
      </div>

      <!-- PASO 1: Fecha -->
      <div id="deskStep1">
        <div class="desb-card">
          <div class="desb-card-title">¿Qué fecha necesitas desbloquear?</div>
          <div class="desb-card-sub">Selecciona el día cuya ventana de captura ya venció y necesitas reabrir.</div>
          <div class="desb-field">
            <div class="desb-field-label">Fecha a desbloquear</div>
            <input type="date" id="deskFecha" max="" />
          </div>
          <div class="desb-field">
            <div class="desb-field-label">Motivo breve</div>
            <input type="text" id="deskMotivo" placeholder="Ej: Se cayó el sistema, empleados sin marcar…" maxlength="200" />
          </div>
          <div class="desb-error" id="deskErr1"></div>
          <button class="desb-btn primary" onclick="deskStep1Next()">Siguiente →</button>
        </div>
      </div>

      <!-- PASO 2: Confirmar sede y turno -->
      <div id="deskStep2" style="display:none">
        <div class="desb-card">
          <div class="desb-card-title">Confirma tu sede y turno</div>
          <div class="desb-card-sub">Se enviará la solicitud con estos datos. Edita si es necesario.</div>

          <div class="desb-info-box" id="deskInfoBox">
            <div class="desb-info-row">
              <span class="desb-info-label">Sede</span>
              <span class="desb-info-val" id="deskInfoSede">—</span>
            </div>
            <div class="desb-info-row">
              <span class="desb-info-label">Jornada</span>
              <span class="desb-info-val" id="deskInfoJornada">—</span>
            </div>
            <div class="desb-info-row">
              <span class="desb-info-label">Fecha</span>
              <span class="desb-info-val" id="deskInfoFecha">—</span>
            </div>
          </div>

          <div class="desb-field">
            <div class="desb-field-label">Sede (editar si es distinta)</div>
            <input type="text" id="deskSede" placeholder="Nombre de la sede" />
          </div>
          <div class="desb-field">
            <div class="desb-field-label">Turno</div>
            <select id="deskJornada">
              <option value="MATUTINO">Matutino</option>
              <option value="VESPERTINO">Vespertino</option>
              <option value="NOCTURNO">Nocturno</option>
            </select>
          </div>

          <div class="desb-error" id="deskErr2"></div>

          <div style="display:flex;gap:10px">
            <button class="desb-btn ghost" onclick="deskGoPaso(1)" style="flex:1">← Atrás</button>
            <button class="desb-btn primary" onclick="deskStep2Next()" style="flex:2">Enviar solicitud</button>
          </div>
        </div>
      </div>

      <!-- PASO 3: Confirmación enviada -->
      <div id="deskStep3" style="display:none">
        <div class="desb-card">
          <div class="desb-sent-icon">📨</div>
          <div class="desb-sent-title">Solicitud enviada</div>
          <div class="desb-sent-sub">Tu solicitud fue recibida por el equipo de soporte con prioridad <b style="color:#fcd34d">ALTA</b>.</div>

          <div class="desb-timeline">
            <div class="desb-tl-item done">
              <div class="desb-tl-title">✅ Solicitud registrada</div>
              <div class="desb-tl-sub" id="deskTlFecha">—</div>
            </div>
            <div class="desb-tl-item">
              <div class="desb-tl-title">⏳ Revisión por soporte</div>
              <div class="desb-tl-sub">SOPORTE o SUPERADMIN revisará y aprobará.</div>
            </div>
            <div class="desb-tl-item">
              <div class="desb-tl-title">🔓 Fecha liberada</div>
              <div class="desb-tl-sub">Recibirás una notificación cuando esté lista.</div>
            </div>
          </div>

          <button class="desb-btn success" onclick="deskClose()">✓ Entendido, volver al inicio</button>
          <button class="desb-btn ghost" onclick="deskNuevaSolicitud()" style="margin-top:-4px">+ Nueva solicitud</button>
        </div>
      </div>

    </div><!-- /desb-body -->
  </div><!-- /screenDesbloqueo -->
  `;

  /* ────────────────────────────────────────────────────────────────────
   * INYECCIÓN en el DOM
   * ──────────────────────────────────────────────────────────────────*/
  function inject() {
    if (document.getElementById('screenDesbloqueo')) return; // ya existe

    /* Estilos */
    var style = document.createElement('style');
    style.id = 'desbloqueo-module-css';
    style.textContent = CSS;
    document.head.appendChild(style);

    /* HTML de la pantalla */
    var wrapper = document.createElement('div');
    wrapper.innerHTML = HTML;
    while (wrapper.firstChild) document.body.appendChild(wrapper.firstChild);

    /* Botón en home — insertarlo en la grilla de módulos */
    injectHomeButton();
  }

  function injectHomeButton() {
    /* Esperamos a que exista el grid de módulos */
    var grid = document.querySelector('.home-modules-grid, .modules-grid, #homeModulesGrid');
    if (!grid) {
      setTimeout(injectHomeButton, 400);
      return;
    }
    if (document.getElementById('btnModDesbloqueo')) return;

    var btn = document.createElement('button');
    btn.id        = 'btnModDesbloqueo';
    btn.className = 'home-module-btn';
    btn.setAttribute('onclick', "deskOpen()");
    btn.innerHTML =
      '<div class="hmb-icon" style="font-size:24px;line-height:1;margin-bottom:4px">🔓</div>' +
      '<div class="hmb-title">Desbloqueo de fechas</div>' +
      '<div class="hmb-sub" style="font-size:10px;color:rgba(147,175,206,.65);line-height:1.4;margin-top:4px">' +
        'Solicita apertura de fechas vencidas a soporte' +
      '</div>';
    btn.style.display = 'none'; // Se muestra solo para SUPERVISOR / ADMIN
    grid.appendChild(btn);
  }

  /* ────────────────────────────────────────────────────────────────────
   * Estado interno
   * ──────────────────────────────────────────────────────────────────*/
  var _state = {
    fecha:     '',
    motivo:    '',
    sede:      '',
    jornada:   '',
    chatId:    '',
    sentAt:    ''
  };

  /* ────────────────────────────────────────────────────────────────────
   * FUNCIONES PÚBLICAS (se cuelgan en window)
   * ──────────────────────────────────────────────────────────────────*/

  /** Abrir el módulo desde el home */
  window.deskOpen = function () {
    var screen = document.getElementById('screenDesbloqueo');
    if (!screen) return;

    /* Prellenar valores desde la sesión actual */
    _state.sede    = (typeof window.mSede      === 'string' && window.mSede)      || '';
    _state.jornada = (typeof window.mJornadaAsignada === 'string' && window.mJornadaAsignada) || 'MATUTINO';

    var sedeInput = document.getElementById('deskSede');
    var jornadaSel = document.getElementById('deskJornada');
    var fechaInput = document.getElementById('deskFecha');

    if (sedeInput)  sedeInput.value   = _state.sede;
    if (jornadaSel) jornadaSel.value  = _state.jornada;

    /* Fecha máxima = hoy, mínimo = hace 30 días */
    var hoy = (typeof todayISO === 'function') ? todayISO() : new Date().toISOString().split('T')[0];
    if (fechaInput) {
      fechaInput.max   = hoy;
      fechaInput.value = '';
    }

    /* Reset errores */
    ['deskErr1', 'deskErr2'].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.className = 'desb-error';
    });

    document.getElementById('deskTopbarSub').textContent = 'Paso 1 de 3';
    deskGoPaso(1);

    screen.classList.add('active');
    try { navigator.vibrate && navigator.vibrate([30]); } catch(e) {}
  };

  /** Cerrar el módulo */
  window.deskClose = function () {
    var screen = document.getElementById('screenDesbloqueo');
    if (screen) screen.classList.remove('active');
  };

  /** Navegar a un paso concreto */
  window.deskGoPaso = function (n) {
    [1, 2, 3].forEach(function (i) {
      var step = document.getElementById('deskStep' + i);
      if (step) step.style.display = i === n ? '' : 'none';

      var dot = document.getElementById('deskDot' + i);
      if (dot) {
        dot.className = 'desb-step-dot' +
          (i < n ? ' done' : i === n ? ' active' : '');
        dot.textContent = i < n ? '✓' : String(i);
      }

      if (i < 3) {
        var line = document.getElementById('deskLine' + i);
        if (line) line.className = 'desb-step-line' + (i < n ? ' done' : '');
      }
    });

    var sub = document.getElementById('deskTopbarSub');
    if (sub) sub.textContent = 'Paso ' + n + ' de 3';
  };

  /** Validar y avanzar del Paso 1 → Paso 2 */
  window.deskStep1Next = function () {
    var fechaEl   = document.getElementById('deskFecha');
    var motivoEl  = document.getElementById('deskMotivo');
    var errEl     = document.getElementById('deskErr1');

    var fecha  = (fechaEl  && fechaEl.value)  ? fechaEl.value.trim()  : '';
    var motivo = (motivoEl && motivoEl.value) ? motivoEl.value.trim() : '';

    if (!fecha) {
      if (errEl) { errEl.textContent = 'Selecciona una fecha.'; errEl.className = 'desb-error show'; }
      return;
    }
    if (!motivo) {
      if (errEl) { errEl.textContent = 'Escribe un motivo breve.'; errEl.className = 'desb-error show'; }
      return;
    }
    if (errEl) errEl.className = 'desb-error';

    _state.fecha  = fecha;
    _state.motivo = motivo;

    /* Prellenar paso 2 */
    var sedeInput  = document.getElementById('deskSede');
    var jornadaSel = document.getElementById('deskJornada');
    _state.sede    = (sedeInput  && sedeInput.value)  ? sedeInput.value.trim()  : (_state.sede || '');
    _state.jornada = (jornadaSel && jornadaSel.value) ? jornadaSel.value        : (_state.jornada || 'MATUTINO');

    document.getElementById('deskInfoSede').textContent    = _state.sede    || '—';
    document.getElementById('deskInfoJornada').textContent = _state.jornada || '—';
    document.getElementById('deskInfoFecha').textContent   = _state.fecha;

    if (sedeInput)  sedeInput.value  = _state.sede;
    if (jornadaSel) jornadaSel.value = _state.jornada;

    deskGoPaso(2);
    try { navigator.vibrate && navigator.vibrate([20]); } catch(e) {}
  };

  /** Validar y enviar solicitud desde Paso 2 */
  window.deskStep2Next = function () {
    var sedeInput  = document.getElementById('deskSede');
    var jornadaSel = document.getElementById('deskJornada');
    var errEl      = document.getElementById('deskErr2');

    var sede    = (sedeInput  && sedeInput.value)  ? sedeInput.value.trim() : '';
    var jornada = (jornadaSel && jornadaSel.value) ? jornadaSel.value       : '';

    if (!sede) {
      if (errEl) { errEl.textContent = 'Escribe el nombre de la sede.'; errEl.className = 'desb-error show'; }
      return;
    }
    if (!jornada) {
      if (errEl) { errEl.textContent = 'Selecciona el turno.'; errEl.className = 'desb-error show'; }
      return;
    }
    if (errEl) errEl.className = 'desb-error';

    _state.sede    = sede;
    _state.jornada = jornada;

    deskEnviarSolicitud();
  };

  /** Enviar la solicitud al backend (via apiEnviarTicket con tipo DESBLOQUEO_FECHA) */
  function deskEnviarSolicitud() {
    var token = (typeof getToken === 'function') ? getToken() : '';
    if (!token) {
      if (typeof showToast === 'function') showToast('Sesión expirada. Vuelve a iniciar sesión.', 'bad');
      return;
    }

    /* Generar chatId único para este ticket */
    _state.chatId = 'desk_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 7);

    var mensaje = '🔓 Solicitud de desbloqueo de fecha\n' +
                  'Fecha: ' + _state.fecha + '\n' +
                  'Sede: '  + _state.sede  + '\n' +
                  'Turno: ' + _state.jornada + '\n' +
                  'Motivo: ' + _state.motivo;

    /* Deshabilitar botón mientras se envía */
    var sendBtn = document.querySelector('#deskStep2 .desb-btn.primary');
    if (sendBtn) { sendBtn.disabled = true; sendBtn.textContent = 'Enviando…'; }

    google.script.run
      .withSuccessHandler(function (res) {
        if (sendBtn) { sendBtn.disabled = false; sendBtn.textContent = 'Enviar solicitud'; }

        if (!res || !res.ok) {
          var errEl = document.getElementById('deskErr2');
          if (errEl) {
            errEl.textContent = (res && res.error) ? res.error : 'No se pudo enviar. Intenta de nuevo.';
            errEl.className = 'desb-error show';
          }
          return;
        }

        _state.sentAt = new Date().toLocaleString('es-MX');
        _state.chatId = res.chatId || _state.chatId;

        /* Actualizar timeline */
        var tlFecha = document.getElementById('deskTlFecha');
        if (tlFecha) {
          tlFecha.textContent = _state.fecha + ' · ' + _state.sede + ' · ' + _state.jornada;
        }

        deskGoPaso(3);
        try { navigator.vibrate && navigator.vibrate([60, 40, 60]); } catch (e) {}
        if (typeof showToast === 'function') showToast('✅ Solicitud de desbloqueo enviada', 'ok');
      })
      .withFailureHandler(function (err) {
        if (sendBtn) { sendBtn.disabled = false; sendBtn.textContent = 'Enviar solicitud'; }
        var errEl = document.getElementById('deskErr2');
        if (errEl) {
          errEl.textContent = (err && err.message) ? err.message : 'Error de red. Verifica tu conexión.';
          errEl.className = 'desb-error show';
        }
      })
      .apiEnviarTicket({
        token:             token,
        mensaje:           mensaje,
        sede:              _state.sede,
        jornada:           _state.jornada,
        chatId:            _state.chatId,
        tipo_ticket:       'DESBLOQUEO_FECHA',
        fecha_solicitud:   _state.fecha,
        motivo:            _state.motivo,
        prioridad_override:'ALTA',
        device_id:         (typeof soporteV3DeviceId === 'function') ? soporteV3DeviceId() : ''
      });
  }

  /** Iniciar una nueva solicitud */
  window.deskNuevaSolicitud = function () {
    _state = { fecha: '', motivo: '', sede: window.mSede || '', jornada: window.mJornadaAsignada || 'MATUTINO', chatId: '', sentAt: '' };
    var fechaInput = document.getElementById('deskFecha');
    if (fechaInput) fechaInput.value = '';
    var motivoEl = document.getElementById('deskMotivo');
    if (motivoEl) motivoEl.value = '';
    deskGoPaso(1);
  };

  /* ────────────────────────────────────────────────────────────────────
   * Controlar visibilidad del botón en el home según el rol del usuario
   * ──────────────────────────────────────────────────────────────────*/
  function updateDeskBtnVisibility() {
    var btn = document.getElementById('btnModDesbloqueo');
    if (!btn) return;
    var role = (typeof getRole === 'function') ? getRole() : '';
    /* El módulo es útil para supervisores y admins (no para CEO ni soporte) */
    var show = ['SUPERVISOR', 'ADMIN', 'SUPERADMIN'].indexOf(role.toUpperCase()) >= 0;
    btn.style.display = show ? '' : 'none';
  }

  /* ────────────────────────────────────────────────────────────────────
   * Integración con updateUserPill: cuando cambia la sesión, re-evaluar
   * ──────────────────────────────────────────────────────────────────*/
  function hookUpdateUserPill() {
    var orig = window.updateUserPill;
    if (typeof orig !== 'function') return false;
    window.updateUserPill = function () {
      var r = orig.apply(this, arguments);
      try { updateDeskBtnVisibility(); } catch (e) {}
      return r;
    };
    return true;
  }

  /* ────────────────────────────────────────────────────────────────────
   * Registrar módulo en homeModulo si existe
   * ──────────────────────────────────────────────────────────────────*/
  function hookHomeModulo() {
    var orig = window.homeModulo;
    window.homeModulo = function (mod) {
      if (mod === 'desbloqueo') {
        window.deskOpen();
        return;
      }
      if (typeof orig === 'function') return orig.apply(this, arguments);
    };
  }

  /* ────────────────────────────────────────────────────────────────────
   * BOOT
   * ──────────────────────────────────────────────────────────────────*/
  function init() {
    inject();
    hookHomeModulo();
    hookUpdateUserPill();
    updateDeskBtnVisibility();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { setTimeout(init, 150); });
  } else {
    setTimeout(init, 150);
  }
  setTimeout(init, 600);
  setTimeout(init, 1500);

})();
