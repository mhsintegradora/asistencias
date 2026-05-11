# MHS Integradora · Sistema de Asistencias
### GitHub Pages + PWA (iOS & Android) + GAS como API

---

## 📁 Estructura del repositorio

```
mhsintegradora/asistencias/
├── mobile.html         ← App principal de supervisores (tú ya lo tienes)
├── index.html          ← Vista desktop (tú ya lo tienes)
├── ceomobile.html      ← Dashboard CEO (tú ya lo tienes)
├── router.html         ← Enrutador de acceso (tú ya lo tienes)
├── api.js              ← ⭐ NUEVO — Adaptador GAS→fetch
├── turno-filter.js     ← ⭐ NUEVO — Filtro de jornadas por supervisor
├── unlock-dates.js     ← ⭐ NUEVO — Módulo "Desbloqueo de fechas"
├── manifest.json       ← ⭐ NUEVO — Configuración PWA
├── sw.js               ← ⭐ NUEVO — Service Worker
├── gas-updates.gs      ← ⭐ NUEVO — Código a agregar a code.gs en GAS
└── icons/
    ├── icon-192.png    ← Crea estos íconos con tu logo
    ├── icon-512.png
    └── icon-180.png
```

---

## 🚀 Paso 1 — Configurar GitHub Pages

1. Ve a tu repositorio: `github.com/mhsintegradora/asistencias`
2. **Settings** → **Pages**
3. Source: **Deploy from a branch**
4. Branch: `main` / raíz `/`
5. Guarda. Tu app estará en: `https://mhsintegradora.github.io/asistencias/`

---

## 🔗 Paso 2 — Conectar GAS como API

### 2a. Actualizar code.gs

Abre tu [Google Apps Script](https://script.google.com) y:

1. Copia el contenido de `gas-updates.gs` y pégalo al final de tu `code.gs`
2. **Agrega la función `doPost()`** (ya está en `gas-updates.gs`)
3. **Crea la hoja "Asignaciones"** en tu Google Sheet (ver sección al final)
4. En `apiMobileSessionBootstrap`, agrega al objeto de retorno:
   ```javascript
   jornadasPorSede: getJornadasPorSede_(ses.username, res.sedesAsignadas || []),
   ```

### 2b. Redesplegar el Web App

1. **Implementar** → **Nueva implementación** → tipo: **Web App**
2. Ejecutar como: **Yo**
3. Acceso: **Cualquier usuario**
4. Copia la URL que termina en `/exec`

### 2c. Actualizar api.js

Abre `api.js` y reemplaza:
```javascript
window.MHS_GAS_URL = 'https://script.google.com/macros/s/TU_DEPLOYMENT_ID/exec';
```
con tu URL real.

---

## 📱 Paso 3 — Agregar 3 líneas a mobile.html

Abre `mobile.html` y haz estos **3 cambios simples**:

### 3a. Eliminar `<base target="_top">`
Busca y elimina esta línea (es exclusiva de GAS):
```html
<base target="_top">   ← ELIMINAR
```

### 3b. Agregar en el `<head>`:
```html
<link rel="manifest" href="manifest.json">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="MHS RH">

<!-- Registrar Service Worker -->
<script>
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js')
      .then(function(r){ console.log('[SW] Registrado:', r.scope); })
      .catch(function(e){ console.warn('[SW] Error:', e); });
  }
</script>
```

### 3c. Agregar ANTES de `</body>` (al final del archivo):
```html
<!-- MHS GitHub Modules — agregar en este orden -->
<script src="api.js"></script>
<script src="turno-filter.js"></script>
<script src="unlock-dates.js"></script>
```

> **Importante:** `api.js` debe cargarse **antes** de cualquier llamada a
> `google.script.run`. Al estar al final del body esto funciona correctamente.

### 3d. Hacer lo mismo en index.html, ceomobile.html y router.html
Mismos cambios: eliminar `<base target="_top">`, agregar manifest y `<script src="api.js">`.

---

## 🌐 Paso 4 — Conectar Wix (mhsintegradora.mx)

1. En tu editor Wix, agrega un elemento **"HTML iFrame"** (Embed Code)
2. Pega en el código de iframe:
```html
<iframe
  src="https://mhsintegradora.github.io/asistencias/mobile.html"
  style="width:100%;height:100vh;border:none"
  allow="geolocation; vibrate; notifications"
></iframe>
```
3. Opcionalmente, apunta `mhsintegradora.mx` directamente a GitHub Pages:
   - En GitHub: Settings → Pages → Custom domain → escribe `mhsintegradora.mx`
   - En tu DNS: agrega un CNAME apuntando a `mhsintegradora.github.io`

---

## 🔔 Paso 5 — Notificaciones Push en iOS y Android

### Android
- Funciona automáticamente cuando el usuario abre la app y acepta notificaciones
- El botón 🔔 Alertas en el inbox activa `Notification.requestPermission()`

### iOS (16.4+)
1. El usuario debe abrir la app en **Safari**
2. Toca el botón compartir → **"Agregar a la pantalla de inicio"**
3. Abre la app desde el ícono en la pantalla de inicio
4. La primera vez que el supervisor abra la sección de Soporte/Tickets,
   el sistema pedirá permiso de notificaciones

### Vibración
- Funciona en Android automáticamente (navigator.vibrate)
- En iOS solo vibra en Safari si el usuario está interactuando activamente

---

## 🎯 Paso 6 — Configurar filtro de turnos por supervisor

### En la hoja Google Sheets, crea una hoja llamada "Asignaciones":

| username | sede     | jornada    |
|----------|----------|------------|
| fer      | OHORAN   | MATUTINO   |
| fer      | OHORAN   | VESPERTINO |
| ivan     | OHORAN   | VESPERTINO |
| ivan     | OHORAN   | NOCTURNO   |
| alex     | OHORAN   | NOCTURNO   |
| fer      | SAN JOSE | MATUTINO   |
| ivan     | SAN JOSE | VESPERTINO |

> El `username` debe coincidir exactamente con el nombre de usuario de login.

### Fallback (mientras no agregas la hoja)
Si no agregas la hoja, el archivo `turno-filter.js` tiene un mapa manual que
puedes editar directamente. Busca `SUPERVISOR_JORNADAS_FALLBACK` en el archivo.

---

## 🔓 Nuevo módulo — Desbloqueo de fechas

### ¿Cómo funciona?

1. **Supervisor** toca el botón "🔓 Desbloqueo de fechas" en el home
2. **Paso 1:** Selecciona la fecha que necesita reabrir + motivo
3. **Paso 2:** Confirma sede y turno
4. Se envía un ticket con prioridad **ALTA** al inbox de SOPORTE/SUPERADMIN
5. **SOPORTE** recibe notificación (sonido + vibración + toast)
6. SOPORTE abre el ticket, toca el botón 🔓 "Liberar fecha" de su módulo
7. El supervisor recibe confirmación (el ticket aparece como RESUELTO)

---

## 📋 Checklist de subida a GitHub

- [ ] `mobile.html` — eliminada `<base target="_top">`, agregados 3 scripts
- [ ] `index.html` — mismo cambio
- [ ] `ceomobile.html` — mismo cambio
- [ ] `router.html` — mismo cambio
- [ ] `api.js` — URL de GAS correcta en `window.MHS_GAS_URL`
- [ ] `manifest.json` — `start_url` correcto (`/asistencias/mobile.html`)
- [ ] `sw.js` — subido
- [ ] `turno-filter.js` — fallback configurado con Fer/Ivan/Alex
- [ ] `unlock-dates.js` — subido
- [ ] `icons/` — íconos de la app creados (192×192, 512×512, 180×180)
- [ ] GAS — `doPost()` agregado a code.gs y re-desplegado
- [ ] GAS — hoja "Asignaciones" creada con los turnos por supervisor
- [ ] GitHub Pages — activado en Settings → Pages

---

## 🆘 Solución de problemas

**"Error de CORS al llamar al GAS"**
→ Asegúrate de que el Web App esté configurado con acceso **"Cualquier usuario"**
→ Verifica que la URL en `api.js` termina en `/exec` (no `/dev`)

**"La app no se instala en iOS"**
→ Debe abrirse en Safari (no Chrome ni Firefox en iOS)
→ El usuario debe tocar: Compartir → Agregar a pantalla de inicio

**"Las notificaciones no llegan"**
→ En Android: verifica que el usuario otorgó permiso de notificaciones
→ En iOS: solo funciona desde la PWA instalada (no desde Safari directo)

**"Los turnos no se filtran"**
→ Verifica que el username en la hoja "Asignaciones" sea idéntico al de login
→ Como respaldo, edita `SUPERVISOR_JORNADAS_FALLBACK` en `turno-filter.js`

---

*MHS Integradora · Sistema de RRHH · mhsintegradora.mx*
