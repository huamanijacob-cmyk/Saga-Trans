/* =====================================================================
   Navegación enlazada (js/nav.js)
   Recuerda por dónde pasaste (módulo + pestaña + filtros) y muestra
   "← Regresar a …" en los tres módulos. Se carga al final, después de
   app.js, ventas.js y nc.js.
   Un "paso" es cualquier clic que cambia de pantalla: menú lateral,
   pestañas, cifras con enlace, número de documento o factura.
   No cuentan: filtros, buscadores, páginas, cambio de mes o corte.
   ===================================================================== */
(function () {
  'use strict';
  const MAX = 30;
  const pila = [];
  let restaurando = false;

  // Qué clics cambian de pantalla
  const NAVEGA = [
    '.rail-item[id^="rail"]',
    '#modRechazos .tabbar .tab-btn', '#modVentas .v-tabbar .tab-btn', '#ncTabs .tab-btn',
    '#modRechazos td.clickable', '#modVentas td.clickable',
    '#modNC td[data-det]', '[data-nav]',
  ].join(',');

  // Controles de mes/corte: no se restauran al regresar
  const NO_RESTAURAR = new Set(['mesSel', 'diaSel', 'vMesSel', 'vDiaSel', 'ncMesSel', 'ncDiaSel', 'vMeta', 'subirInput', 'ncAlexFile']);

  function moduloActual() {
    for (const m of ['Rechazos', 'Ventas', 'NC']) {
      const el = document.getElementById('mod' + m);
      if (el && el.style.display !== 'none') return m;
    }
    return 'Rechazos';
  }
  const NOMBRE = { Rechazos: 'Rechazos', Ventas: 'Ventas', NC: 'Notas de crédito' };

  // --- Rechazos (su estado vive en app.js: state, PDE_*) ---
  const Rech = {
    snapshot() {
      const campos = {};
      document.querySelectorAll('#modRechazos input, #modRechazos select').forEach(el => {
        if (!el.id || NO_RESTAURAR.has(el.id) || el.type === 'file') return;
        campos[el.id] = el.type === 'checkbox' ? el.checked : el.value;
      });
      return {
        tab: state.tab, cond: { ...state.cond }, vend: { ...state.vend }, doc: { ...state.doc },
        pde: { s: PDE_SEARCH, d: PDE_DESDE, h: PDE_HASTA, p: PDE_PAGE }, campos,
      };
    },
    async restore(s) {
      activarModulo('Rechazos');
      Object.assign(state.cond, s.cond); Object.assign(state.vend, s.vend); state.doc = { ...s.doc };
      PDE_SEARCH = s.pde.s; PDE_DESDE = s.pde.d; PDE_HASTA = s.pde.h; PDE_PAGE = s.pde.p;
      Object.entries(s.campos || {}).forEach(([id, v]) => { const el = document.getElementById(id); if (!el) return; if (el.type === 'checkbox') el.checked = v; else el.value = v; });
      const lbl = (id, txt) => { const e = document.getElementById(id); if (e) e.textContent = txt; };
      lbl('condAlertLabel', state.cond.onlyAlerts ? 'Ver todos' : `Solo alertas ≥${state.cond.umbral}%`);
      lbl('vendAlertLabel', state.vend.onlyAlerts ? 'Ver todos' : `Solo alertas ≥${state.vend.umbral}%`);
      setTab(s.tab);
      if (s.tab === 'pdes' && typeof renderPdes === 'function') renderPdes();
    },
    etiqueta() { const b = document.querySelector('#modRechazos .tabbar .tab-btn.tab-active'); return b ? b.textContent.trim() : ''; },
  };
  const ADAPT = {
    Rechazos: () => Rech,
    Ventas: () => (window.VENTAS && VENTAS.nav) || null,
    NC: () => (window.NCMOD && NCMOD.nav) || null,
  };

  function foto() {
    const mod = moduloActual(); const a = ADAPT[mod]();
    if (!a) return null;
    const et = a.etiqueta();
    return { mod, data: a.snapshot(), label: `${NOMBRE[mod]}${et ? ' · ' + et : ''}` };
  }
  const clave = f => (f ? `${f.mod}|${f.label}` : '');

  function pintar() {
    const ult = pila[pila.length - 1];
    document.querySelectorAll('[data-nav-back]').forEach(b => {
      b.hidden = !ult;
      if (ult) { b.querySelector('span').textContent = `Regresar a ${ult.label}`; b.title = `Regresar a ${ult.label}`; }
    });
  }

  // Antes de cada clic que navega, se guarda la pantalla actual.
  // Si luego no cambió nada (p. ej. clic en la pestaña que ya estaba abierta), se descarta.
  document.addEventListener('click', e => {
    if (restaurando) return;
    const el = e.target.closest(NAVEGA);
    if (!el || e.target.closest('[data-nav-back]')) return;
    const f = foto(); if (!f) return;
    pila.push(f); if (pila.length > MAX) pila.shift();
    const revisar = () => { if (pila[pila.length - 1] === f && clave(foto()) === clave(f)) { pila.pop(); } pintar(); };
    setTimeout(revisar, 60); setTimeout(revisar, 1500);
    pintar();
  }, true);

  async function regresar() {
    const f = pila.pop(); pintar();
    if (!f) return;
    const a = ADAPT[f.mod](); if (!a) return;
    restaurando = true;
    try { await a.restore(f.data); } catch (err) { console.error('No se pudo regresar:', err); }
    finally { setTimeout(() => { restaurando = false; pintar(); }, 50); }
  }
  document.addEventListener('click', e => { if (e.target.closest('[data-nav-back]')) { e.preventDefault(); regresar(); } });

  // El recorrido se borra al cerrar sesión
  if (typeof supabaseClient !== 'undefined') supabaseClient.auth.onAuthStateChange(ev => { if (ev === 'SIGNED_OUT') { pila.length = 0; pintar(); } });
  window.NAV = { get pila() { return pila; }, regresar };
  pintar();
})();
