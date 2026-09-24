/* =====================================================================
   Módulo Ventas — Saga Trans Confitería (js/ventas.js)
   Se carga DESPUÉS de js/app.js y reutiliza de ahí:
     supabaseClient · loadModuleFile · FILE_CACHE · downloadStyledXlsx
   Todo lo propio vive dentro de esta función: no crea variables globales
   salvo window.VENTAS (para diagnóstico desde la consola, F12).
   ===================================================================== */
(function () {
  'use strict';

  // ---------------- Configuración del módulo ----------------
  const VCFG = {
    VENTAS_FOLDER: 'Ventas', VENTAS_PREFIX: 'Ventas',   // se lee con loadModuleFile (bucket de Rechazos)
    DATA_BUCKET: 'ventas-data',
    CUOTAS: m => `Cuotas/Cuotas_${m}.xlsx`,
    CARTERA: m => `Cartera/Cartera_${m}.xlsx`,
    CALENDARIO: y => `Calendario/Calendario_${y}.xlsx`,
    META: m => `Overrides/meta_${m}.json`,
    CATEGORIAS: [
      { codigo: '00001', nombre: 'Chocolate', columnaCuota: 'Cuota Chocolate' },
      { codigo: '00002', nombre: 'Galleta', columnaCuota: 'Cuota Galleta' },
      { codigo: '00003', nombre: 'Paneton', columnaCuota: 'Cuota Paneton' },
    ],
    CATEGORIAS_AVANCE_DEFAULT: ['00001', '00002'],
    VENDEDORES_SIN_RANKING: ['99999'],
    UMBRAL_ATENCION: 0.85,
    UMBRAL_EN_RITMO: 1.00,
    DOCS_POR_PAGINA: 15,
  };

  const C = window.VentasCalc;
  const $ = id => document.getElementById(id);
  const CAT = Object.fromEntries(VCFG.CATEGORIAS.map(c => [c.codigo, c]));
  const ALL_CATS = VCFG.CATEGORIAS.map(c => c.codigo);
  const MES_ABBR = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Set', 'Oct', 'Nov', 'Dic'];
  const MES_NOMBRE = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'setiembre', 'octubre', 'noviembre', 'diciembre'];
  const EST_LABEL = { riesgo: 'En riesgo', atencion: 'Atención', ritmo: 'En ritmo' };

  const vs = {
    booted: false, modo: 'mes', mesSel: null, diaSel: null, ym: null, corte: null, tab: 'avance',
    cats: new Set(VCFG.CATEGORIAS_AVANCE_DEFAULT), catCob: ALL_CATS[0],
    meta: 1, metaGuardada: 1, metaRegistro: null,
    doc: { buscar: '', vend: 'all', cat: 'all', cats: null, desde: '', hasta: '', page: 0 },
    cmp: { cats: new Set(VCFG.CATEGORIAS_AVANCE_DEFAULT), vista: 'acum' },   // siempre toda la empresa
  };
  const PREVIO = {};   // ym → ventas parseadas del mismo mes del año anterior (o null si no existe)
  let CMP_CHART = null;
  let VD = null;   // datos del mes cargado
  let VR = {};     // últimos resultados calculados
  window.VENTAS = { state: vs, get data() { return VD; }, get res() { return VR; }, config: VCFG };

  // ---------------- Formato ----------------
  const pad = n => String(n).padStart(2, '0');
  const nf0 = new Intl.NumberFormat('es-PE', { maximumFractionDigits: 0 });
  const nf2 = new Intl.NumberFormat('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fS = n => (n === null || n === undefined ? '—' : 'S/ ' + nf0.format(Math.round(n)));
  const fS2 = n => 'S/ ' + nf2.format(n || 0);
  const fN = n => nf0.format(n || 0);
  const fP = n => (n === null || n === undefined || !isFinite(n) ? '—' : (Math.round(n * 1000) / 10).toFixed(1) + '%');
  const fD = f => (f ? `${f.slice(8, 10)}/${f.slice(5, 7)}/${f.slice(0, 4)}` : '');
  const toIso = f => (f ? f.replace(/\//g, '-') : '');
  const fromIso = s => (s ? s.replace(/-/g, '/') : '');
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const title = s => String(s || '').toLowerCase().replace(/(^|\s)\S/g, m => m.toUpperCase());
  const largo = f => { const [y, m, d] = f.split('/').map(Number); return `${d} de ${MES_NOMBRE[m - 1]} del ${y}`; };
  const estado = p => C.estado(p, VCFG);
  const barCls = e => ({ ritmo: 'v-bar-ok', atencion: 'v-bar-mid', riesgo: 'v-bar-bad' }[e] || 'v-bar-neutral');
  const teorico = () => (VR.hab && VR.hab.total ? VR.hab.transcurridos / VR.hab.total : 0);
  const fT = t => Math.round(t * 100) + '%';

  // % primero y la barra después. ref = contra qué se evalúa (teórico o 1 = meta).
  // Verde ≥ ref · amarillo ≥ 85% de ref · rojo menos. Con ref < 1 se dibuja la marca del teórico.
  function pctHtml(p, ref) {
    if (p === null || p === undefined || !isFinite(p)) return '<span class="pct-num">—</span>';
    const w = Math.max(0, Math.min(p, 1)) * 100;
    const cls = ref ? barCls(estado(p / ref)) : 'v-bar-neutral';
    const marca = ref && ref < 1 ? `<i class="v-marca" style="left:${ref * 100}%"></i>` : '';
    return `<div class="v-pct"><span class="pct-num">${fP(p)}</span><div class="pct-bar-track v-track"><div class="pct-bar ${cls}" style="width:${w}%"></div>${marca}</div></div>`;
  }
  const estHtml = e => (e ? `<span class="v-est ${e}">${EST_LABEL[e]}</span>` : '');
  function linkTd(txt, data) {
    const attrs = Object.entries(data).map(([k, v]) => `data-${k}="${esc(v)}"`).join(' ');
    return `<td class="num clickable" title="Ver documentos" ${attrs}><span class="v-link">${txt}</span></td>`;
  }
  function cuotaHtml(cuota, q100) {
    if (!(q100 > 0)) return '—';
    return Math.round(vs.meta * 100) !== 100
      ? `<span class="v-ajustada">${fS(cuota)}</span><div class="sub">100% ${fS(q100)}</div>` : fS(cuota);
  }
  function kpi(color, label, value, sub) {
    return `<div class="kpi kpi-${color}"><div class="kpi-label">${label}</div><div class="kpi-val">${value}</div>${sub ? `<div class="v-kpi-sub">${sub}</div>` : ''}</div>`;
  }
  function aviso(html, err) {
    const d = document.createElement('div');
    d.className = 'pendiente-banner' + (err ? ' v-aviso-err' : '');
    d.innerHTML = html;
    $('vAvisos').appendChild(d);
  }

  // ---------------- Carga de archivos ----------------
  function overlay(on, txt) {
    const o = $('loadingOverlay'); if (!o) return;
    const t = o.querySelector('.loading-text');
    if (on) { o.dataset.prev = o.dataset.prev || (t ? t.textContent : ''); if (t) t.textContent = txt || 'Cargando ventas...'; o.style.display = ''; }
    else { o.style.display = 'none'; if (t && o.dataset.prev) t.textContent = o.dataset.prev; }
  }
  async function token() { const { data } = await supabaseClient.auth.getSession(); return data?.session?.access_token; }
  // Descarga sin caché del navegador (para que un archivo recién subido se vea al instante).
  async function descargar(bucket, path) {
    const url = `${SUPABASE_URL}/storage/v1/object/${bucket}/${path.split('/').map(encodeURIComponent).join('/')}?cachebust=${Date.now()}`;
    const res = await fetch(url, { cache: 'no-store', headers: { Authorization: `Bearer ${await token()}`, apikey: SUPABASE_ANON_KEY } });
    if (res.status === 400 || res.status === 404) return null; // el archivo no existe
    if (!res.ok) throw new Error(`HTTP ${res.status} al descargar ${bucket}/${path}`);
    return res.arrayBuffer();
  }
  async function leerXlsx(bucket, path) {
    const buf = await descargar(bucket, path);
    if (!buf) return null;
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '', raw: true });
  }
  async function leerJson(bucket, path) {
    const buf = await descargar(bucket, path);
    if (!buf) return null;
    try { return JSON.parse(new TextDecoder().decode(buf)); } catch { return null; }
  }
  // Ventas: se comparte con Rechazos (mismo archivo, mismo caché: no se descarga dos veces).
  async function leerVentas(ym) {
    const rows = await loadModuleFile(VCFG.VENTAS_FOLDER, VCFG.VENTAS_PREFIX, ym);
    return rows && rows.length ? rows : null;
  }

  async function cargarMes(ym, { forzar } = {}) {
    overlay(true, 'Cargando ventas...');
    $('vAvisos').innerHTML = '';
    if (forzar && typeof FILE_CACHE !== 'undefined') delete FILE_CACHE[`${VCFG.VENTAS_FOLDER}/${VCFG.VENTAS_PREFIX}_${ym}.xlsx`];
    if (forzar) {
      const py = `${+ym.slice(0, 4) - 1}${ym.slice(4)}`;
      delete PREVIO[py];
      if (typeof FILE_CACHE !== 'undefined') delete FILE_CACHE[`${VCFG.VENTAS_FOLDER}/${VCFG.VENTAS_PREFIX}_${py}.xlsx`];
    }
    const year = ym.slice(0, 4);
    const intento = async (fn, nombre) => {
      try { return await fn(); } catch (e) { console.error(nombre, e); aviso(`No se pudo leer <b>&nbsp;${esc(nombre)}&nbsp;</b>: ${esc(e.message)}`, true); return null; }
    };
    try {
      const [vRows, qRows, cRows, calRows, metaJson] = await Promise.all([
        intento(() => leerVentas(ym), `Ventas ${ym}`),
        intento(() => leerXlsx(VCFG.DATA_BUCKET, VCFG.CUOTAS(ym)), `Cuotas ${ym}`),
        intento(() => leerXlsx(VCFG.DATA_BUCKET, VCFG.CARTERA(ym)), `Cartera ${ym}`),
        intento(() => leerXlsx(VCFG.DATA_BUCKET, VCFG.CALENDARIO(year)), `Calendario ${year}`),
        intento(() => leerJson(VCFG.DATA_BUCKET, VCFG.META(ym)), `Meta ${ym}`),
      ]);
      const pref = ym.replace('-', '/');
      const todas = vRows ? C.parseVentas(vRows) : [];
      const ventas = todas.filter(r => r.fecha.startsWith(pref));
      if (!vRows) aviso(`Todavía no hay datos de Ventas para ${MES_NOMBRE[+ym.slice(5) - 1]} ${year}. Sube <b>&nbsp;Ventas_${ym}.xlsx&nbsp;</b> a Storage y este panel se completa solo.`);
      else if (!('vtadvo' in vRows[0])) aviso('El archivo de Ventas no tiene la columna <b>&nbsp;vtadvo&nbsp;</b>. ¿Se subió el archivo correcto?', true);
      if (todas.length > ventas.length) aviso(`${fN(todas.length - ventas.length)} filas de Ventas tienen fecha fuera del mes y no se consideran.`);

      let cuotas = {};
      if (qRows) { const r = C.parseCuotas(qRows, VCFG.CATEGORIAS); cuotas = r.cuotas; r.avisos.forEach(a => aviso(esc(a))); }
      else aviso(`No se encontró <b>&nbsp;ventas-data/${esc(VCFG.CUOTAS(ym))}&nbsp;</b>. Las cuotas se muestran en 0.`);

      const m = metaJson && Number(metaJson.meta);
      vs.meta = vs.metaGuardada = m > 0 ? m / 100 : 1;
      vs.metaRegistro = m > 0 ? metaJson : null;
      $('vMeta').value = Math.round(vs.meta * 100);

      VD = {
        ym, ventas, cuotas,
        cartera: cRows ? C.parseCartera(cRows) : null,
        calendario: calRows ? C.parseCalendario(calRows) : {},
        tieneCalendario: !!calRows,
        fechas: C.fechasDisponibles(ventas),
      };
      vs.ym = ym;
    } finally {
      overlay(false);
    }
  }

  // ---------------- Corte (mismo patrón que Rechazos: Cierre de mes / Día específico) ----------------
  function corteDelMes() {
    if (VD.fechas.length) return VD.fechas[VD.fechas.length - 1];
    const [y, m] = VD.ym.split('-').map(Number);
    return `${y}/${pad(m)}/${pad(new Date(y, m, 0).getDate())}`;
  }
  async function aplicarCorte({ forzar } = {}) {
    const ym = vs.modo === 'dia' ? vs.diaSel.slice(0, 7) : vs.mesSel;
    if (!VD || VD.ym !== ym || forzar) await cargarMes(ym, { forzar });
    vs.corte = vs.modo === 'dia' ? fromIso(vs.diaSel) : corteDelMes();
    const desde = vs.corte.slice(0, 8) + '01';
    Object.assign(vs.doc, { desde, hasta: vs.corte, page: 0 });
    renderTodo();
  }

  function initControles() {
    const hoy = new Date();
    vs.mesSel = `${hoy.getFullYear()}-${pad(hoy.getMonth() + 1)}`;
    vs.diaSel = `${vs.mesSel}-${pad(hoy.getDate())}`;
    const sel = $('vMesSel');
    for (let i = 11; i >= 0; i--) {
      const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
      const val = `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
      sel.appendChild(new Option(`${MES_ABBR[d.getMonth()]} ${d.getFullYear()}`, val, false, val === vs.mesSel));
    }
    sel.addEventListener('change', () => { vs.mesSel = sel.value; aplicarCorte(); });
    const dia = $('vDiaSel');
    dia.addEventListener('change', () => { if (dia.value) { vs.diaSel = dia.value; aplicarCorte(); } });
    $('vModoMes').addEventListener('click', () => {
      vs.modo = 'mes'; $('vModoMes').classList.add('tab-active'); $('vModoDia').classList.remove('tab-active');
      sel.style.display = ''; dia.style.display = 'none'; aplicarCorte();
    });
    $('vModoDia').addEventListener('click', () => {
      vs.modo = 'dia'; $('vModoDia').classList.add('tab-active'); $('vModoMes').classList.remove('tab-active');
      if (VD && vs.corte) vs.diaSel = toIso(vs.corte);
      dia.value = vs.diaSel; sel.style.display = 'none'; dia.style.display = ''; aplicarCorte();
    });
  }

  // ---------------- Render ----------------
  function renderTodo() {
    if (!VD) return;
    VR.hab = C.habiles(VD.ym, VD.calendario, vs.corte);
    document.querySelectorAll('.vCorteInline').forEach(e => { e.textContent = largo(vs.corte); });
    $('vCorteLabel').textContent = largo(vs.corte);
    renderHabiles();
    renderCats();
    renderMetaInfo();
    renderAvance();
    renderCobertura();
    renderAlertas();
    renderDocs();
    if (vs.tab === 'comparativo') renderCmp();
  }

  function renderHabiles() {
    const h = VR.hab;
    const ex = h.dias.filter(d => d.excepcion);
    const exTxt = ex.length ? ` · ${ex.length} día${ex.length > 1 ? 's' : ''} del calendario` : '';
    $('vHabiles').innerHTML = `<span class="v-dot"></span>Teórico: <b>${fT(teorico())}</b> · <b>${h.transcurridos}</b> de <b>${h.total}</b> días hábiles · restan <b>${h.restantes}</b>${exTxt}`;
    const [y, m] = VD.ym.split('-').map(Number);
    const first = (new Date(y, m - 1, 1).getDay() + 6) % 7;
    let html = '<div class="v-cal-grid">' + ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'].map(d => `<div class="v-cal-hd">${d}</div>`).join('') + '<div></div>'.repeat(first);
    h.dias.forEach(d => {
      const cls = ['v-cal-d', d.trabaja ? '' : 'off', d.trabaja && d.excepcion ? 'ex-si' : '', d.fecha === vs.corte ? 'corte' : '', d.fecha < vs.corte ? 'past' : ''].join(' ');
      const lbl = !d.trabaja ? (d.motivo || 'No se trabaja') : (d.excepcion ? (d.motivo || 'Se trabaja') : '');
      html += `<div class="${cls}" title="${esc(fD(d.fecha) + (lbl ? ' · ' + lbl : ''))}">${d.dia}<small>${esc(lbl)}</small></div>`;
    });
    html += `</div><div class="v-cal-leg"><span><i></i>Se trabaja</span><span><i style="background:#FDEEEA;border-color:#F2C9BE"></i>No se trabaja</span><span><i style="background:#E6F4F2;border-color:#BFE3DE"></i>Domingo que sí se trabaja</span><span><i style="box-shadow:inset 0 0 0 2px var(--navy)"></i>Fecha de corte</span><span>${VD.tieneCalendario ? `Excepciones de <span class="mono">Calendario_${VD.ym.slice(0, 4)}.xlsx</span>` : 'Sin archivo de calendario: lunes a sábado'}</span></div>`;
    $('vCalendario').innerHTML = html;
  }

  function renderCats() {
    $('vCats').innerHTML = '<span class="v-cats-lbl">Categorías</span>' + VCFG.CATEGORIAS.map(c =>
      `<label class="v-chip ${vs.cats.has(c.codigo) ? 'on' : ''}"><input type="checkbox" value="${c.codigo}" ${vs.cats.has(c.codigo) ? 'checked' : ''}>${c.nombre}</label>`).join('');
    document.querySelectorAll('.vCatsInline').forEach(e => { e.textContent = [...vs.cats].map(c => CAT[c].nombre).join(' + '); });
  }

  function renderMetaInfo(err) {
    const box = $('vMetaInfo');
    box.classList.toggle('err', !!err);
    if (err) { box.textContent = err; $('vMetaGuardar').style.display = 'none'; return; }
    const pct = Math.round(vs.meta * 100);
    const dirty = Math.abs(vs.meta - vs.metaGuardada) > 1e-9;
    $('vMetaGuardar').style.display = dirty ? '' : 'none';
    if (dirty) box.textContent = `Meta sin guardar: ${pct}%. Guárdala para que todos vean las mismas cuotas.`;
    else if (vs.metaRegistro && pct !== 100) box.textContent = `Meta del mes: ${pct}% · guardada por ${vs.metaRegistro.usuario || '—'} el ${vs.metaRegistro.fechaTexto || ''}`;
    else box.textContent = 'Meta al 100%: se usan las cuotas del Excel.';
  }

  function renderAvance() {
    const cats = [...vs.cats];
    const tb = $('vAvanceTable');
    if (!cats.length) {
      $('vAvanceKpis').innerHTML = '';
      tb.querySelector('thead').innerHTML = ''; tb.querySelector('tfoot').innerHTML = '';
      tb.querySelector('tbody').innerHTML = '<tr><td class="muted">Marca al menos una categoría.</td></tr>';
      VR.avance = null; return;
    }
    const ag = C.avanceGeneral({ ventas: VD.ventas, cuotas: VD.cuotas, corte: vs.corte, hab: VR.hab, cats, meta: vs.meta, cfg: VCFG });
    VR.avance = ag;
    const T = ag.total, desde = vs.corte.slice(0, 8) + '01', catNames = cats.map(c => CAT[c].nombre).join(' + ');
    const T0 = teorico();
    $('vAvanceKpis').innerHTML =
      kpi('navy', 'Cuota', fS(T.cuota), Math.round(vs.meta * 100) !== 100 ? `100% ${fS(T.cuota100)} · meta ${Math.round(vs.meta * 100)}%` : catNames) +
      kpi('teal', 'Facturado', fS(T.facturado), 'venta neta sin IGV') +
      kpi('navy', '% Avance', fP(T.pct), `teórico ${fT(T0)}`) +
      kpi(T.estado === 'ritmo' ? 'green' : T.estado === 'atencion' ? 'amber' : 'red', 'Proyección de cierre', fP(T.proy), EST_LABEL[T.estado] || '');
    tb.querySelector('thead').innerHTML = '<tr><th class="c">Código</th><th>Vendedor</th><th class="r">Cuota</th><th class="r">Facturado</th><th class="c">% Avance</th><th class="c">Proyección</th><th class="c">Estado</th><th class="c">Ranking</th></tr>';
    tb.querySelector('tbody').innerHTML = ag.filas.map(f => `<tr class="${f.sinRanking ? 'v-muted' : ''}">
      <td class="mono c">${esc(f.vendedor)}</td><td>${esc(title(f.nombre))}</td>
      <td class="num">${cuotaHtml(f.cuota, f.cuota100)}</td>
      ${linkTd(fS(f.facturado), { v: f.vendedor, cats: cats.join(','), desde, hasta: vs.corte })}
      <td class="num">${pctHtml(f.pct, T0)}</td><td class="num">${pctHtml(f.proy, 1)}</td>
      <td class="c">${estHtml(f.estado)}</td><td class="c v-rk">${f.ranking || ''}</td></tr>`).join('');
    tb.querySelector('tfoot').innerHTML = `<tr><td></td><td>Saga Trans Confitería</td><td class="num">${cuotaHtml(T.cuota, T.cuota100)}</td>
      ${linkTd(fS(T.facturado), { v: '', cats: cats.join(','), desde, hasta: vs.corte })}
      <td class="num">${pctHtml(T.pct, T0)}</td><td class="num">${pctHtml(T.proy, 1)}</td><td class="c">${estHtml(T.estado)}</td><td></td></tr>`;
  }

  function renderCobertura() {
    const cat = CAT[vs.catCob];
    $('vCobCats').innerHTML = VCFG.CATEGORIAS.map(c => `<button class="tab-btn ${c.codigo === vs.catCob ? 'tab-active' : ''}" data-vcat="${c.codigo}">${c.nombre}</button>`).join('');
    $('vCobCatLabel').textContent = cat.nombre;
    document.querySelectorAll('.vCobCatInline').forEach(e => { e.textContent = cat.nombre; });
    $('vCobSinCartera').style.display = VD.cartera ? 'none' : '';
    const cartera = VD.cartera || { todos: {}, activosPorVend: {} };
    const cb = C.cobertura({ ventas: VD.ventas, cuotas: VD.cuotas, cartera, corte: vs.corte, hab: VR.hab, cat: cat.codigo, meta: vs.meta, cfg: VCFG });
    VR.cobertura = cb;
    const M = cb.total.mes, D = cb.total.dia, desde = vs.corte.slice(0, 8) + '01';
    const T0 = teorico();
    $('vCobKpis').innerHTML =
      kpi('navy', 'Clientes en cartera', fN(M.clientes), 'cliente AC y local AC') +
      kpi('teal', `Con venta de ${cat.nombre}`, fN(M.conVenta), `cobertura ${fP(M.pctCob)} · teórico ${fT(T0)}`) +
      kpi('navy', 'Avance del mes', fS(M.avance), `${fP(M.pct)} de ${fS(M.cuota)}`) +
      kpi('red', `Venta del ${fD(vs.corte).slice(0, 5)}`, fS(D.avance), `${fN(D.nuevos)} de ${fN(D.pendientes)} pendientes · ${fP(D.pct)} de la cuota diaria`);
    $('vCobTitMes').textContent = `Resumen del mes · acumulado al ${largo(vs.corte)}`;
    $('vCobTitDia').textContent = `Acumulado del día · ${largo(vs.corte)} · restan ${VR.hab.restantes} días hábiles`;
    const head = (c1, c2, s1, extra) => `<tr class="v-grp"><th></th><th></th><th colspan="${extra ? 3 : 2}" class="v-grp-sep">Cliente</th><th class="v-grp-sep"></th><th colspan="2" class="v-grp-sep">Soles</th><th class="v-grp-sep"></th></tr>
      <tr><th class="c">Código</th><th>Vendedor</th><th class="c">${c1}</th>${extra ? `<th class="c">${extra}</th>` : ''}<th class="c">${c2}</th><th class="c">% Cobertura</th><th class="r">${s1}</th><th class="r">Avance</th><th class="c">% Avance</th></tr>`;
    const tm = $('vCobMesTable'), td = $('vCobDiaTable');
    tm.querySelector('thead').innerHTML = head('Totales', `Venta ${cat.nombre}`, 'Cuota');
    td.querySelector('thead').innerHTML = head('Pendientes', `Venta ${cat.nombre}`, 'Cuota diaria');
    tm.querySelector('tbody').innerHTML = cb.filas.map(f => `<tr><td class="mono c">${esc(f.vendedor)}</td><td>${esc(title(f.nombre))}</td>
      <td class="num c">${fN(f.mes.clientes)}</td><td class="num c">${fN(f.mes.conVenta)}</td><td class="num">${pctHtml(f.mes.pctCob, T0)}</td>
      <td class="num">${cuotaHtml(f.mes.cuota, f.mes.cuota100)}</td>${linkTd(fS(f.mes.avance), { v: f.vendedor, cats: cat.codigo, desde, hasta: vs.corte })}
      <td class="num">${pctHtml(f.mes.pct, T0)}</td></tr>`).join('');
    tm.querySelector('tfoot').innerHTML = `<tr><td></td><td>Total</td><td class="num c">${fN(M.clientes)}</td><td class="num c">${fN(M.conVenta)}</td><td class="num">${pctHtml(M.pctCob, T0)}</td>
      <td class="num">${cuotaHtml(M.cuota, M.cuota100)}</td><td class="num">${fS(M.avance)}</td><td class="num">${pctHtml(M.pct, T0)}</td></tr>`;
    td.querySelector('tbody').innerHTML = cb.filas.map(f => {
      const cd = f.dia.cuotaDiaria === null ? '—' : fS(f.dia.cuotaDiaria);
      return `<tr><td class="mono c">${esc(f.vendedor)}</td><td>${esc(title(f.nombre))}</td>
      <td class="num c">${fN(f.dia.pendientes)}</td><td class="num c">${fN(f.dia.nuevos)}</td><td class="num">${pctHtml(f.dia.pctCob, 1)}</td>
      <td class="num">${cd}</td>${linkTd(fS(f.dia.avance), { v: f.vendedor, cats: cat.codigo, desde: vs.corte, hasta: vs.corte })}
      <td class="num">${pctHtml(f.dia.pct, 1)}</td></tr>`;
    }).join('');
    td.querySelector('tfoot').innerHTML = `<tr><td></td><td>Total</td><td class="num c">${fN(D.pendientes)}</td><td class="num c">${fN(D.nuevos)}</td><td class="num">${pctHtml(D.pctCob, 1)}</td>
      <td class="num">${D.cuotaDiaria === null ? '—' : fS(D.cuotaDiaria)}</td><td class="num">${fS(D.avance)}</td><td class="num">${pctHtml(D.pct, 1)}</td></tr>`;
  }

  function renderAlertas() {
    const tb = $('vAlertasTable'), badge = $('vAlertasCount');
    const setBadge = n => { badge.textContent = n; badge.classList.toggle('zero', !n); };
    tb.querySelector('thead').innerHTML = '<tr><th>Vendedor</th><th class="c">Código cliente</th><th>Cliente</th><th>Motivo</th><th class="c">Docs</th><th class="r">Soles</th><th class="c">Última venta</th></tr>';
    if (!VD.cartera) { tb.querySelector('tbody').innerHTML = '<tr><td colspan="7" class="muted">Sin archivo de Cartera: no se pueden revisar ventas fuera de cartera.</td></tr>'; setBadge(0); return; }
    const al = C.fueraDeCartera({ ventas: VD.ventas, cartera: VD.cartera, corte: vs.corte, catCodes: ALL_CATS, cfg: VCFG });
    VR.alertas = al; setBadge(al.length);
    const desde = vs.corte.slice(0, 8) + '01';
    tb.querySelector('tbody').innerHTML = al.length ? al.map(a => `<tr><td>${esc(a.vendedor)} · ${esc(title(a.nombreVendedor))}</td><td class="mono c">${esc(a.cli)}</td><td>${esc(a.nombre)}</td><td>${esc(a.motivo)}</td>
      <td class="num c">${a.docs}</td>${linkTd(fS(a.soles), { v: a.vendedor, cats: ALL_CATS.join(','), desde, hasta: vs.corte, cli: a.cli })}<td class="num c">${fD(a.ultima)}</td></tr>`).join('')
      : '<tr><td colspan="7" class="muted">Todas las ventas del mes son a clientes de la cartera activa.</td></tr>';
  }

  // ---------------- Documentos (drill-down, mismo patrón que Rechazos) ----------------
  function docCats() { return vs.doc.cats || (vs.doc.cat === 'all' ? ALL_CATS : [vs.doc.cat]); }
  function filtrarDocs() {
    return C.documentos(VD.ventas, { vendedor: vs.doc.vend === 'all' ? null : vs.doc.vend, cats: docCats(), desde: vs.doc.desde, hasta: vs.doc.hasta, buscar: vs.doc.buscar });
  }
  function renderDocs() {
    const vends = {};
    VD.ventas.forEach(r => { vends[r.vendedor] = vends[r.vendedor] || r.nombrevendedor; });
    const vSel = $('vDocVend');
    vSel.innerHTML = '<option value="all">Todos</option>' + Object.keys(vends).sort().map(v => `<option value="${esc(v)}">${esc(v)} · ${esc(title(vends[v]))}</option>`).join('');
    vSel.value = vends[vs.doc.vend] ? vs.doc.vend : 'all';
    vs.doc.vend = vSel.value;
    const cSel = $('vDocCat');
    let opts = '<option value="all">Todas</option>' + VCFG.CATEGORIAS.map(c => `<option value="${c.codigo}">${c.nombre}</option>`).join('');
    if (vs.doc.cats && vs.doc.cats.length > 1 && vs.doc.cats.length < ALL_CATS.length) opts += `<option value="custom">${vs.doc.cats.map(c => CAT[c].nombre).join(' + ')}</option>`;
    cSel.innerHTML = opts;
    cSel.value = vs.doc.cats ? (vs.doc.cats.length === 1 ? vs.doc.cats[0] : (vs.doc.cats.length === ALL_CATS.length ? 'all' : 'custom')) : vs.doc.cat;
    $('vDocSearch').value = vs.doc.buscar;
    $('vDocDesde').value = toIso(vs.doc.desde);
    $('vDocHasta').value = toIso(vs.doc.hasta);
    dibujarDocs();
  }
  function dibujarDocs() {
    const docs = filtrarDocs();
    VR.docs = docs;
    const tot = docs.reduce((a, d) => ({ venta: a.venta + d.venta, devolucion: a.devolucion + d.devolucion, neto: a.neto + d.neto }), { venta: 0, devolucion: 0, neto: 0 });
    VR.docsTot = tot;
    $('vDocKpis').innerHTML = kpi('navy', 'Documentos', fN(docs.length)) + kpi('teal', 'Venta', fS(tot.venta)) + kpi('red', 'Devolución', fS(tot.devolucion)) + kpi('navy', 'Neto', fS(tot.neto));
    const vendTxt = vs.doc.vend === 'all' ? 'todos los vendedores' : `vendedor ${vs.doc.vend}`;
    $('vDocLabel').textContent = `${fN(docs.length)} documentos · ${vendTxt} · ${docCats().map(c => CAT[c].nombre).join(' + ')} · ${fD(vs.doc.desde)} al ${fD(vs.doc.hasta)} · neto ${fS2(tot.neto)}`;
    const per = VCFG.DOCS_POR_PAGINA, maxPage = Math.max(0, Math.ceil(docs.length / per) - 1);
    vs.doc.page = Math.min(vs.doc.page, maxPage);
    const page = docs.slice(vs.doc.page * per, vs.doc.page * per + per);
    $('vDocPage').textContent = `${vs.doc.page + 1} / ${maxPage + 1}`;
    $('vDocPrev').disabled = vs.doc.page === 0; $('vDocNext').disabled = vs.doc.page === maxPage;
    const tb = $('vDocTable');
    tb.querySelector('thead').innerHTML = '<tr><th class="c">Fecha</th><th class="c">Documento</th><th class="c">Vendedor</th><th class="c">Código</th><th>Cliente</th><th>Categoría</th><th class="r">Venta</th><th class="r">Devolución</th><th class="r">Neto</th></tr>';
    tb.querySelector('tbody').innerHTML = page.length ? page.map(d => `<tr><td class="mono c">${fD(d.fecha)}</td><td class="mono c">${esc(d.documento)}</td><td class="mono c" title="${esc(title(d.nombreVendedor))}">${esc(d.vendedor)}</td>
      <td class="mono c">${esc(d.codigo)}</td><td>${esc(d.cliente)}</td><td>${esc(title(d.cats))}</td>
      <td class="num">${fS2(d.venta)}</td><td class="num rej">${d.devolucion ? fS2(d.devolucion) : ''}</td><td class="num">${fS2(d.neto)}</td></tr>`).join('')
      : '<tr><td colspan="9" class="muted">Sin documentos para este filtro.</td></tr>';
    tb.querySelector('tfoot').innerHTML = `<tr><td colspan="6">Total (${fN(docs.length)} documentos)</td><td class="num">${fS2(tot.venta)}</td><td class="num rej">${fS2(tot.devolucion)}</td><td class="num">${fS2(tot.neto)}</td></tr>`;
  }
  function irADocs(ds) {
    vs.doc = { buscar: ds.cli || '', vend: ds.v || 'all', cat: 'all', cats: ds.cats ? ds.cats.split(',') : null, desde: ds.desde, hasta: ds.hasta, page: 0 };
    setTab('documentos');
    renderDocs();
  }

  // ---------------- Comparativo por día de venta ----------------
  const DIA_ABBR = ['dom', 'lun', 'mar', 'mié', 'jue', 'vie', 'sáb'];
  const diaCorto = f => { if (!f) return ''; const [y, m, d] = f.split('/').map(Number); return `${DIA_ABBR[new Date(y, m - 1, d).getDay()]} ${pad(d)}/${pad(m)}`; };
  const ymPrevio = ym => `${+ym.slice(0, 4) - 1}${ym.slice(4)}`;
  const mesTxt = ym => `${MES_NOMBRE[+ym.slice(5) - 1]} ${ym.slice(0, 4)}`;
  async function ventasPrevio(ym) {
    const py = ymPrevio(ym);
    if (!(py in PREVIO)) {
      overlay(true, `Cargando ventas de ${mesTxt(py)}...`);
      try {
        const rows = await leerVentas(py);
        PREVIO[py] = rows ? C.parseVentas(rows).filter(r => r.fecha.startsWith(py.replace('-', '/'))) : null;
      } catch (e) { console.error(e); PREVIO[py] = null; }
      finally { overlay(false); }
    }
    return PREVIO[py];
  }
  function renderCmpFiltros() {
    $('vCmpCats').innerHTML = '<span class="v-cats-lbl">Categorías</span>' + VCFG.CATEGORIAS.map(c =>
      `<label class="v-chip ${vs.cmp.cats.has(c.codigo) ? 'on' : ''}"><input type="checkbox" value="${c.codigo}" ${vs.cmp.cats.has(c.codigo) ? 'checked' : ''}>${c.nombre}</label>`).join('');
    $('vCmpDiario').classList.toggle('tab-active', vs.cmp.vista === 'diario');
    $('vCmpAcum').classList.toggle('tab-active', vs.cmp.vista === 'acum');
  }
  async function renderCmp() {
    if (!VD) return;
    renderCmpFiltros();
    const ym = VD.ym, py = ymPrevio(ym);
    const tituloTxt = `${mesTxt(py)} vs ${mesTxt(ym)}`;
    $('vCmpTitulo').textContent = tituloTxt;
    document.querySelectorAll('.vCmpTituloInline').forEach(e => { e.textContent = tituloTxt; });
    const previo = await ventasPrevio(ym);
    if (VD.ym !== ym) return; // cambió el mes mientras cargaba
    const aviso2 = $('vCmpSinPrevio');
    aviso2.style.display = previo ? 'none' : '';
    aviso2.innerHTML = previo ? '' : `No se encontró <b>&nbsp;Ventas_${py}.xlsx&nbsp;</b> en Storage: el gráfico muestra solo ${mesTxt(ym)}.`;
    const cats = [...vs.cmp.cats];
    const r = C.comparativo({ actual: VD.ventas, previo: previo || [], cats, vendedor: null, corte: vs.corte });
    VR.comparativo = r;
    const yA = ym.slice(0, 4), yP = py.slice(0, 4);
    const filtroTxt = cats.map(c => CAT[c].nombre).join(' + ') || 'sin categorías';
    $('vCmpFiltroPrint').textContent = filtroTxt;
    const n = r.nAct, v = r.variacion;
    const vTxt = v === null ? '—' : `<span class="${v >= 0 ? 'v-up' : 'v-down'}">${v >= 0 ? '+' : ''}${(v * 100).toFixed(1)}%</span>`;
    const dif = r.prevMismoDia === null ? null : r.acumAct - r.prevMismoDia;
    $('vCmpKpis').innerHTML =
      kpi('navy', `${yA} · ${n} días de venta`, fS(r.acumAct), `al ${diaCorto(vs.corte)}`) +
      kpi('red', `${yP} · mismos ${Math.min(n, r.nPrev)} días`, r.prevMismoDia === null ? '—' : fS(r.prevMismoDia), r.nPrev ? `al ${diaCorto(r.dias[Math.min(n, r.nPrev) - 1].fechaPrev)}` : 'sin datos') +
      kpi(v === null || v >= 0 ? 'green' : 'red', 'Variación al mismo día', vTxt, dif === null ? '' : `${dif >= 0 ? '+' : '−'}${fS(Math.abs(dif))}`) +
      kpi('amber', `${yP} · mes completo`, r.nPrev ? fS(r.prevMes) : '—', r.nPrev ? `${r.nPrev} días de venta` : '');
    $('vCmpLegend').innerHTML = `<span><i class="prev"></i>${title(mesTxt(py))}</span><span><i></i>${title(mesTxt(ym))}</span>`;
    // Gráfico
    const acum = vs.cmp.vista === 'acum';
    const labels = r.dias.map(d => d.n);
    const dA = r.dias.map(d => (acum ? d.acumAct : d.ventaAct));
    const dP = r.dias.map(d => (acum ? d.acumPrev : d.ventaPrev));
    if (typeof Chart === 'undefined') { $('vCmpLegend').innerHTML += '<span class="rej">No se pudo cargar la librería de gráficos (Chart.js).</span>'; }
    else {
      if (CMP_CHART) CMP_CHART.destroy();
      CMP_CHART = new Chart($('vCmpChart'), {
        type: 'line',
        data: { labels, datasets: [
          { label: yP, data: dP, borderColor: '#C1432B', backgroundColor: '#C1432B', borderDash: [6, 4], borderWidth: 2.5, pointRadius: 0, pointHoverRadius: 5, tension: 0.25, spanGaps: false },
          { label: yA, data: dA, borderColor: '#16356B', backgroundColor: '#16356B', borderWidth: 2.5, pointRadius: 0, pointHoverRadius: 5, tension: 0.25, spanGaps: false },
        ] },
        options: {
          responsive: true, maintainAspectRatio: false, animation: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: {
              title: it => `Día de venta ${it[0].label}`,
              label: c => {
                const d = r.dias[c.dataIndex], prev = c.datasetIndex === 0;
                const f = prev ? d.fechaPrev : d.fechaAct;
                return f ? `${c.dataset.label} (${diaCorto(f)}): ${fS(c.parsed.y)}` : '';
              },
              afterBody: it => {
                const d = r.dias[it[0].dataIndex];
                const a = acum ? d.acumAct : d.ventaAct, b = acum ? d.acumPrev : d.ventaPrev;
                return a !== null && b ? `Variación: ${a >= b ? '+' : ''}${((a / b - 1) * 100).toFixed(1)}%` : '';
              },
            } },
          },
          scales: {
            x: { title: { display: true, text: 'Día de venta', color: '#6B675F', font: { family: 'IBM Plex Sans' } }, grid: { display: false }, ticks: { color: '#6B675F', font: { family: 'IBM Plex Mono' } } },
            y: { grid: { color: '#F0EDE5' }, ticks: { color: '#6B675F', font: { family: 'IBM Plex Mono' }, callback: x => 'S/ ' + nf0.format(Math.round(x / 1000)) + 'k' } },
          },
        },
      });
    }
    // Tabla
    const tb = $('vCmpTable');
    tb.querySelector('thead').innerHTML = `<tr class="v-grp"><th></th><th colspan="3" class="v-grp-sep">${yP}</th><th colspan="3" class="v-grp-sep">${yA}</th><th class="v-grp-sep"></th></tr>
      <tr><th class="c">Día de venta</th><th class="c">Fecha</th><th class="r">Venta</th><th class="r">Acumulado</th><th class="c">Fecha</th><th class="r">Venta</th><th class="r">Acumulado</th><th class="c">Var. acumulado</th></tr>`;
    const varCell = (a, b) => (a === null || !b ? '—' : `<span class="${a >= b ? 'v-up' : 'v-down'}">${a >= b ? '+' : ''}${((a / b - 1) * 100).toFixed(1)}%</span>`);
    tb.querySelector('tbody').innerHTML = r.dias.map(d => `<tr><td class="c v-rk">${d.n}</td>
      <td class="mono c">${diaCorto(d.fechaPrev)}</td><td class="num">${d.ventaPrev === null ? '' : fS(d.ventaPrev)}</td><td class="num">${d.acumPrev === null ? '' : fS(d.acumPrev)}</td>
      <td class="mono c">${diaCorto(d.fechaAct)}</td><td class="num">${d.ventaAct === null ? '' : fS(d.ventaAct)}</td><td class="num">${d.acumAct === null ? '' : fS(d.acumAct)}</td>
      <td class="num c">${varCell(d.acumAct, d.acumPrev)}</td></tr>`).join('') || '<tr><td colspan="8" class="muted">Sin días de venta.</td></tr>';
    tb.querySelector('tfoot').innerHTML = `<tr><td class="c">Total</td><td></td><td class="num">${r.nPrev ? fS(r.prevMes) : ''}</td><td></td><td></td><td class="num">${fS(r.acumAct)}</td><td></td>
      <td class="num c">${varCell(r.acumAct, r.prevMismoDia)}</td></tr>`;
  }
  function exportCmp() {
    const r = VR.comparativo; if (!r) return;
    const yA = VD.ym.slice(0, 4), yP = ymPrevio(VD.ym).slice(0, 4);
    const vr = (a, b) => (a === null || !b ? '' : (a / b - 1) * 100);
    downloadStyledXlsx({
      filename: `comparativo_${ymPrevio(VD.ym)}_vs_${VD.ym}.xlsx`, sheetName: 'Comparativo',
      title: `VENTAS SAGA TRANS · Comparativo por día de venta — ${mesTxt(ymPrevio(VD.ym))} vs ${mesTxt(VD.ym)}`,
      subtitle: `${$('vCmpFiltroPrint').textContent} · corte al ${largo(vs.corte)}`,
      columns: [colInt('Día de venta'), colTxt(`Fecha ${yP}`, 12), colMoney(`Venta ${yP}`), colMoney(`Acumulado ${yP}`), colTxt(`Fecha ${yA}`, 12), colMoney(`Venta ${yA}`), colMoney(`Acumulado ${yA}`), colPct('Var. acumulado')],
      rows: r.dias.map(d => [d.n, d.fechaPrev ? fD(d.fechaPrev) : '', d.ventaPrev ?? '', d.acumPrev ?? '', d.fechaAct ? fD(d.fechaAct) : '', d.ventaAct ?? '', d.acumAct ?? '', vr(d.acumAct, d.acumPrev)]),
      totalsRow: ['Total', '', r.prevMes, '', '', r.acumAct, '', vr(r.acumAct, r.prevMismoDia)],
    });
  }

  // ---------------- Meta % (Overrides en ventas-data) ----------------
  async function guardarMeta() {
    const pct = Math.round(vs.meta * 100), ahora = new Date();
    const { data } = await supabaseClient.auth.getSession();
    const registro = { meta: pct, usuario: data?.session?.user?.email || '', fecha: ahora.toISOString(), fechaTexto: ahora.toLocaleString('es-PE', { dateStyle: 'short', timeStyle: 'short' }) };
    const prev = await leerJson(VCFG.DATA_BUCKET, VCFG.META(VD.ym)).catch(() => null);
    const historial = (prev && Array.isArray(prev.historial) ? prev.historial : []).concat([registro]);
    const blob = new Blob([JSON.stringify({ ...registro, historial }, null, 2)], { type: 'application/json' });
    const btn = $('vMetaGuardar'); btn.disabled = true; btn.textContent = 'Guardando...';
    const { error } = await supabaseClient.storage.from(VCFG.DATA_BUCKET).upload(VCFG.META(VD.ym), blob, { upsert: true, contentType: 'application/json', cacheControl: '0' });
    btn.disabled = false; btn.textContent = 'Guardar meta';
    if (error) { console.error(error); renderMetaInfo(`No se pudo guardar la meta: ${error.message}. Revisa la política de escritura de Overrides/ en ventas-data.`); return; }
    vs.metaGuardada = vs.meta; vs.metaRegistro = registro; renderMetaInfo();
  }

  // ---------------- Exportar (reutiliza downloadStyledXlsx de Rechazos) ----------------
  const colCode = h => ({ header: h, width: 12, isCode: true, align: 'center' });
  const colTxt = (h, w) => ({ header: h, width: w || 30 });
  const colMoney = h => ({ header: h, width: 16, numFmt: '#,##0', align: 'right' });
  const colMoney2 = h => ({ header: h, width: 16, numFmt: '#,##0.00', align: 'right' });
  const colInt = h => ({ header: h, width: 12, numFmt: '#,##0', align: 'center' });
  const colPct = h => ({ header: h, width: 13, numFmt: '0.0"%"', align: 'right' });
  const p100 = v => (v === null || v === undefined || !isFinite(v) ? '' : v * 100);
  const metaSub = () => (Math.round(vs.meta * 100) === 100 ? '' : ` · meta ${Math.round(vs.meta * 100)}%`);
  const fileDate = () => toIso(vs.corte);

  function exportAvance() {
    const ag = VR.avance; if (!ag) return;
    const cats = [...vs.cats].map(c => CAT[c].nombre).join(' + ');
    downloadStyledXlsx({
      filename: `avance_ventas_${fileDate()}.xlsx`, sheetName: 'Avance general',
      title: `VENTAS SAGA TRANS · Confitería — Avance de ventas (${cats})`,
      subtitle: `Corte al ${largo(vs.corte)} · teórico ${fT(teorico())} (${VR.hab.transcurridos} de ${VR.hab.total} días hábiles)${metaSub()}`,
      columns: [colCode('Código'), colTxt('Vendedor'), colMoney('Cuota 100%'), colMoney('Cuota'), colMoney('Facturado'), colPct('% Avance'), colPct('Proyección'), colTxt('Estado', 12), colInt('Ranking')],
      rows: ag.filas.map(f => [f.vendedor, title(f.nombre), f.cuota100, f.cuota, f.facturado, p100(f.pct), p100(f.proy), EST_LABEL[f.estado] || '', f.ranking || '']),
      totalsRow: ['', 'Saga Trans Confitería', ag.total.cuota100, ag.total.cuota, ag.total.facturado, p100(ag.total.pct), p100(ag.total.proy), EST_LABEL[ag.total.estado] || '', ''],
    });
  }
  function exportCob(bloque) {
    const cb = VR.cobertura; if (!cb) return;
    const cat = CAT[vs.catCob].nombre, dia = bloque === 'dia';
    const T = dia ? cb.total.dia : cb.total.mes;
    downloadStyledXlsx({
      filename: `cobertura_${cat.toLowerCase()}_${dia ? 'dia' : 'mes'}_${fileDate()}.xlsx`, sheetName: dia ? 'Acumulado del día' : 'Resumen del mes',
      title: `VENTAS SAGA TRANS · Análisis de cobertura ${cat} — ${dia ? 'Acumulado del día' : 'Resumen del mes'}`,
      subtitle: dia ? `Día ${largo(vs.corte)} · restan ${VR.hab.restantes} días hábiles${metaSub()}` : `Acumulado al ${largo(vs.corte)} · teórico ${fT(teorico())}${metaSub()}`,
      columns: dia
        ? [colCode('Código'), colTxt('Vendedor'), colInt('Pendientes'), colInt(`Venta ${cat}`), colPct('% Cobertura'), colMoney('Cuota diaria'), colMoney('Avance'), colPct('% Avance')]
        : [colCode('Código'), colTxt('Vendedor'), colInt('Clientes totales'), colInt(`Venta ${cat}`), colPct('% Cobertura'), colMoney('Cuota'), colMoney('Avance'), colPct('% Avance')],
      rows: cb.filas.map(f => dia
        ? [f.vendedor, title(f.nombre), f.dia.pendientes, f.dia.nuevos, p100(f.dia.pctCob), f.dia.cuotaDiaria ?? '', f.dia.avance, p100(f.dia.pct)]
        : [f.vendedor, title(f.nombre), f.mes.clientes, f.mes.conVenta, p100(f.mes.pctCob), f.mes.cuota, f.mes.avance, p100(f.mes.pct)]),
      totalsRow: dia ? ['', 'Total', T.pendientes, T.nuevos, p100(T.pctCob), T.cuotaDiaria ?? '', T.avance, p100(T.pct)]
        : ['', 'Total', T.clientes, T.conVenta, p100(T.pctCob), T.cuota, T.avance, p100(T.pct)],
    });
  }
  function exportDocs() {
    const docs = VR.docs || [], t = VR.docsTot || { venta: 0, devolucion: 0, neto: 0 };
    downloadStyledXlsx({
      filename: `documentos_ventas_${toIso(vs.doc.desde)}_al_${toIso(vs.doc.hasta)}.xlsx`, sheetName: 'Documentos',
      title: 'VENTAS SAGA TRANS · Confitería — Documentos', subtitle: $('vDocLabel').textContent,
      columns: [colTxt('Fecha', 12), colCode('Documento'), colCode('Vendedor'), colCode('Código'), colTxt('Cliente', 34), colTxt('Categoría', 16), colMoney2('Venta'), colMoney2('Devolución'), colMoney2('Neto')],
      rows: docs.map(d => [fD(d.fecha), d.documento, d.vendedor, d.codigo, d.cliente, title(d.cats), d.venta, d.devolucion || '', d.neto]),
      totalsRow: ['Total', '', '', '', `${docs.length} documentos`, '', t.venta, t.devolucion, t.neto],
    });
  }
  function imprimir(id) {
    const node = $(id); node.classList.add('print-target'); window.print();
    setTimeout(() => node.classList.remove('print-target'), 500);
  }

  // ---------------- Pestañas y módulos ----------------
  function setTab(tab) {
    vs.tab = tab;
    document.querySelectorAll('.v-tabbar .tab-btn').forEach(b => b.classList.toggle('tab-active', b.dataset.vtab === tab));
    ['avance', 'cobertura', 'comparativo', 'documentos', 'alertas'].forEach(t => { $('vpanel-' + t).style.display = t === tab ? '' : 'none'; });
    if (tab === 'comparativo' && VD) renderCmp();
  }
  async function mostrarModulo(nombre) {
    $('modRechazos').style.display = nombre === 'rechazos' ? '' : 'none';
    $('modVentas').style.display = nombre === 'ventas' ? '' : 'none';
    $('railRechazos').classList.toggle('rail-active', nombre === 'rechazos');
    $('railVentas').classList.toggle('rail-active', nombre === 'ventas');
    window.scrollTo(0, 0);
    if (nombre !== 'ventas') return;
    const { data } = await supabaseClient.auth.getSession();
    const email = data?.session?.user?.email || '';
    $('vUserEmail').textContent = email; $('vUserAvatar').textContent = (email || '?').charAt(0).toUpperCase();
    if (!vs.booted) {
      vs.booted = true;
      try { await aplicarCorte(); } catch (e) {
        console.error(e); vs.booted = false;
        aviso(`No se pudo cargar el módulo de Ventas: ${esc(e.message)}`, true);
      }
    }
  }

  function wire() {
    $('railRechazos').addEventListener('click', () => mostrarModulo('rechazos'));
    $('railVentas').addEventListener('click', () => mostrarModulo('ventas'));
    $('vLogoutBtn').addEventListener('click', async () => { try { await supabaseClient.auth.signOut(); } catch (e) { console.error(e); } });
    initControles();
    document.querySelectorAll('.v-tabbar .tab-btn').forEach(b => b.addEventListener('click', () => setTab(b.dataset.vtab)));
    document.querySelectorAll('[data-vprint]').forEach(b => b.addEventListener('click', () => imprimir(b.dataset.vprint)));
    $('vBtnCalendario').addEventListener('click', () => {
      const box = $('vCalendario'), abrir = box.style.display === 'none';
      box.style.display = abrir ? '' : 'none';
      $('vBtnCalendarioTxt').textContent = abrir ? 'Ocultar calendario' : 'Ver calendario';
      $('vBtnCalendario').setAttribute('aria-expanded', String(abrir));
    });
    $('vBtnRecargar').addEventListener('click', () => VD && aplicarCorte({ forzar: true }));
    $('vCats').addEventListener('change', e => {
      if (e.target.type !== 'checkbox') return;
      e.target.checked ? vs.cats.add(e.target.value) : vs.cats.delete(e.target.value);
      renderCats(); renderAvance();
    });
    $('vMeta').addEventListener('input', e => {
      const v = parseFloat(e.target.value);
      if (!(v > 0 && v <= 200)) { renderMetaInfo('Ingresa una meta entre 1% y 200%.'); return; }
      vs.meta = v / 100; renderMetaInfo(); renderAvance(); renderCobertura();
    });
    $('vMetaGuardar').addEventListener('click', guardarMeta);
    $('vCobCats').addEventListener('click', e => { const b = e.target.closest('[data-vcat]'); if (b) { vs.catCob = b.dataset.vcat; renderCobertura(); } });
    $('modVentas').addEventListener('click', e => { const td = e.target.closest('td.clickable[data-desde]'); if (td) irADocs(td.dataset); });
    $('vExportAvance').addEventListener('click', exportAvance);
    $('vExportCobMes').addEventListener('click', () => exportCob('mes'));
    $('vExportCobDia').addEventListener('click', () => exportCob('dia'));
    $('vExportDocs').addEventListener('click', exportDocs);
    // Comparativo
    $('vCmpCats').addEventListener('change', e => {
      if (e.target.type !== 'checkbox') return;
      e.target.checked ? vs.cmp.cats.add(e.target.value) : vs.cmp.cats.delete(e.target.value);
      renderCmp();
    });
    $('vCmpDiario').addEventListener('click', () => { vs.cmp.vista = 'diario'; renderCmp(); });
    $('vCmpAcum').addEventListener('click', () => { vs.cmp.vista = 'acum'; renderCmp(); });
    $('vExportCmp').addEventListener('click', exportCmp);
    // Filtros de Documentos
    $('vDocSearch').addEventListener('input', e => { vs.doc.buscar = e.target.value; vs.doc.page = 0; dibujarDocs(); });
    $('vDocVend').addEventListener('change', e => { vs.doc.vend = e.target.value; vs.doc.page = 0; dibujarDocs(); });
    $('vDocCat').addEventListener('change', e => { if (e.target.value !== 'custom') { vs.doc.cats = null; vs.doc.cat = e.target.value; } vs.doc.page = 0; dibujarDocs(); });
    $('vDocDesde').addEventListener('change', e => { if (e.target.value) { vs.doc.desde = fromIso(e.target.value); vs.doc.page = 0; dibujarDocs(); } });
    $('vDocHasta').addEventListener('change', e => { if (e.target.value) { vs.doc.hasta = fromIso(e.target.value); vs.doc.page = 0; dibujarDocs(); } });
    $('vDocLimpiar').addEventListener('click', () => { vs.doc = { buscar: '', vend: 'all', cat: 'all', cats: null, desde: vs.corte.slice(0, 8) + '01', hasta: vs.corte, page: 0 }; renderDocs(); });
    $('vDocPrev').addEventListener('click', () => { vs.doc.page = Math.max(0, vs.doc.page - 1); dibujarDocs(); });
    $('vDocNext').addEventListener('click', () => { vs.doc.page += 1; dibujarDocs(); });
  }

  if (!C) { console.error('Módulo Ventas: falta js/ventas-calc.js (cárgalo antes de js/ventas.js).'); return; }
  if (!$('modVentas') || !$('railVentas')) { console.error('Módulo Ventas: falta el bloque #modVentas o los ids del menú en index.html.'); return; }
  wire();
})();
