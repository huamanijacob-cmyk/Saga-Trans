/* =====================================================================
   Módulo Notas de crédito — Saga Trans Confitería (js/nc.js)
   Se carga después de app.js, ventas.js y nc-calc.js. Reutiliza:
     supabaseClient · loadModuleFile · FILE_CACHE · activarModulo · nombreDesdeCorreo
   Diagnóstico desde la consola (F12): NCMOD.data · NCMOD.res · NCMOD.state
   ===================================================================== */
(function () {
  'use strict';
  const C = window.NCCalc;
  const $ = id => document.getElementById(id);
  if (!C || !$('modNC')) { console.error('Módulo NC: falta nc-calc.js o el bloque #modNC en index.html'); return; }

  const CFG = {
    NC_FOLDER: 'ContabilidadNC', VENTAS_FOLDER: 'Ventas', BUCKET: 'ventas-data',
    AJUSTES: m => `Overrides/nc_ajustes_${m}.json`,
    REGISTRO: m => `Overrides/nc_registro_${m}.json`,
    ALEX: m => `Overrides/alex_${m}.xlsx`,
    POR_PAG: 15,
  };
  const MES_ABBR = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Set', 'Oct', 'Nov', 'Dic'];
  const MES_NOMBRE = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'setiembre', 'octubre', 'noviembre', 'diciembre'];
  const CLASE = { rechazo: 'Rechazo', parcial: 'Parcial', error: 'Error' };
  const ESTADO = { coincide: 'Coincide', distinto: 'Monto distinto', rechazo: 'Rechazo', parcial: 'Parcial', facanulada: 'Factura anulada', falta: 'Falta emitir NC', sinalex: 'NC sin Alex' };
  const GRUPO = { regular: 'Regular', adicional: 'Adicional', paneton: 'Panetón' };

  const st = {
    booted: false, modo: 'mes', mesSel: null, diaSel: null, ym: null, corte: null, tab: 'resumen', igv: 'c',
    anu: { filtro: 'todas', page: 0, cambios: {} },
    con: { filtro: 'pend', page: 0 },
    doc: { buscar: '', grupo: 'all', page: 0, sel: new Set() },
    correo: {}, correoS: {},
  };
  let D = null, R = null, CHART = null;
  window.NCMOD = { state: st, get data() { return D; }, get res() { return R; } };

  // ---------------- Formato ----------------
  const pad = n => String(n).padStart(2, '0');
  const nf2 = new Intl.NumberFormat('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const nf0 = new Intl.NumberFormat('es-PE', { maximumFractionDigits: 0 });
  const fS = n => (n === null || n === undefined || !isFinite(n) ? '—' : 'S/ ' + nf2.format(n));
  const fN = n => nf0.format(n || 0);
  const f2 = n => (n === null || n === undefined ? '' : nf2.format(n));
  const fP = n => (n === null || n === undefined || !isFinite(n) ? '—' : (n * 100).toFixed(2) + '%');
  const fP1 = n => (n === null || n === undefined || !isFinite(n) ? '—' : (n * 100).toFixed(1) + '%');
  const fD = f => (f ? `${f.slice(8, 10)}/${f.slice(5, 7)}/${f.slice(0, 4)}` : '');
  const toIso = f => (f ? f.replace(/\//g, '-') : '');
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const largo = f => { const [y, m, d] = f.split('/').map(Number); return `${d} de ${MES_NOMBRE[m - 1]} del ${y}`; };
  const chip = (cls, t) => `<span class="nc-chip nc-${cls}">${esc(t)}</span>`;
  const kpi = (color, label, value, sub) => `<div class="kpi kpi-${color}"><div class="kpi-label">${label}</div><div class="kpi-val">${value}</div>${sub ? `<div class="v-kpi-sub">${sub}</div>` : ''}</div>`;
  function aviso(html, err) { const d = document.createElement('div'); d.className = 'pendiente-banner' + (err ? ' v-aviso-err' : ''); d.innerHTML = html; $('ncAvisos').appendChild(d); }

  // ---------------- Storage ----------------
  function overlay(on, t) {
    const o = $('loadingOverlay'); if (!o) return; const lt = o.querySelector('.loading-text');
    if (on) { o.dataset.prev = o.dataset.prev || (lt ? lt.textContent : ''); if (lt) lt.textContent = t || 'Cargando…'; o.style.display = ''; }
    else { o.style.display = 'none'; if (lt && o.dataset.prev) lt.textContent = o.dataset.prev; }
  }
  async function sesion() { const { data } = await supabaseClient.auth.getSession(); return data?.session || null; }
  async function descargar(path) {
    const s = await sesion();
    const url = `${SUPABASE_URL}/storage/v1/object/${CFG.BUCKET}/${path.split('/').map(encodeURIComponent).join('/')}?cachebust=${Date.now()}`;
    const res = await fetch(url, { cache: 'no-store', headers: { Authorization: `Bearer ${s?.access_token}`, apikey: SUPABASE_ANON_KEY } });
    if (res.status === 400 || res.status === 404) return null;
    if (!res.ok) throw new Error(`HTTP ${res.status} al leer ${path}`);
    return res.arrayBuffer();
  }
  async function leerJson(path) { const b = await descargar(path); if (!b) return null; try { return JSON.parse(new TextDecoder().decode(b)); } catch { return null; } }
  async function subir(path, blob, tipo) {
    const { error } = await supabaseClient.storage.from(CFG.BUCKET).upload(path, blob, { upsert: true, contentType: tipo, cacheControl: '0' });
    if (error) throw error;
  }
  const guardarJson = (path, obj) => subir(path, new Blob([JSON.stringify(obj)], { type: 'application/json' }), 'application/json');

  // Lee todas las hojas del libro de Alex. El archivo suele declarar ~1 millón de filas vacías:
  // se recalcula el rango real para no recorrerlas.
  function hojasAlex(buf) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true, dense: false });
    const out = {};
    wb.SheetNames.forEach(n => {
      const ws = wb.Sheets[n]; let maxR = 0, maxC = 0;
      Object.keys(ws).forEach(k => { if (k[0] === '!') return; const c = XLSX.utils.decode_cell(k); if (c.r > maxR) maxR = c.r; if (c.c > maxC) maxC = c.c; });
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: Math.max(maxC, 15) } });
      out[n] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    });
    return out;
  }

  // ---------------- Carga del mes ----------------
  async function cargarMes(ym, { forzar } = {}) {
    overlay(true, 'Cargando notas de crédito...');
    $('ncAvisos').innerHTML = '';
    if (forzar && typeof FILE_CACHE !== 'undefined') { delete FILE_CACHE[`${CFG.NC_FOLDER}/${CFG.NC_FOLDER}_${ym}.xlsx`]; delete FILE_CACHE[`${CFG.VENTAS_FOLDER}/${CFG.VENTAS_FOLDER}_${ym}.xlsx`]; }
    try {
      const [ncRows, vRows, ajustes, registro, alexBuf] = await Promise.all([
        loadModuleFile(CFG.NC_FOLDER, CFG.NC_FOLDER, ym),
        loadModuleFile(CFG.VENTAS_FOLDER, CFG.VENTAS_FOLDER, ym),
        leerJson(CFG.AJUSTES(ym)).catch(() => null),
        leerJson(CFG.REGISTRO(ym)).catch(() => null),
        descargar(CFG.ALEX(ym)).catch(e => { console.error(e); return null; }),
      ]);
      const nc = C.parseNC(ncRows || []);
      if (!nc.length) aviso(`Todavía no hay notas de crédito de ${MES_NOMBRE[+ym.slice(5) - 1]} ${ym.slice(0, 4)}. Sube <b>&nbsp;ContabilidadNC_${ym}.xlsx&nbsp;</b> a Storage.`);
      if (!vRows || !vRows.length) aviso(`No se encontró <b>&nbsp;Ventas_${ym}.xlsx&nbsp;</b>: el sell out y el % de descuento de cada factura quedan en blanco.`);
      let alex = [];
      if (alexBuf) { try { alex = C.parseAlex(hojasAlex(alexBuf)); } catch (e) { console.error(e); aviso('No se pudo leer el archivo de Alex guardado para este mes.', true); } }
      D = { ym, nc, ventas: vRows || [], alex, ajustes: ajustes || { marcas: {}, manuales: [], correo: {} }, registro: registro || {} };
      D.ajustes.marcas = D.ajustes.marcas || {}; D.ajustes.manuales = D.ajustes.manuales || []; D.ajustes.correo = D.ajustes.correo || {};
      // Guardar montos de NC vigentes para conocerlos si luego se anulan (el sistema los deja en 0).
      const nuevos = C.nuevosParaRegistro(nc, D.registro);
      if (Object.keys(nuevos).length) {
        Object.assign(D.registro, nuevos);
        guardarJson(CFG.REGISTRO(ym), D.registro).catch(e => console.warn('No se pudo actualizar el registro de montos:', e.message));
      }
    } finally { overlay(false); }
  }

  function calcular() {
    R = C.calcular({ nc: D.nc, ventas: D.ventas, alex: D.alex, ajustes: D.ajustes, registro: D.registro, corte: st.corte, ym: D.ym });
  }

  async function aplicarCorte({ forzar } = {}) {
    const ym = st.modo === 'dia' ? st.diaSel.slice(0, 7) : st.mesSel;
    if (!D || D.ym !== ym || forzar) await cargarMes(ym, { forzar });
    st.ym = ym;
    if (st.modo === 'dia') st.corte = st.diaSel.replace(/-/g, '/');
    else { const fs = C.fechasNC(D.nc); const [y, m] = ym.split('-').map(Number); st.corte = fs.length ? fs[fs.length - 1] : `${y}/${pad(m)}/${pad(new Date(y, m, 0).getDate())}`; }
    st.anu.page = st.con.page = st.doc.page = 0; st.anu.cambios = {}; st.correo = {}; st.correoS = {}; st.doc.sel.clear();
    calcular(); renderTodo();
  }

  // ---------------- Render ----------------
  function renderTodo() {
    if (!R) return;
    $('ncCorteLabel').textContent = largo(st.corte);
    const fa = D.alex.length ? D.alex.map(a => a.fecha).sort().slice(-1)[0] : '';
    $('ncAlexLabel').textContent = D.alex.length ? ` · Alex actualizado al ${fD(fa).slice(0, 5)}` : ' · sin archivo de Alex';
    renderResumen(); renderDiario(); renderAnuladas(); renderConcil(); renderProducto(); renderAjustes(); renderDocs();
  }

  function renderResumen() {
    const t = R.tot;
    const ajTxt = (a) => (a ? ` · ajuste ${fS(a)}` : '');
    $('ncKpis1').innerHTML =
      kpi('navy', 'Nestlé', fS(t.N), `s/IGV ${f2(t.Ns)} · ${fP1(t.shareN)}${ajTxt(t.ajN)}`) +
      kpi('teal', 'ST', fS(t.S), `s/IGV ${f2(t.Ss)} · ${fP1(t.shareS)}${ajTxt(t.ajS)}`) +
      kpi('navy', 'Total descuentos', fS(t.total), `s/IGV ${f2(t.totals)}`) +
      kpi('red', '%DCTOS ST', fP(t.pctST), `c/IGV ${fP(t.pctSTc)} · sell out ${fS(t.sellout)}`);
    const nR = R.anuladas.filter(a => a.clase !== 'error');
    $('ncKpis2').innerHTML =
      kpi('red', 'Rechazos Nestlé', fS(t.rechN), `${nR.filter(a => a.cargo === 'N').length} notas anuladas`) +
      kpi('red', 'Rechazos ST', fS(t.rechS), `${nR.filter(a => a.cargo === 'S').length} notas anuladas`) +
      kpi('navy', 'Adicionales', fS(t.adicional), `${R.vigentes.filter(r => r.grupo === 'adicional').length} notas`) +
      kpi('amber', 'Panetón', fS(t.paneton), `${R.vigentes.filter(r => r.grupo === 'paneton').length} notas`);
    const [y, m] = D.ym.split('-').map(Number);
    $('ncCurvaTitulo').textContent = `% Descuentos ST · ${MES_NOMBRE[m - 1]} ${y}`;
    if (typeof Chart === 'undefined') return;
    if (CHART) CHART.destroy();
    const pts = R.curva.map(c => ({ x: c.fecha.slice(8), y: c.pct === null ? null : +(c.pct * 100).toFixed(2) }));
    CHART = new Chart($('ncCurva'), {
      type: 'line',
      data: { labels: pts.map(p => p.x), datasets: [{ data: pts.map(p => p.y), borderColor: '#16356B', backgroundColor: '#16356B', pointRadius: 5, pointHoverRadius: 6, borderWidth: 2, tension: 0.25 }] },
      options: {
        responsive: true, maintainAspectRatio: false, animation: false, layout: { padding: { top: 20, right: 24, left: 8 } },
        plugins: { legend: { display: false }, tooltip: { callbacks: { title: i => `Día ${i[0].label}`, label: c => `${c.parsed.y.toFixed(2)}%` } } },
        scales: { x: { grid: { display: false }, ticks: { color: '#6B675F', font: { family: 'IBM Plex Mono' } } }, y: { min: 0, grid: { color: '#F0EDE5' }, ticks: { color: '#6B675F', font: { family: 'IBM Plex Mono' }, callback: v => v.toFixed(1) + '%' } } },
      },
      plugins: [{ id: 'etiquetas', afterDatasetsDraw(ch) { const g = ch.ctx; g.save(); g.font = '600 10px "IBM Plex Mono", monospace'; g.fillStyle = '#16356B'; g.textAlign = 'center'; ch.getDatasetMeta(0).data.forEach((p, i) => { const v = pts[i].y; if (v !== null) g.fillText(v.toFixed(2) + '%', p.x, p.y - 10); }); g.restore(); } }],
    });
  }

  function renderDiario() {
    document.querySelectorAll('#ncIgvSel .tab-btn').forEach(b => b.classList.toggle('tab-active', b.dataset.igv === st.igv));
    const sN = R.series.filter(s => s.charAt(1) === 'N'), sS = R.series.filter(s => s.charAt(1) === 'S');
    const k = st.igv; const tb = $('ncDiarioTable');
    tb.querySelector('thead').innerHTML = `<tr><th class="c">Fecha</th>${sN.map(s => `<th class="r">${s}</th>`).join('')}<th class="r">Nestlé</th>${sS.map(s => `<th class="r">${s}</th>`).join('')}<th class="r">ST</th><th class="r">Total</th></tr>`;
    const tot = {}; const sum = (d, arr) => arr.reduce((a, s) => a + (d[k][s] || 0), 0);
    tb.querySelector('tbody').innerHTML = R.diario.map(d => {
      [...sN, ...sS].forEach(s => { tot[s] = (tot[s] || 0) + (d[k][s] || 0); });
      const n = sum(d, sN), s = sum(d, sS);
      return `<tr><td class="mono c">${fD(d.fecha)}</td>${sN.map(x => `<td class="num">${f2(d[k][x] || 0)}</td>`).join('')}<td class="num"><b>${f2(n)}</b></td>${sS.map(x => `<td class="num">${f2(d[k][x] || 0)}</td>`).join('')}<td class="num"><b>${f2(s)}</b></td><td class="num">${f2(n + s)}</td></tr>`;
    }).join('') || '<tr><td class="muted" colspan="9">Sin notas de crédito vigentes en el período.</td></tr>';
    const tn = sN.reduce((a, s) => a + (tot[s] || 0), 0), ts = sS.reduce((a, s) => a + (tot[s] || 0), 0);
    tb.querySelector('tfoot').innerHTML = `<tr><td class="c">Total</td>${sN.map(s => `<td class="num">${f2(tot[s] || 0)}</td>`).join('')}<td class="num">${f2(tn)}</td>${sS.map(s => `<td class="num">${f2(tot[s] || 0)}</td>`).join('')}<td class="num">${f2(ts)}</td><td class="num">${f2(tn + ts)}</td></tr>`;
  }

  function renderAnuladas() {
    const A = R.anuladas;
    const cuenta = c => A.filter(a => a.clase === c);
    const monto = arr => arr.reduce((s, a) => s + a.montoAnulada, 0);
    $('ncAnuKpis').innerHTML = kpi('red', 'Rechazo', fN(cuenta('rechazo').length), fS(monto(cuenta('rechazo')))) + kpi('amber', 'Parcial', fN(cuenta('parcial').length), fS(monto(cuenta('parcial')))) + kpi('navy', 'Error', fN(cuenta('error').length));
    document.querySelectorAll('#ncAnuFiltro .tab-btn').forEach(b => b.classList.toggle('tab-active', b.dataset.f === st.anu.filtro));
    const lista = A.filter(a => st.anu.filtro === 'todas' || a.clase === st.anu.filtro).sort((a, b) => a.fecha.localeCompare(b.fecha) || a.doc.localeCompare(b.doc));
    const per = CFG.POR_PAG, max = Math.max(0, Math.ceil(lista.length / per) - 1); st.anu.page = Math.min(st.anu.page, max);
    $('ncAnuPage').textContent = `${st.anu.page + 1} / ${max + 1}`; $('ncAnuPrev').disabled = st.anu.page === 0; $('ncAnuNext').disabled = st.anu.page === max;
    const tb = $('ncAnuTable');
    tb.querySelector('thead').innerHTML = '<tr><th class="c">Nota de crédito</th><th class="c">Emisión</th><th class="c">Anulada</th><th class="c">Cliente</th><th class="c">Asume</th><th class="r">Monto</th><th class="c">Origen</th><th class="c">Clasificación</th></tr>';
    const ORIG = { alex: 'Alex', registro: 'Registro', manual: 'Manual', '': '—' };
    tb.querySelector('tbody').innerHTML = lista.slice(st.anu.page * per, st.anu.page * per + per).map(a => {
      const cam = st.anu.cambios[a.doc] || {};
      const clase = cam.tipo || a.clase; const m = cam.monto !== undefined ? cam.monto : a.montoAnulada;
      return `<tr><td class="mono c">${esc(a.doc)}</td><td class="mono c">${fD(a.fecha).slice(0, 5)}</td><td class="mono c">${esc(a.fecanu ? fD(a.fecanu).slice(0, 5) : '—')}</td><td class="mono c">${esc(a.codcli)}</td><td class="c">${a.cargo === 'N' ? 'Nestlé' : 'ST'}</td>
        <td class="num">${clase === 'error' ? '—' : `<input class="nc-inp ${cam.monto !== undefined ? 'dirty' : ''}" type="number" step="0.01" data-doc="${esc(a.doc)}" value="${m}">`}</td>
        <td class="c nc-origen">${ORIG[a.origen] || '—'}</td>
        <td class="c"><select class="nc-sel ${cam.tipo ? 'dirty' : ''}" data-doc="${esc(a.doc)}">${Object.entries(CLASE).map(([k, v]) => `<option value="${k}" ${k === clase ? 'selected' : ''}>${v}</option>`).join('')}</select></td></tr>`;
    }).join('') || '<tr><td class="muted" colspan="8">No hay notas anuladas en este filtro.</td></tr>';
    $('ncAnuGuardar').style.display = Object.keys(st.anu.cambios).length ? '' : 'none';
  }

  function renderConcil() {
    const c = R.concil;
    $('ncConcilVacio').style.display = c ? 'none' : '';
    $('ncConcilBox').style.display = c ? '' : 'none';
    if (!c) return;
    const n = k => c.cuenta[k] || 0;
    const defs = [['coincide', 'Coinciden', 'green', n('coincide')], ['distinto', 'Monto distinto', 'amber', n('distinto')], ['rechparc', 'Rechazo / parcial', 'red', n('rechazo') + n('parcial') + n('facanulada')], ['falta', 'Falta emitir NC', 'navy', n('falta')], ['sinalex', 'NC sin Alex', 'teal', n('sinalex')]];
    $('ncConKpis').innerHTML = defs.map(([k, l, col, v]) => `<div class="kpi kpi-${col} ${st.con.filtro === k ? 'on' : ''}" data-cf="${k}"><div class="kpi-label">${l}</div><div class="kpi-val">${fN(v)}</div></div>`).join('');
    const filtro = st.con.filtro;
    const lista = c.filas.filter(f => filtro === 'pend' ? f.estado !== 'coincide' : filtro === 'rechparc' ? (f.estado === 'rechazo' || f.estado === 'parcial' || f.estado === 'facanulada') : f.estado === filtro)
      .sort((a, b) => a.fecha.localeCompare(b.fecha) || a.doc.localeCompare(b.doc));
    $('ncConTitulo').textContent = (c.finAlex ? `Alex hasta el ${fD(c.finAlex).slice(0, 5)} · ` : '') + (filtro === 'pend' ? `Por revisar · ${fN(lista.length)} facturas` : `${defs.find(d => d[0] === filtro)?.[1] || ''} · ${fN(lista.length)} facturas`);
    const per = CFG.POR_PAG, max = Math.max(0, Math.ceil(lista.length / per) - 1); st.con.page = Math.min(st.con.page, max);
    $('ncConPage').textContent = `${st.con.page + 1} / ${max + 1}`; $('ncConPrev').disabled = st.con.page === 0; $('ncConNext').disabled = st.con.page === max;
    const tb = $('ncConTable');
    tb.querySelector('thead').innerHTML = '<tr><th class="c">Fecha</th><th class="c">Factura</th><th class="c">Cliente</th><th>Razón social</th><th class="r">Alex Nestlé</th><th class="r">Sistema Nestlé</th><th class="r">Dif.</th><th class="r">Alex ST</th><th class="r">Sistema ST</th><th class="r">Dif.</th><th class="c">Resultado</th></tr>';
    const dif = v => (Math.abs(v) <= 0.05 ? '<span class="muted">0.00</span>' : `<span class="${v < 0 ? 'rej' : ''}">${f2(v)}</span>`);
    tb.querySelector('tbody').innerHTML = lista.slice(st.con.page * per, st.con.page * per + per).map(f => `<tr><td class="mono c">${fD(f.fecha).slice(0, 5)}</td><td class="mono c" title="${esc(f.ncs.join(' · '))}">${esc(f.doc)}</td><td class="mono c">${esc(f.cli)}</td><td>${esc(f.razon)}</td>
      <td class="num">${f2(f.alexN)}</td><td class="num">${f2(f.sisN)}</td><td class="num">${dif(f.dN)}</td><td class="num">${f2(f.alexS)}</td><td class="num">${f2(f.sisS)}</td><td class="num">${dif(f.dS)}</td><td class="c">${chip(f.estado === 'facanulada' ? 'rechazo' : f.estado, ESTADO[f.estado] + (f.auto ? ' · auto' : ''))}</td></tr>`).join('')
      || '<tr><td class="muted" colspan="11">Nada que revisar en este filtro.</td></tr>';
    // Sr. Alex
    const ta = $('ncSrAlexTable');
    ta.querySelector('thead').innerHTML = '<tr class="v-grp"><th></th><th colspan="4" class="v-grp-sep">Nestlé</th><th colspan="4" class="v-grp-sep">ST</th></tr>'
      + '<tr><th class="c">Fecha</th><th class="r">Diario (Alex)</th><th class="r">Acumulado</th><th class="c">Informa correo</th><th class="r">Diferencia</th><th class="r">Diario (Alex)</th><th class="r">Acumulado</th><th class="c">Informa correo</th><th class="r">Diferencia</th></tr>';
    const celdas = (fecha, acum, guardado, pend, attr) => {
      const val = pend[fecha] !== undefined ? pend[fecha] : (guardado ?? '');
      const d = val === '' || val === null ? null : Math.round((Number(val) - acum) * 100) / 100;
      return `<td class="c"><input class="nc-inp ${pend[fecha] !== undefined ? 'dirty' : ''}" type="number" step="0.01" ${attr}="${fecha}" value="${val}"></td>
        <td class="num">${d === null ? '—' : (Math.abs(d) <= 0.05 ? chip('coincide', '0.00') : `<span class="rej">${f2(d)}</span>`)}</td>`;
    };
    ta.querySelector('tbody').innerHTML = c.srAlex.map(r => `<tr><td class="mono c">${fD(r.fecha)}</td>
      <td class="num">${f2(r.diario)}</td><td class="num">${f2(r.acumulado)}</td>${celdas(r.fecha, r.acumulado, r.correo, st.correo, 'data-correo')}
      <td class="num">${f2(r.diarioS)}</td><td class="num">${f2(r.acumuladoS)}</td>${celdas(r.fecha, r.acumuladoS, r.correoS, st.correoS, 'data-correos')}</tr>`).join('');
    $('ncCorreoGuardar').style.display = Object.keys(st.correo).length || Object.keys(st.correoS).length ? '' : 'none';
  }

  function renderProducto() {
    const tb = $('ncProdTable'); const P = R.productos;
    tb.querySelector('thead').innerHTML = '<tr><th>Producto / promoción</th><th class="c">Facturas</th><th class="r">Nestlé</th><th class="r">ST</th><th class="r">Total</th></tr>';
    tb.querySelector('tbody').innerHTML = P.map(p => `<tr><td>${esc(p.producto)}</td><td class="num c">${fN(p.facturas)}</td><td class="num">${f2(p.N)}</td><td class="num">${f2(p.S)}</td><td class="num">${f2(p.total)}</td></tr>`).join('')
      || `<tr><td class="muted" colspan="5">${D.alex.length ? 'El archivo de Alex no trae la columna de producto (solo la traen las hojas por día).' : 'Sube el archivo de Alex para ver los descuentos por producto.'}</td></tr>`;
    const s = P.reduce((a, p) => ({ f: a.f + p.facturas, N: a.N + p.N, S: a.S + p.S }), { f: 0, N: 0, S: 0 });
    tb.querySelector('tfoot').innerHTML = P.length ? `<tr><td>Total</td><td class="num c">${fN(s.f)}</td><td class="num">${f2(s.N)}</td><td class="num">${f2(s.S)}</td><td class="num">${f2(s.N + s.S)}</td></tr>` : '';
  }

  function renderAjustes() {
    const tb = $('ncAjTable'); const M = D.ajustes.marcas;
    tb.querySelector('thead').innerHTML = '<tr><th>Concepto</th><th class="c">Tipo</th><th class="c">Notas</th><th class="r">Monto c/IGV</th><th>Registrado por</th><th class="c"></th></tr>';
    const adi = {};
    R.vigentes.filter(r => r.grupo === 'adicional').forEach(r => { const e = r.etiqueta || 'Adicional'; const a = (adi[e] = adi[e] || { n: 0, m: 0, quien: '' }); a.n++; a.m += r.mondoc; const mk = M[r.doc]; if (mk) a.quien = `${mk.usuario || ''} · ${mk.fechaTexto || ''}`; });
    const manAnu = Object.entries(M).filter(([, v]) => ['rechazo', 'parcial', 'error'].includes(v.tipo));
    let h = Object.entries(adi).map(([e, a]) => `<tr><td>${esc(e)}</td><td class="c">${chip('adicional', 'Adicional')}</td><td class="num c">${a.n}</td><td class="num">${f2(a.m)}</td><td>${esc(a.quien)}</td><td class="c"><button class="clear-btn" data-quitar-etq="${esc(e)}">Quitar</button></td></tr>`).join('');
    h += D.ajustes.manuales.map((a, i) => `<tr><td>${esc(a.concepto)}</td><td class="c">${chip('distinto', a.cargo === 'N' ? 'Ajuste Nestlé' : 'Ajuste ST')}</td><td class="c">—</td><td class="num">${f2(Number(a.monto))}</td><td>${esc(`${a.usuario || ''} · ${a.fechaTexto || ''}`)}</td><td class="c"><button class="clear-btn" data-quitar-aj="${i}">Quitar</button></td></tr>`).join('');
    if (manAnu.length) h += `<tr><td>Anuladas clasificadas a mano</td><td class="c">${chip('rechazo', 'Anuladas')}</td><td class="num c">${manAnu.length}</td><td class="num">${f2(manAnu.reduce((s, [, v]) => s + (Number(v.monto) || 0), 0))}</td><td></td><td></td></tr>`;
    tb.querySelector('tbody').innerHTML = h || '<tr><td class="muted" colspan="6">No hay adicionales ni ajustes en este mes. Los adicionales se marcan desde la pestaña Documentos.</td></tr>';
  }

  function docsFiltrados() {
    const q = st.doc.buscar.trim().toLowerCase();
    return R.documentos.filter(d => (st.doc.grupo === 'all' || d.grupo === st.doc.grupo) &&
      (!q || d.factura.toLowerCase().includes(q) || d.cli.includes(q) || d.cliente.toLowerCase().includes(q) || d.ncs.some(n => n.toLowerCase().includes(q))));
  }
  function renderDocs() {
    const lista = docsFiltrados();
    const per = CFG.POR_PAG, max = Math.max(0, Math.ceil(lista.length / per) - 1); st.doc.page = Math.min(st.doc.page, max);
    $('ncDocPage').textContent = `${st.doc.page + 1} / ${max + 1}`; $('ncDocPrev').disabled = st.doc.page === 0; $('ncDocNext').disabled = st.doc.page === max;
    const pag = lista.slice(st.doc.page * per, st.doc.page * per + per);
    const tb = $('ncDocTable');
    const todos = pag.length && pag.every(d => st.doc.sel.has(d.factura));
    tb.querySelector('thead').innerHTML = `<tr><th class="c"><input type="checkbox" id="ncSelPag" ${todos ? 'checked' : ''} aria-label="Seleccionar página"></th><th class="c">Fecha</th><th class="c">Factura</th><th class="c">Cliente</th><th>Razón social</th><th class="r">Monto c/IGV</th><th class="c">% Nestlé</th><th class="c">% ST</th><th class="c">% Total</th><th class="c">Notas de crédito</th><th>Producto</th><th class="c">Grupo</th></tr>`;
    tb.querySelector('tbody').innerHTML = pag.map(d => `<tr><td class="c"><input type="checkbox" data-sel="${esc(d.factura)}" ${st.doc.sel.has(d.factura) ? 'checked' : ''} aria-label="Seleccionar ${esc(d.factura)}"></td><td class="mono c">${fD(d.fecha).slice(0, 5)}</td><td class="mono c">${esc(d.factura)}</td><td class="mono c">${esc(d.cli)}</td><td>${esc(d.cliente)}</td>
      <td class="num">${f2(d.ventaIgv)}</td><td class="num c">${fP1(d.pctN)}</td><td class="num c">${fP1(d.pctS)}</td><td class="num c">${fP1(d.pctT)}</td><td class="mono c" style="font-size:12px;">${esc(d.ncs.join(' · '))}</td><td>${esc(d.producto)}</td><td class="c">${chip(d.grupo, d.etiqueta || GRUPO[d.grupo])}</td></tr>`).join('')
      || '<tr><td class="muted" colspan="12">Sin documentos para este filtro.</td></tr>';
    const s = lista.reduce((a, d) => ({ v: a.v + d.ventaIgv, N: a.N + d.N, S: a.S + d.S }), { v: 0, N: 0, S: 0 });
    tb.querySelector('tfoot').innerHTML = `<tr><td></td><td colspan="4">Total (${fN(lista.length)} facturas) · Nestlé ${f2(s.N)} · ST ${f2(s.S)}</td><td class="num">${f2(s.v)}</td><td colspan="6"></td></tr>`;
    const bar = $('ncSelBar'); const n = st.doc.sel.size;
    bar.style.display = n ? '' : 'none';
    if (n) bar.innerHTML = `<span><b>${n}</b> factura${n > 1 ? 's' : ''} seleccionada${n > 1 ? 's' : ''}</span>
      <span style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;"><input type="text" id="ncEtiqueta" class="input-select" placeholder="Etiqueta (ej. Sagastegui julio)" style="width:220px;">
      <button class="v-meta-save" id="ncMarcarAdi">Marcar como Adicional</button><button class="export-btn" id="ncMarcarReg">Marcar como Regular</button>
      <button class="export-btn" id="ncSelCliente">Todas las de este cliente</button><button class="clear-btn" id="ncSelLimpiar">Quitar selección</button></span>`;
  }

  // ---------------- Guardar marcas y ajustes ----------------
  async function guardarAjustes(msg) {
    const s = await sesion(); const ahora = new Date();
    D.ajustes.actualizado = { usuario: s?.user?.email || '', fecha: ahora.toISOString() };
    overlay(true, msg || 'Guardando...');
    try { await guardarJson(CFG.AJUSTES(D.ym), D.ajustes); }
    catch (e) { console.error(e); aviso(`No se pudo guardar: ${esc(e.message)}. Revisa la política de escritura de Overrides/ en ventas-data.`, true); }
    finally { overlay(false); }
    calcular(); renderTodo();
  }
  async function firma() { const s = await sesion(); const n = nombreDesdeCorreo(s?.user?.email || ''); return { usuario: n.nombre, fechaTexto: new Date().toLocaleString('es-PE', { dateStyle: 'short', timeStyle: 'short' }) }; }

  async function marcarFacturas(tipo) {
    const etq = tipo === 'adicional' ? ($('ncEtiqueta')?.value || '').trim() : '';
    const fm = await firma();
    R.documentos.filter(d => st.doc.sel.has(d.factura)).forEach(d => d.ncs.forEach(nc => {
      if (tipo === 'regular') delete D.ajustes.marcas[nc];
      else D.ajustes.marcas[nc] = { tipo: 'adicional', etiqueta: etq || 'Adicional', ...fm };
    }));
    st.doc.sel.clear();
    await guardarAjustes('Guardando marcas...');
  }

  // ---------------- Subir archivo de Alex ----------------
  async function subirAlex(file) {
    overlay(true, 'Leyendo archivo de Alex...');
    let buf, filas;
    try { buf = await file.arrayBuffer(); filas = C.parseAlex(hojasAlex(buf)); }
    catch (e) { overlay(false); console.error(e); aviso(`No se pudo leer ${esc(file.name)}: ${esc(e.message)}`, true); return; }
    if (!filas.length) { overlay(false); aviso(`${esc(file.name)} no tiene hojas con el formato de Alex (encabezado "CodClie" en la columna A).`, true); return; }
    const meses = {}; filas.forEach(f => { const m = f.fecha.slice(0, 7).replace('/', '-'); meses[m] = (meses[m] || 0) + 1; });
    const ym = Object.entries(meses).sort((a, b) => b[1] - a[1])[0][0];
    try { overlay(true, 'Guardando archivo de Alex...'); await subir(CFG.ALEX(ym), new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); }
    catch (e) { overlay(false); aviso(`No se pudo guardar el archivo de Alex: ${esc(e.message)}`, true); return; }
    overlay(false);
    if (ym !== D.ym) { st.modo = 'mes'; st.mesSel = ym; $('ncMesSel').value = ym; $('ncModoMes').click(); await aplicarCorte({ forzar: true }); }
    else { D.alex = filas; calcular(); renderTodo(); }
    aviso(`Archivo de Alex guardado para ${MES_NOMBRE[+ym.slice(5) - 1]} ${ym.slice(0, 4)}: ${fN(filas.length)} facturas, del ${fD(filas.map(f => f.fecha).sort()[0])} al ${fD(filas.map(f => f.fecha).sort().slice(-1)[0])}.`);
    setTab('concil');
  }

  // ---------------- Exportar ----------------
  async function xlsx(filename, hojas) {
    const wb = new ExcelJS.Workbook(); wb.creator = 'Panel de Gestión · Saga Trans Confitería';
    const NAVY = 'FF16356B', BRAND = 'FF2E6FB7', CREAM = 'FFF5F3ED', STRIPE = 'FFFBF9F4', BORDER = { style: 'thin', color: { argb: 'FFE4E0D4' } };
    hojas.forEach(h => {
      const ws = wb.addWorksheet(h.name.slice(0, 31), { views: [{ state: 'frozen', ySplit: 3 }] });
      const nc = h.columns.length;
      ws.mergeCells(1, 1, 1, nc); const t = ws.getCell(1, 1); t.value = h.title; t.font = { name: 'Arial', bold: true, size: 13, color: { argb: 'FFFFFFFF' } }; t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; ws.getRow(1).height = 24;
      ws.mergeCells(2, 1, 2, nc); const s2 = ws.getCell(2, 1); s2.value = h.subtitle || ''; s2.font = { name: 'Arial', italic: true, size: 10, color: { argb: 'FF6B675F' } };
      const hr = ws.getRow(3);
      h.columns.forEach((c, i) => { const cell = hr.getCell(i + 1); cell.value = c.header; cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BRAND } }; cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; ws.getColumn(i + 1).width = c.width || 14; });
      const put = (vals, r, tot) => {
        const row = ws.getRow(r);
        h.columns.forEach((c, i) => {
          const cell = row.getCell(i + 1); const v = vals[i];
          cell.value = c.code ? (v == null ? '' : String(v)) : (v === null || v === undefined ? '' : v);
          if (c.code) cell.numFmt = '@'; else if (c.fmt) cell.numFmt = c.fmt;
          cell.font = { name: 'Arial', bold: !!tot, color: { argb: tot ? NAVY : 'FF211F1A' } };
          cell.alignment = { horizontal: c.align || (typeof v === 'number' ? 'right' : 'left') };
          cell.border = tot ? { top: { style: 'double', color: { argb: NAVY } } } : { bottom: BORDER };
          if (tot) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CREAM } };
          else if ((r % 2) === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: STRIPE } };
        });
      };
      h.rows.forEach((v, i) => put(v, 4 + i, false));
      if (h.totals) put(h.totals, 4 + h.rows.length, true);
    });
    const buf = await wb.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = Object.assign(document.createElement('a'), { href: url, download: filename }); document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 3000);
  }
  const M2 = '#,##0.00', PCT = '0.00%';
  const sub = () => `Corte al ${largo(st.corte)}`;
  function hojaResumen() {
    const t = R.tot;
    return { name: 'Resumen', title: `NOTAS DE CRÉDITO SAGA TRANS · Resumen de descuentos`, subtitle: sub(), columns: [{ header: 'Concepto', width: 34 }, { header: 'Nestlé', width: 16, fmt: M2 }, { header: 'ST', width: 16, fmt: M2 }, { header: 'Total', width: 16, fmt: M2 }],
      rows: [['Descuentos confitería c/IGV', t.N, t.S, t.total], ['Descuentos confitería s/IGV', t.Ns, t.Ss, t.totals], ['Participación', t.shareN, t.shareS, 1], ['Rechazos c/IGV', t.rechN, t.rechS, t.rechN + t.rechS], ['Ajustes manuales c/IGV', t.ajN, t.ajS, t.ajN + t.ajS], ['Adicionales c/IGV', '', '', t.adicional], ['Panetón c/IGV', '', '', t.paneton], ['Sell out s/IGV (Chocolate + Galleta)', '', '', t.sellout], ['%DCTOS ST sobre sell out s/IGV', '', t.pctST, ''], ['%DCTOS ST sobre sell out c/IGV', '', t.pctSTc, '']]
        .map(r => (r[0] === 'Participación' || r[0].startsWith('%DCTOS') ? r.map((v, i) => (i && typeof v === 'number' ? fP(v) : v)) : r)) };
  }
  function hojaDiario() {
    const sN = R.series.filter(s => s.charAt(1) === 'N'), sS = R.series.filter(s => s.charAt(1) === 'S'); const k = st.igv;
    const cols = [{ header: 'Fecha', width: 12, align: 'center' }, ...sN.map(s => ({ header: s, fmt: M2 })), { header: 'Nestlé', fmt: M2 }, ...sS.map(s => ({ header: s, fmt: M2 })), { header: 'ST', fmt: M2 }, { header: 'Total', fmt: M2 }];
    const rows = R.diario.map(d => { const n = sN.reduce((a, s) => a + (d[k][s] || 0), 0), s2 = sS.reduce((a, s) => a + (d[k][s] || 0), 0); return [fD(d.fecha), ...sN.map(s => d[k][s] || 0), n, ...sS.map(s => d[k][s] || 0), s2, n + s2]; });
    const totals = ['Total', ...cols.slice(1).map((_, i) => rows.reduce((a, r) => a + (r[i + 1] || 0), 0))];
    return { name: 'Diario por serie', title: `NOTAS DE CRÉDITO · Diario por serie (${k === 'c' ? 'con' : 'sin'} IGV)`, subtitle: sub(), columns: cols, rows, totals };
  }
  function hojaAnuladas() {
    return { name: 'Anuladas', title: 'NOTAS DE CRÉDITO · Anuladas', subtitle: sub(), columns: [{ header: 'Nota de crédito', width: 18, code: true }, { header: 'Emisión', width: 12 }, { header: 'Anulada', width: 12 }, { header: 'Cliente', code: true }, { header: 'Asume' }, { header: 'Monto', fmt: M2 }, { header: 'Origen' }, { header: 'Clasificación' }],
      rows: R.anuladas.map(a => [a.doc, fD(a.fecha), a.fecanu ? fD(a.fecanu) : '', a.codcli, a.cargo === 'N' ? 'Nestlé' : 'ST', a.montoAnulada, a.origen, CLASE[a.clase]]) };
  }
  function hojaConcil(soloPend) {
    const c = R.concil; if (!c) return null;
    const filas = c.filas.filter(f => !soloPend || f.estado !== 'coincide');
    return { name: 'Conciliación Alex', title: 'NOTAS DE CRÉDITO · Conciliación con Alex', subtitle: `${sub()} · ${soloPend ? 'solo facturas por revisar' : 'todas las facturas'}`, columns: [{ header: 'Fecha', width: 12 }, { header: 'Factura', width: 18, code: true }, { header: 'Cliente', code: true }, { header: 'Razón social', width: 32 }, { header: 'Alex Nestlé', fmt: M2 }, { header: 'Sistema Nestlé', fmt: M2 }, { header: 'Dif. Nestlé', fmt: M2 }, { header: 'Alex ST', fmt: M2 }, { header: 'Sistema ST', fmt: M2 }, { header: 'Dif. ST', fmt: M2 }, { header: 'Resultado', width: 16 }, { header: 'Notas de crédito', width: 34 }],
      rows: filas.map(f => [fD(f.fecha), f.doc, f.cli, f.razon, f.alexN, f.sisN, f.dN, f.alexS, f.sisS, f.dS, ESTADO[f.estado], f.ncs.join(' · ')]) };
  }
  function hojaProd() { return { name: 'Por producto', title: 'NOTAS DE CRÉDITO · Descuentos por producto', subtitle: sub(), columns: [{ header: 'Producto / promoción', width: 28 }, { header: 'Facturas' }, { header: 'Nestlé', fmt: M2 }, { header: 'ST', fmt: M2 }, { header: 'Total', fmt: M2 }], rows: R.productos.map(p => [p.producto, p.facturas, p.N, p.S, p.total]) }; }
  function hojaDocs() {
    return { name: 'Documentos', title: 'NOTAS DE CRÉDITO · Documentos', subtitle: sub(), columns: [{ header: 'Fecha', width: 12 }, { header: 'Factura', width: 18, code: true }, { header: 'Cliente', code: true }, { header: 'Razón social', width: 32 }, { header: 'Monto c/IGV', fmt: M2 }, { header: 'Nestlé c/IGV', fmt: M2 }, { header: 'ST c/IGV', fmt: M2 }, { header: '% Nestlé', fmt: PCT }, { header: '% ST', fmt: PCT }, { header: '% Total', fmt: PCT }, { header: 'Notas de crédito', width: 34 }, { header: 'Producto', width: 16 }, { header: 'Grupo', width: 14 }],
      rows: docsFiltrados().map(d => [fD(d.fecha), d.factura, d.cli, d.cliente, d.ventaIgv, d.N, d.S, d.pctN, d.pctS, d.pctT, d.ncs.join(' · '), d.producto, d.etiqueta || GRUPO[d.grupo]]) };
  }
  const fileTag = () => toIso(st.corte);
  function reporteXlsx() { xlsx(`reporte_nc_${fileTag()}.xlsx`, [hojaResumen(), hojaDiario(), hojaAnuladas(), hojaConcil(true), hojaProd()].filter(Boolean)); }
  function reportePdf() {
    const t = R.tot; const img = CHART ? CHART.toBase64Image('image/png', 1) : '';
    const fila = (l, a, b, c) => `<tr><td>${l}</td><td class="num">${a}</td><td class="num">${b}</td><td class="num">${c}</td></tr>`;
    const c = R.concil;
    $('ncReporte').innerHTML = `<div class="print-title" style="display:block"><div class="print-title-main">Reporte diario de notas de crédito</div><div class="print-title-sub">Saga Trans Confitería · Corte al ${largo(st.corte)}</div></div>
      <table><thead><tr><th>Concepto</th><th>Nestlé</th><th>ST</th><th>Total</th></tr></thead><tbody>
      ${fila('Descuentos c/IGV', f2(t.N), f2(t.S), f2(t.total))}${fila('Descuentos s/IGV', f2(t.Ns), f2(t.Ss), f2(t.totals))}${fila('Participación', fP1(t.shareN), fP1(t.shareS), '100%')}
      ${fila('Rechazos c/IGV', f2(t.rechN), f2(t.rechS), f2(t.rechN + t.rechS))}${fila('Adicionales / Panetón', '', '', `${f2(t.adicional)} / ${f2(t.paneton)}`)}
      ${fila('Sell out s/IGV · %DCTOS ST', '', `${fP(t.pctST)} (c/IGV ${fP(t.pctSTc)})`, f2(t.sellout))}</tbody></table>
      ${img ? `<h2>% Descuentos ST</h2><img src="${img}" alt="Curva % descuentos ST">` : ''}
      ${c ? `<h2>Conciliación con Alex</h2><table><tbody>${Object.entries(ESTADO).map(([k, v]) => `<tr><td>${v}</td><td class="num">${fN(c.cuenta[k] || 0)}</td></tr>`).join('')}</tbody></table>` : ''}
      <h2>Diario por serie (con IGV)</h2>${$('ncDiarioTable').outerHTML}`;
    const node = $('ncReporte'); node.classList.add('print-target'); window.print(); setTimeout(() => node.classList.remove('print-target'), 800);
  }

  // ---------------- Pestañas, módulo y eventos ----------------
  function setTab(tab) {
    st.tab = tab;
    document.querySelectorAll('#ncTabs .tab-btn').forEach(b => b.classList.toggle('tab-active', b.dataset.nctab === tab));
    ['resumen', 'diario', 'anuladas', 'concil', 'producto', 'ajustes', 'documentos'].forEach(t => { $('ncpanel-' + t).style.display = t === tab ? '' : 'none'; });
  }
  async function abrir() {
    activarModulo('NC');
    const s = await sesion(); const q = nombreDesdeCorreo(s?.user?.email || '');
    $('ncUserEmail').textContent = q.nombre; $('ncUserEmail').title = s?.user?.email || ''; $('ncUserAvatar').textContent = q.iniciales;
    if (!st.booted) { st.booted = true; try { await aplicarCorte(); } catch (e) { console.error(e); st.booted = false; aviso(`No se pudo cargar el módulo: ${esc(e.message)}`, true); } }
  }

  function wire() {
    $('railNC').addEventListener('click', abrir);
    $('ncLogoutBtn').addEventListener('click', async () => { try { await supabaseClient.auth.signOut(); } catch (e) { console.error(e); } });
    const hoy = new Date(); st.mesSel = `${hoy.getFullYear()}-${pad(hoy.getMonth() + 1)}`; st.diaSel = `${st.mesSel}-${pad(hoy.getDate())}`;
    const sel = $('ncMesSel');
    for (let i = 11; i >= 0; i--) { const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1); const v = `${d.getFullYear()}-${pad(d.getMonth() + 1)}`; sel.appendChild(new Option(`${MES_ABBR[d.getMonth()]} ${d.getFullYear()}`, v, false, v === st.mesSel)); }
    sel.addEventListener('change', () => { st.mesSel = sel.value; aplicarCorte(); });
    $('ncDiaSel').addEventListener('change', e => { if (e.target.value) { st.diaSel = e.target.value; aplicarCorte(); } });
    $('ncModoMes').addEventListener('click', () => { st.modo = 'mes'; $('ncModoMes').classList.add('tab-active'); $('ncModoDia').classList.remove('tab-active'); sel.style.display = ''; $('ncDiaSel').style.display = 'none'; if (D) aplicarCorte(); });
    $('ncModoDia').addEventListener('click', () => { st.modo = 'dia'; $('ncModoDia').classList.add('tab-active'); $('ncModoMes').classList.remove('tab-active'); if (st.corte) st.diaSel = toIso(st.corte); $('ncDiaSel').value = st.diaSel; sel.style.display = 'none'; $('ncDiaSel').style.display = ''; if (D) aplicarCorte(); });
    document.querySelectorAll('#ncTabs .tab-btn').forEach(b => b.addEventListener('click', () => setTab(b.dataset.nctab)));
    $('ncBtnRecargar').addEventListener('click', () => D && aplicarCorte({ forzar: true }));
    const pedirAlex = () => $('ncAlexFile').click();
    $('ncBtnAlex').addEventListener('click', pedirAlex); $('ncBtnAlex2').addEventListener('click', pedirAlex);
    $('ncAlexFile').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) subirAlex(f); });
    $('ncReportePdf').addEventListener('click', () => R && reportePdf());
    $('ncReporteXlsx').addEventListener('click', () => R && reporteXlsx());
    // Diario
    $('ncIgvSel').addEventListener('click', e => { const b = e.target.closest('[data-igv]'); if (b) { st.igv = b.dataset.igv; renderDiario(); } });
    $('ncExportDiario').addEventListener('click', () => xlsx(`nc_diario_${fileTag()}.xlsx`, [hojaDiario()]));
    // Anuladas
    $('ncAnuFiltro').addEventListener('click', e => { const b = e.target.closest('[data-f]'); if (b) { st.anu.filtro = b.dataset.f; st.anu.page = 0; renderAnuladas(); } });
    $('ncAnuTable').addEventListener('change', e => {
      const doc = e.target.dataset.doc; if (!doc) return;
      const c = (st.anu.cambios[doc] = st.anu.cambios[doc] || {});
      if (e.target.tagName === 'SELECT') c.tipo = e.target.value; else c.monto = parseFloat(e.target.value) || 0;
      renderAnuladas();
    });
    $('ncAnuGuardar').addEventListener('click', async () => {
      const fm = await firma();
      Object.entries(st.anu.cambios).forEach(([doc, c]) => {
        const a = R.anuladas.find(x => x.doc === doc); const tipo = c.tipo || a.clase;
        D.ajustes.marcas[doc] = { tipo, monto: tipo === 'error' ? 0 : (c.monto !== undefined ? c.monto : a.montoAnulada), ...fm };
      });
      st.anu.cambios = {}; await guardarAjustes('Guardando clasificación...');
    });
    $('ncAnuPrev').addEventListener('click', () => { st.anu.page = Math.max(0, st.anu.page - 1); renderAnuladas(); });
    $('ncAnuNext').addEventListener('click', () => { st.anu.page++; renderAnuladas(); });
    // Conciliación
    $('ncConKpis').addEventListener('click', e => { const k = e.target.closest('[data-cf]'); if (k) { st.con.filtro = st.con.filtro === k.dataset.cf ? 'pend' : k.dataset.cf; st.con.page = 0; renderConcil(); } });
    $('ncConPrev').addEventListener('click', () => { st.con.page = Math.max(0, st.con.page - 1); renderConcil(); });
    $('ncConNext').addEventListener('click', () => { st.con.page++; renderConcil(); });
    $('ncExportConcil').addEventListener('click', () => { const h = hojaConcil(false); if (h) xlsx(`nc_conciliacion_alex_${fileTag()}.xlsx`, [h]); });
    $('ncSrAlexTable').addEventListener('change', e => {
      const v = e.target.value === '' ? '' : parseFloat(e.target.value);
      if (e.target.dataset.correo) st.correo[e.target.dataset.correo] = v;
      else if (e.target.dataset.correos) st.correoS[e.target.dataset.correos] = v;
      else return;
      renderConcil();
    });
    $('ncCorreoGuardar').addEventListener('click', async () => {
      D.ajustes.correoST = D.ajustes.correoST || {};
      Object.assign(D.ajustes.correo, st.correo); Object.assign(D.ajustes.correoST, st.correoS);
      st.correo = {}; st.correoS = {}; await guardarAjustes('Guardando montos del correo...');
    });
    // Producto
    $('ncExportProd').addEventListener('click', () => xlsx(`nc_productos_${fileTag()}.xlsx`, [hojaProd()]));
    // Ajustes
    $('ncAjAgregar').addEventListener('click', async () => {
      const concepto = $('ncAjConcepto').value.trim(), monto = parseFloat($('ncAjMonto').value);
      if (!concepto || !isFinite(monto)) { aviso('Escribe el concepto y el monto del ajuste.', true); return; }
      D.ajustes.manuales.push({ concepto, cargo: $('ncAjCargo').value, monto, ...(await firma()) });
      $('ncAjConcepto').value = ''; $('ncAjMonto').value = '';
      await guardarAjustes('Guardando ajuste...');
    });
    $('ncAjTable').addEventListener('click', async e => {
      const q = e.target.closest('[data-quitar-aj]'), qe = e.target.closest('[data-quitar-etq]');
      if (q) { D.ajustes.manuales.splice(+q.dataset.quitarAj, 1); await guardarAjustes(); }
      if (qe) { Object.keys(D.ajustes.marcas).forEach(k => { const m = D.ajustes.marcas[k]; if (m.tipo === 'adicional' && (m.etiqueta || 'Adicional') === qe.dataset.quitarEtq) delete D.ajustes.marcas[k]; }); await guardarAjustes(); }
    });
    // Documentos
    $('ncDocBuscar').addEventListener('input', e => { st.doc.buscar = e.target.value; st.doc.page = 0; renderDocs(); });
    $('ncDocGrupo').addEventListener('change', e => { st.doc.grupo = e.target.value; st.doc.page = 0; renderDocs(); });
    $('ncDocLimpiar').addEventListener('click', () => { st.doc.buscar = ''; st.doc.grupo = 'all'; st.doc.page = 0; $('ncDocBuscar').value = ''; $('ncDocGrupo').value = 'all'; renderDocs(); });
    $('ncDocPrev').addEventListener('click', () => { st.doc.page = Math.max(0, st.doc.page - 1); renderDocs(); });
    $('ncDocNext').addEventListener('click', () => { st.doc.page++; renderDocs(); });
    $('ncDocTable').addEventListener('change', e => {
      if (e.target.id === 'ncSelPag') { const lista = docsFiltrados().slice(st.doc.page * CFG.POR_PAG, st.doc.page * CFG.POR_PAG + CFG.POR_PAG); lista.forEach(d => (e.target.checked ? st.doc.sel.add(d.factura) : st.doc.sel.delete(d.factura))); renderDocs(); return; }
      const f = e.target.dataset.sel; if (!f) return; e.target.checked ? st.doc.sel.add(f) : st.doc.sel.delete(f); renderDocs();
    });
    $('ncSelBar').addEventListener('click', e => {
      if (e.target.id === 'ncMarcarAdi') marcarFacturas('adicional');
      if (e.target.id === 'ncMarcarReg') marcarFacturas('regular');
      if (e.target.id === 'ncSelLimpiar') { st.doc.sel.clear(); renderDocs(); }
      if (e.target.id === 'ncSelCliente') { const clis = new Set(R.documentos.filter(d => st.doc.sel.has(d.factura)).map(d => d.cli)); R.documentos.filter(d => clis.has(d.cli)).forEach(d => st.doc.sel.add(d.factura)); renderDocs(); }
    });
    $('ncExportDocs').addEventListener('click', () => xlsx(`nc_documentos_${fileTag()}.xlsx`, [hojaDocs()]));
  }
  wire();
})();
