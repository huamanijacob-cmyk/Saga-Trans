/* ============================================================
   Panel de Rechazos — lógica de la app
   Lee los Excel directo de Supabase Storage (sin tablas/SQL),
   igual que el patrón de tu otro proyecto: SheetJS en el navegador.
   ============================================================ */

console.log('Panel de Rechazos — app.js versión 2 (con login endurecido)');

// Bloquea el bfcache: si el navegador restaura una foto congelada de la
// página (Atrás/Adelante después de cerrar sesión), fuerza una recarga real
// en vez de mostrar el panel tal como estaba en el momento de salir.
window.addEventListener('pageshow', (e) => {
  if (e.persisted) location.reload();
});

// Sesión SOLO en memoria (nada en localStorage/sessionStorage): al recargar
// la página, o si alguien cierra la pestaña, hay que loguearse de nuevo.
// Elegido así porque este panel muestra montos, clientes y vendedores reales.
const memoryAuthStorage = (() => {
  const store = new Map();
  return {
    getItem: (key) => (store.has(key) ? store.get(key) : null),
    setItem: (key, value) => { store.set(key, value); },
    removeItem: (key) => { store.delete(key); },
  };
})();

const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: { storage: memoryAuthStorage, persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
});

const MES_ABBR = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Set','Oct','Nov','Dic'];

const state = {
  tab: 'general',
  modo: 'mes',            // 'mes' | 'dia'
  mesSel: null,           // 'YYYY-MM'
  diaSel: null,           // 'YYYY-MM-DD'
  cond: { search: '', onlyAlerts: false, umbral: 2, sortKey: 'pct', sortDir: 'desc' },
  vend: { search: '' },
  doc:  { search: '', vendor: 'all', motivo: 'all', montoMin: '', montoMax: '', from: null, to: null, page: 0 },
};

let VENDOR_NAMES = {};                              // código -> nombre (se llena desde Ventas)
let CORTE_DATA = { transportistas: [], ventas: [], nc: [] };
const FILE_CACHE = {};                              // 'Transportistas/Transportistas_2026-09.xlsx' -> filas ya parseadas

/* ---------------- utilidades ---------------- */

function fmtMoney(n) {
  n = Number(n) || 0;
  return 'S/ ' + n.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtPct(n) {
  n = Number(n) || 0;
  return (n * 100).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';
}
function pctLevel(pct, umbral) {
  const p = pct * 100;
  if (p >= umbral) return 'pct-high';
  if (p >= umbral / 2) return 'pct-mid';
  return 'pct-low';
}
function pad(n) { return n.toString().padStart(2, '0'); }
function isoDate(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function toISO(v) {
  // Normaliza fecha de Excel (Date real gracias a cellDates:true, número serial, o string) a 'YYYY-MM-DD'.
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : isoDate(v);
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    return d ? `${d.y}-${pad(d.m)}-${pad(d.d)}` : null;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); // dd/mm/aaaa
  if (m) return `${m[3]}-${pad(+m[2])}-${pad(+m[1])}`;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : isoDate(d);
}
function fmtFecha(v) {
  const iso = toISO(v);
  if (!iso) return '';
  const d = new Date(iso + 'T00:00:00');
  return d.toLocaleDateString('es-PE', { day: '2-digit', month: 'short', year: 'numeric' });
}
function lastNMonths(n, ref) {
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(ref.getFullYear(), ref.getMonth() - i, 1);
    out.push({ year: d.getFullYear(), month: d.getMonth() + 1 });
  }
  return out;
}
function monthBounds(periodo) {
  const [y, m] = periodo.split('-').map(Number);
  const lastDay = new Date(y, m, 0).getDate();
  return { from: `${periodo}-01`, to: `${periodo}-${pad(lastDay)}` };
}
function monthsBetween(from, to) {
  const out = [];
  let [y, m] = from.slice(0, 7).split('-').map(Number);
  const [ye, me] = to.slice(0, 7).split('-').map(Number);
  while (y < ye || (y === ye && m <= me)) {
    out.push(`${y}-${pad(m)}`);
    m++; if (m > 12) { m = 1; y++; }
  }
  return out;
}
function csvEscape(v) {
  const s = String(v == null ? '' : v);
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
function downloadBlob(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function el(tag, cls, html) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (html != null) e.innerHTML = html;
  return e;
}

/* ---------------- lectura de Excel desde Storage ---------------- */

async function loadModuleFile(folder, prefix, periodo) {
  const path = `${folder}/${prefix}_${periodo}.xlsx`;
  if (FILE_CACHE[path]) return FILE_CACHE[path];

  const { data, error } = await supabase.storage.from(STORAGE_BUCKET).download(path);
  if (error) {
    // Es normal no tener todavía el archivo de un mes (p.ej. futuro o no subido aún).
    console.warn(`No se encontró "${path}" en el bucket "${STORAGE_BUCKET}":`, error.message);
    FILE_CACHE[path] = [];
    return [];
  }
  const buf = await data.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, cellDates: true });
  FILE_CACHE[path] = rows;
  return rows;
}

/* ---------------- corte (mes / día) ---------------- */

function getCorteRange() {
  if (state.modo === 'dia') return { from: state.diaSel, to: state.diaSel };
  return monthBounds(state.mesSel);
}
function getCorteLabel() {
  if (state.modo === 'dia') {
    const d = new Date(state.diaSel + 'T00:00:00');
    return d.toLocaleDateString('es-PE', { day: '2-digit', month: 'long', year: 'numeric' });
  }
  const [y, m] = state.mesSel.split('-').map(Number);
  const esMesActual = state.mesSel === isoDate(new Date()).slice(0, 7);
  const lastDay = new Date(y, m, 0).getDate();
  const dt = new Date(y, m - 1, esMesActual ? new Date().getDate() : lastDay);
  return dt.toLocaleDateString('es-PE', { day: '2-digit', month: 'long', year: 'numeric' }) + (esMesActual ? ' (a la fecha)' : ' (cierre de mes)');
}

function initCorteControls() {
  const hoy = new Date();
  state.mesSel = isoDate(hoy).slice(0, 7);
  state.diaSel = isoDate(hoy);

  const sel = document.getElementById('mesSel');
  lastNMonths(12, hoy).forEach(({ year, month }) => {
    const val = `${year}-${pad(month)}`;
    const opt = el('option', null, `${MES_ABBR[month - 1]} ${year}`);
    opt.value = val;
    if (val === state.mesSel) opt.selected = true;
    sel.appendChild(opt);
  });
  sel.addEventListener('change', () => { state.mesSel = sel.value; onCorteChanged(); });

  const dia = document.getElementById('diaSel');
  dia.value = state.diaSel;
  dia.addEventListener('change', () => { state.diaSel = dia.value; onCorteChanged(); });

  document.getElementById('btnModoMes').addEventListener('click', () => {
    state.modo = 'mes';
    document.getElementById('btnModoMes').classList.add('tab-active');
    document.getElementById('btnModoDia').classList.remove('tab-active');
    sel.style.display = ''; dia.style.display = 'none';
    onCorteChanged();
  });
  document.getElementById('btnModoDia').addEventListener('click', () => {
    state.modo = 'dia';
    document.getElementById('btnModoDia').classList.add('tab-active');
    document.getElementById('btnModoMes').classList.remove('tab-active');
    sel.style.display = 'none'; dia.style.display = '';
    onCorteChanged();
  });
}

async function onCorteChanged() {
  const range = getCorteRange();
  state.doc.from = range.from;
  state.doc.to = range.to;
  document.getElementById('docFechaFrom').value = range.from;
  document.getElementById('docFechaTo').value = range.to;

  await loadCorteData();
  await renderDocumentos();
  renderAll();
}

/* ---------------- carga de datos del corte activo ---------------- */

async function loadCorteData() {
  const periodo = state.modo === 'dia' ? state.diaSel.slice(0, 7) : state.mesSel;
  const [t, v, n] = await Promise.all([
    loadModuleFile('Transportistas', 'Transportistas', periodo),
    loadModuleFile('Ventas', 'Ventas', periodo),
    loadModuleFile('ContabilidadNC', 'ContabilidadNC', periodo),
  ]);

  let transportistas = t, ventas = v, nc = n;
  if (state.modo === 'dia') {
    const d = state.diaSel;
    transportistas = transportistas.filter(r => toISO(r.fecemi) === d);
    ventas = ventas.filter(r => toISO(r.fecha) === d);
    nc = nc.filter(r => toISO(r.emision) === d);
  }
  CORTE_DATA = { transportistas, ventas, nc };

  // Nombres de vendedor: se arman con lo que haya en Ventas de este corte.
  ventas.forEach(r => { if (r.vendedor && r.nombrevendedor) VENDOR_NAMES[r.vendedor] = r.nombrevendedor; });
}

async function loadDocumentosData() {
  if (!state.doc.from || !state.doc.to) return [];
  const periodos = monthsBetween(state.doc.from, state.doc.to);
  const arrays = await Promise.all(periodos.map(p => loadModuleFile('ContabilidadNC', 'ContabilidadNC', p)));
  const all = arrays.flat();
  return all.filter(r => {
    const iso = toISO(r.emision);
    return iso && iso >= state.doc.from && iso <= state.doc.to;
  });
}

/* ---------------- Vista general (cruce Chofer x Vendedor) ---------------- */

function renderGeneral() {
  document.querySelectorAll('.corteLabelInline').forEach(e => e.textContent = getCorteLabel());

  const docToChofer = {};
  CORTE_DATA.transportistas.forEach(r => {
    if (r.Documento) docToChofer[r.Documento] = r.nomcho || r.codcho || 'Sin identificar';
  });

  const cross = [];
  let matched = 0;
  CORTE_DATA.nc.forEach(r => {
    const ref = r['Factura/Boleta'];
    if (!ref) return;
    const chofer = docToChofer[ref];
    if (!chofer) return;
    matched++;
    cross.push({ chofer, vendedor: r.vendedor || '—', monto: Number(r.monimp) || 0 });
  });

  const kpis = document.getElementById('cross-kpis');
  kpis.innerHTML = '';
  kpis.appendChild(kpiCard('navy', 'Documentos cruzados', `${matched} de ${CORTE_DATA.nc.length}`));
  kpis.appendChild(kpiCard('teal', 'Choferes involucrados', new Set(cross.map(c => c.chofer)).size));
  kpis.appendChild(kpiCard('red', 'Monto cruzado', fmtMoney(cross.reduce((s, c) => s + c.monto, 0))));

  const vendCols = Array.from(new Set(cross.map(c => c.vendedor))).sort();
  const choferes = Array.from(new Set(cross.map(c => c.chofer))).sort();
  const matrix = {};
  choferes.forEach(c => matrix[c] = {});
  cross.forEach(c => { matrix[c.chofer][c.vendedor] = (matrix[c.chofer][c.vendedor] || 0) + c.monto; });

  const table = document.getElementById('cross-table');
  const thead = table.querySelector('thead'), tbody = table.querySelector('tbody'), tfoot = table.querySelector('tfoot');
  thead.innerHTML = ''; tbody.innerHTML = ''; tfoot.innerHTML = '';

  if (!choferes.length) {
    document.getElementById('cross-note').textContent = 'No hay coincidencias todavía entre las notas de crédito y los documentos de Transportistas en este periodo.';
    return;
  }

  const trTop = el('tr'); trTop.style.background = '#F5F3ED';
  trTop.appendChild(el('th'));
  const thVend = el('th', null, 'Vendedores'); thVend.colSpan = vendCols.length; thVend.style.textAlign = 'center';
  trTop.appendChild(thVend);
  const thTotal = el('th', null, 'Total<br>general'); thTotal.rowSpan = 2; thTotal.style.textAlign = 'center';
  trTop.appendChild(thTotal);
  thead.appendChild(trTop);

  const trHead = el('tr'); trHead.style.background = '#F5F3ED';
  trHead.appendChild(el('th', null, 'Chofer'));
  vendCols.forEach(v => {
    const th = el('th', null, VENDOR_NAMES[v] ? VENDOR_NAMES[v].toUpperCase() : ('VEND. ' + v));
    th.style.textAlign = 'center';
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  const colTotals = vendCols.map(() => 0);
  let grandTotal = 0;
  choferes.forEach(ch => {
    const tr = el('tr');
    tr.appendChild(el('td', null, ch.toUpperCase()));
    let rowTotal = 0;
    vendCols.forEach((v, i) => {
      const val = matrix[ch][v];
      const td = el('td', 'num rej', val ? val.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '');
      td.style.textAlign = 'center';
      tr.appendChild(td);
      if (val) { rowTotal += val; colTotals[i] += val; }
    });
    grandTotal += rowTotal;
    tr.appendChild(el('td', 'num rej', rowTotal.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })));
    tbody.appendChild(tr);
  });

  const trFoot = el('tr');
  trFoot.appendChild(el('td', null, 'Total general'));
  colTotals.forEach(t => {
    const td = el('td', 'num rej', t.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
    td.style.textAlign = 'center';
    trFoot.appendChild(td);
  });
  trFoot.appendChild(el('td', 'num rej', grandTotal.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })));
  tfoot.appendChild(trFoot);

  document.getElementById('cross-note').textContent =
    'Cruce real: enlaza el N° de Documento Referencia de cada nota de crédito contra el chofer de ese mismo documento en Transportistas.';
}

/* ---------------- Por conductor ---------------- */

function aggregateConductor() {
  const map = {};
  CORTE_DATA.transportistas.forEach(r => {
    const key = r.codcho || '—';
    if (!map[key]) map[key] = { cod: key, nombre: r.nomcho || key, facturado: 0, ventaReal: 0, monto: 0, pedidos: 0 };
    map[key].facturado += Number(r.totdsp) || 0;
    map[key].ventaReal += Number(r.totent) || 0;
    map[key].monto += Number(r.totfal) || 0;
    if ((Number(r.totfal) || 0) > 0) map[key].pedidos += 1;
  });
  return Object.values(map).map(r => ({ ...r, pct: r.facturado ? r.monto / r.facturado : 0 }));
}

function kpiCard(color, label, value, pending) {
  const card = el('div', `kpi kpi-${color}`);
  card.appendChild(el('div', 'kpi-label', label));
  card.appendChild(el('div', pending ? 'kpi-val pending' : 'kpi-val', value));
  return card;
}

function renderConductor() {
  document.querySelectorAll('.corteLabelInline').forEach(e => e.textContent = getCorteLabel());
  const all = aggregateConductor();
  const totals = all.reduce((a, r) => ({ facturado: a.facturado + r.facturado, ventaReal: a.ventaReal + r.ventaReal, monto: a.monto + r.monto, pedidos: a.pedidos + r.pedidos }), { facturado: 0, ventaReal: 0, monto: 0, pedidos: 0 });
  const totalPct = totals.facturado ? totals.monto / totals.facturado : 0;

  const kpis = document.getElementById('cond-kpis');
  kpis.innerHTML = '';
  kpis.appendChild(kpiCard('navy', 'Facturado', fmtMoney(totals.facturado)));
  kpis.appendChild(kpiCard('teal', 'Venta real', fmtMoney(totals.ventaReal)));
  kpis.appendChild(kpiCard('red', 'Monto rechazado', fmtMoney(totals.monto)));
  kpis.appendChild(kpiCard('navy', '% Rechazo global', fmtPct(totalPct)));

  const q = state.cond.search.toLowerCase();
  let rows = all.filter(r => r.nombre.toLowerCase().includes(q) || String(r.cod).toLowerCase().includes(q));
  if (state.cond.onlyAlerts) rows = rows.filter(r => r.pct * 100 >= state.cond.umbral);
  const dir = state.cond.sortDir === 'asc' ? 1 : -1;
  rows.sort((a, b) => state.cond.sortKey === 'nombre' ? a.nombre.localeCompare(b.nombre) * dir : (a[state.cond.sortKey] - b[state.cond.sortKey]) * dir);

  const table = document.getElementById('cond-table');
  const thead = table.querySelector('thead'), tbody = table.querySelector('tbody'), tfoot = table.querySelector('tfoot');
  thead.innerHTML = `<tr>
    <th><button data-sort="nombre">Conductor</button></th>
    <th style="text-align:right"><button data-sort="facturado">Facturado</button></th>
    <th style="text-align:right">Venta real</th>
    <th style="text-align:right"><button data-sort="monto">Monto rech.</button></th>
    <th style="text-align:right"><button data-sort="pedidos">Pedidos</button></th>
    <th><button data-sort="pct">% Rechazo</button></th>
  </tr>`;
  thead.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
    const k = b.dataset.sort;
    state.cond.sortDir = (state.cond.sortKey === k && state.cond.sortDir === 'desc') ? 'asc' : 'desc';
    state.cond.sortKey = k;
    renderConductor();
  }));

  tbody.innerHTML = '';
  rows.forEach(r => {
    const tr = el('tr');
    tr.appendChild(el('td', null, `${r.nombre}<div class="sub">${r.cod}</div>`));
    tr.appendChild(el('td', 'num', fmtMoney(r.facturado)));
    tr.appendChild(el('td', 'num muted', fmtMoney(r.ventaReal)));
    tr.appendChild(el('td', 'num rej', fmtMoney(r.monto)));
    tr.appendChild(el('td', 'num', r.pedidos));
    tr.appendChild(pctCell(r.pct, state.cond.umbral));
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = '';
  const trT = el('tr');
  trT.appendChild(el('td', null, 'Total general'));
  trT.appendChild(el('td', 'num', fmtMoney(totals.facturado)));
  trT.appendChild(el('td', 'num', fmtMoney(totals.ventaReal)));
  trT.appendChild(el('td', 'num rej', fmtMoney(totals.monto)));
  trT.appendChild(el('td', 'num', totals.pedidos));
  trT.appendChild(pctCell(totalPct, state.cond.umbral));
  tfoot.appendChild(trT);

  document.getElementById('cond-note').textContent = `Mostrando ${rows.length} de ${all.length} conductores · datos reales de Transportistas (${getCorteLabel()}).`;
}

function pctCell(pct, umbral) {
  const td = el('td');
  const wrap = el('div', 'pct-cell');
  const track = el('div', 'pct-bar-track');
  const bar = el('div', `pct-bar ${pctLevel(pct, umbral)}`);
  bar.style.width = Math.min(100, Math.round(pct * 100 * 3)) + '%';
  track.appendChild(bar);
  wrap.appendChild(track);
  wrap.appendChild(el('span', 'pct-num', fmtPct(pct)));
  td.appendChild(wrap);
  return td;
}

/* ---------------- Por vendedor ---------------- */

function renderVendedor() {
  document.querySelectorAll('.corteLabelInline').forEach(e => e.textContent = getCorteLabel());

  const ncMap = {};
  CORTE_DATA.nc.forEach(r => {
    const key = r.vendedor || '—';
    if (!ncMap[key]) ncMap[key] = { cod: key, monto: 0, docs: 0 };
    ncMap[key].monto += Number(r.monimp) || 0;
    ncMap[key].docs += 1;
  });

  const ventasByVendor = {};
  CORTE_DATA.ventas.forEach(r => {
    const key = r.vendedor || '—';
    if (!ventasByVendor[key]) ventasByVendor[key] = { facturado: 0, ventaReal: 0 };
    ventasByVendor[key].facturado += Number(r.soles) || 0;
    ventasByVendor[key].ventaReal += Number(r.vtadvo) || 0;
  });
  const hayVentas = CORTE_DATA.ventas.length > 0;

  const montoTotal = Object.values(ncMap).reduce((s, v) => s + v.monto, 0);
  const docsTotal = CORTE_DATA.nc.length;
  const facturadoTotal = Object.values(ventasByVendor).reduce((s, v) => s + v.facturado, 0);
  const ventaRealTotal = Object.values(ventasByVendor).reduce((s, v) => s + v.ventaReal, 0);
  const pctTotal = facturadoTotal ? montoTotal / facturadoTotal : 0;

  const kpis = document.getElementById('vend-kpis');
  kpis.innerHTML = '';
  kpis.appendChild(kpiCard('navy', 'Facturado', hayVentas ? fmtMoney(facturadoTotal) : 'pendiente', !hayVentas));
  kpis.appendChild(kpiCard('teal', 'Venta real', hayVentas ? fmtMoney(ventaRealTotal) : 'pendiente', !hayVentas));
  kpis.appendChild(kpiCard('red', 'Monto rechazado', fmtMoney(montoTotal)));
  kpis.appendChild(kpiCard('navy', '% Rechazo global', hayVentas ? fmtPct(pctTotal) : 'pendiente', !hayVentas));

  const q = state.vend.search.toLowerCase();
  let rows = Object.values(ncMap).filter(v => {
    const nm = (VENDOR_NAMES[v.cod] || '').toLowerCase();
    return !q || String(v.cod).toLowerCase().includes(q) || nm.includes(q);
  }).sort((a, b) => b.monto - a.monto);

  const table = document.getElementById('vend-table');
  const thead = table.querySelector('thead'), tbody = table.querySelector('tbody'), tfoot = table.querySelector('tfoot');
  thead.innerHTML = `<tr>
    <th>Vendedor</th>
    <th style="text-align:right">Facturado</th>
    <th style="text-align:right">Venta real</th>
    <th style="text-align:right">Monto rech.</th>
    <th style="text-align:right">Pedidos</th>
    <th>% Rechazo</th>
    <th style="width:20px"></th>
  </tr>`;

  tbody.innerHTML = '';
  rows.forEach(v => {
    const vv = ventasByVendor[v.cod];
    const facturado = vv ? vv.facturado : null;
    const ventaReal = vv ? vv.ventaReal : null;
    const pct = facturado ? v.monto / facturado : null;

    const tr = el('tr', 'clickable');
    tr.addEventListener('click', () => goToVendorDocs(v.cod));
    const nombre = VENDOR_NAMES[v.cod];
    tr.appendChild(el('td', null, nombre ? `${nombre}<div class="sub">Vendedor ${v.cod}</div>` : `Vendedor ${v.cod}<div class="sub" style="font-style:italic">nombre pendiente</div>`));
    tr.appendChild(el('td', 'num', facturado != null ? fmtMoney(facturado) : '<span class="muted" style="font-style:italic">pendiente</span>'));
    tr.appendChild(el('td', 'num', ventaReal != null ? fmtMoney(ventaReal) : '<span class="muted" style="font-style:italic">pendiente</span>'));
    tr.appendChild(el('td', 'num rej', fmtMoney(v.monto)));
    tr.appendChild(el('td', 'num', v.docs));
    tr.appendChild(pct != null ? pctCell(pct, state.cond.umbral) : el('td', 'muted', '<span style="font-style:italic">pendiente</span>'));
    tr.appendChild(el('td', 'muted', '&rsaquo;'));
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = '';
  const trT = el('tr');
  trT.appendChild(el('td', null, 'Total general'));
  trT.appendChild(el('td', 'num', hayVentas ? fmtMoney(facturadoTotal) : '<span class="muted" style="font-style:italic">pendiente</span>'));
  trT.appendChild(el('td', 'num', hayVentas ? fmtMoney(ventaRealTotal) : '<span class="muted" style="font-style:italic">pendiente</span>'));
  trT.appendChild(el('td', 'num rej', fmtMoney(montoTotal)));
  trT.appendChild(el('td', 'num', docsTotal));
  trT.appendChild(el('td', 'muted', hayVentas ? fmtPct(pctTotal) : '<span style="font-style:italic">pendiente</span>'));
  trT.appendChild(el('td'));
  tfoot.appendChild(trT);

  document.getElementById('vend-note').textContent = `Monto rechazado y Pedidos: reales (Contabilidad NC). Facturado/Venta real: ${hayVentas ? 'reales (Ventas)' : 'pendientes — sin filas de Ventas en este periodo'} · haz clic en un vendedor para ver sus documentos.`;
}

function goToVendorDocs(cod) {
  state.doc.vendor = cod;
  setTab('documentos');
}

/* ---------------- Documentos rechazados ---------------- */

let DOC_ROWS_CACHE = [];

async function renderDocumentos() {
  DOC_ROWS_CACHE = await loadDocumentosData();
  populateDocFilterOptions(DOC_ROWS_CACHE);
  drawDocumentos();
}

function populateDocFilterOptions(rows) {
  const vendorSel = document.getElementById('docVendor');
  const motivoSel = document.getElementById('docMotivo');
  const prevVendor = state.doc.vendor, prevMotivo = state.doc.motivo;

  const vendors = Array.from(new Set(rows.map(r => r.vendedor).filter(Boolean))).sort();
  vendorSel.innerHTML = '';
  vendorSel.appendChild(new Option('Todos', 'all'));
  vendors.forEach(v => vendorSel.appendChild(new Option(VENDOR_NAMES[v] || ('Vendedor ' + v), v)));
  vendorSel.value = vendors.includes(prevVendor) ? prevVendor : 'all';
  state.doc.vendor = vendorSel.value;

  const motivos = Array.from(new Set(rows.map(r => r.motivo).filter(Boolean))).sort();
  motivoSel.innerHTML = '';
  motivoSel.appendChild(new Option('Todos', 'all'));
  motivos.forEach(m => motivoSel.appendChild(new Option(m, m)));
  motivoSel.value = motivos.includes(prevMotivo) ? prevMotivo : 'all';
  state.doc.motivo = motivoSel.value;
}

function filteredDocRows() {
  const q = state.doc.search.toLowerCase();
  const min = state.doc.montoMin === '' ? null : parseFloat(state.doc.montoMin);
  const max = state.doc.montoMax === '' ? null : parseFloat(state.doc.montoMax);
  return DOC_ROWS_CACHE.filter(r => {
    if (q && !((String(r.Documento || '')).toLowerCase().includes(q) || (String(r.razsoc || '')).toLowerCase().includes(q))) return false;
    if (state.doc.vendor !== 'all' && r.vendedor !== state.doc.vendor) return false;
    if (state.doc.motivo !== 'all' && r.motivo !== state.doc.motivo) return false;
    if (min !== null && (Number(r.monimp) || 0) < min) return false;
    if (max !== null && (Number(r.monimp) || 0) > max) return false;
    return true;
  });
}

function drawDocumentos() {
  const rows = filteredDocRows();
  const sum = rows.reduce((s, r) => s + (Number(r.monimp) || 0), 0);
  const avg = rows.length ? sum / rows.length : 0;
  const clientes = new Set(rows.map(r => r.razsoc)).size;

  const kpis = document.getElementById('doc-kpis');
  kpis.innerHTML = '';
  kpis.appendChild(kpiCard('navy', 'Documentos (filtro)', rows.length));
  kpis.appendChild(kpiCard('red', 'Suma total', fmtMoney(sum)));
  kpis.appendChild(kpiCard('teal', 'Promedio por documento', fmtMoney(avg)));
  kpis.appendChild(kpiCard('navy', 'Clientes únicos', clientes));

  document.getElementById('doc-result-label').textContent = `${rows.length} documento(s) encontrado(s) · Contabilidad NC`;

  const pageSize = 8;
  const maxPage = Math.max(0, Math.ceil(rows.length / pageSize) - 1);
  state.doc.page = Math.min(state.doc.page, maxPage);
  const pageRows = rows.slice(state.doc.page * pageSize, state.doc.page * pageSize + pageSize);

  const table = document.getElementById('doc-table');
  table.querySelector('thead').innerHTML = `<tr>
    <th>Nota de Crédito</th><th>Doc. Referencia</th><th>Fecha</th><th>Cliente</th>
    <th>Vendedor</th><th style="text-align:right">Monto</th><th>Motivo</th>
  </tr>`;
  const tbody = table.querySelector('tbody');
  tbody.innerHTML = '';
  pageRows.forEach(r => {
    const tr = el('tr');
    tr.appendChild(el('td', 'mono', r.Documento || ''));
    tr.appendChild(el('td', 'mono', r['Factura/Boleta'] || `<span class="muted">— (${(r.estdoc || 'sin ref.').toLowerCase()})</span>`));
    tr.appendChild(el('td', 'mono muted', fmtFecha(r.emision)));
    tr.appendChild(el('td', null, r.razsoc || ''));
    tr.appendChild(el('td', 'mono muted', VENDOR_NAMES[r.vendedor] || (r.vendedor ? 'Vendedor ' + r.vendedor : '')));
    tr.appendChild(el('td', 'num rej', fmtMoney(r.monimp)));
    tr.appendChild(el('td', 'muted', r.motivo || ''));
    tbody.appendChild(tr);
  });

  document.getElementById('docPageLabel').textContent = `${state.doc.page + 1} / ${maxPage + 1}`;
  document.getElementById('docPrev').disabled = state.doc.page === 0;
  document.getElementById('docNext').disabled = state.doc.page === maxPage;
}

function exportDocumentosCSV() {
  const rows = filteredDocRows();
  const header = ['Nota de Credito', 'Doc. Referencia', 'Fecha', 'Cliente', 'Vendedor', 'Monto', 'Motivo'];
  const lines = [header.map(csvEscape).join(',')];
  rows.forEach(r => {
    lines.push([r.Documento, r['Factura/Boleta'] || 'SIN REFERENCIA', toISO(r.emision), r.razsoc, r.vendedor, (Number(r.monimp) || 0).toFixed(2), r.motivo].map(csvEscape).join(','));
  });
  downloadBlob(`documentos_rechazados_${state.doc.from || 'todos'}.csv`, '\uFEFF' + lines.join('\r\n'), 'text/csv;charset=utf-8;');
}

/* ---------------- tabs / render general ---------------- */

function setTab(tab) {
  state.tab = tab;
  document.querySelectorAll('.tabbar .tab-btn').forEach(b => b.classList.toggle('tab-active', b.dataset.tab === tab));
  ['general', 'conductor', 'vendedor', 'documentos'].forEach(t => {
    document.getElementById('panel-' + t).style.display = (t === tab) ? '' : 'none';
  });
  renderAll();
}

function renderAll() {
  const bannerHayDatos = CORTE_DATA.transportistas.length || CORTE_DATA.ventas.length || CORTE_DATA.nc.length;
  const banner = document.getElementById('pendienteBanner');
  document.getElementById('corteLabel').textContent = getCorteLabel();

  if (state.tab !== 'documentos' && !bannerHayDatos) {
    banner.style.display = 'flex';
    banner.textContent = `Todavía no hay datos cargados para ${getCorteLabel()}. Sube los 3 Excel de ese mes a Storage y este panel se completa solo.`;
  } else {
    banner.style.display = 'none';
  }

  if (state.tab === 'general') renderGeneral();
  else if (state.tab === 'conductor') renderConductor();
  else if (state.tab === 'vendedor') renderVendedor();
  else if (state.tab === 'documentos') drawDocumentos();
}

/* ---------------- impresión (solo el cuadro) ---------------- */

function printTarget(id) {
  const node = document.getElementById(id);
  node.classList.add('print-target');
  window.print();
  setTimeout(() => node.classList.remove('print-target'), 500);
}

/* ---------------- wiring de eventos ---------------- */

function wireEvents() {
  document.querySelectorAll('.tabbar .tab-btn').forEach(b => b.addEventListener('click', () => setTab(b.dataset.tab)));
  document.querySelectorAll('[data-print]').forEach(b => b.addEventListener('click', () => printTarget(b.dataset.print)));

  document.getElementById('condSearch').addEventListener('input', e => { state.cond.search = e.target.value; renderConductor(); });
  document.getElementById('condUmbral').addEventListener('input', e => { state.cond.umbral = parseFloat(e.target.value) || 0; document.getElementById('condAlertLabel').textContent = `Solo alertas ≥${state.cond.umbral}%`; renderConductor(); });
  document.getElementById('condAlertToggle').addEventListener('click', () => {
    state.cond.onlyAlerts = !state.cond.onlyAlerts;
    document.getElementById('condAlertLabel').textContent = state.cond.onlyAlerts ? 'Ver todos' : `Solo alertas ≥${state.cond.umbral}%`;
    renderConductor();
  });

  document.getElementById('vendSearch').addEventListener('input', e => { state.vend.search = e.target.value; renderVendedor(); });

  document.getElementById('docSearch').addEventListener('input', e => { state.doc.search = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docVendor').addEventListener('change', e => { state.doc.vendor = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docMotivo').addEventListener('change', e => { state.doc.motivo = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docMontoMin').addEventListener('input', e => { state.doc.montoMin = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docMontoMax').addEventListener('input', e => { state.doc.montoMax = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docFechaFrom').addEventListener('change', async e => { state.doc.from = e.target.value; await renderDocumentos(); });
  document.getElementById('docFechaTo').addEventListener('change', async e => { state.doc.to = e.target.value; await renderDocumentos(); });
  document.getElementById('docClear').addEventListener('click', async () => {
    state.doc.search = ''; state.doc.vendor = 'all'; state.doc.motivo = 'all';
    state.doc.montoMin = ''; state.doc.montoMax = ''; state.doc.page = 0;
    document.getElementById('docSearch').value = '';
    document.getElementById('docMontoMin').value = '';
    document.getElementById('docMontoMax').value = '';
    await renderDocumentos();
  });
  document.getElementById('docPrev').addEventListener('click', () => { state.doc.page = Math.max(0, state.doc.page - 1); drawDocumentos(); });
  document.getElementById('docNext').addEventListener('click', () => { state.doc.page += 1; drawDocumentos(); });
  document.getElementById('docExportCSV').addEventListener('click', exportDocumentosCSV);
}

/* ---------------- arranque ---------------- */

async function bootApp() {
  initCorteControls();
  wireEvents();
  await loadCorteData();
  await renderDocumentos();
  renderAll();
}

const loginScreen = document.getElementById('loginScreen');
const appRoot = document.getElementById('appRoot');
const loginForm = document.getElementById('loginForm');
const loginError = document.getElementById('loginError');
const loginSubmitBtn = document.getElementById('loginSubmitBtn');
const loginSpinner = document.getElementById('loginSpinner');
const loginBtnText = document.getElementById('loginBtnText');
const logoutBtn = document.getElementById('logoutBtn');

// Cierre de sesión automático por inactividad (30 min — ajusta si quieres).
const INACTIVITY_LIMIT_MS = 30 * 60 * 1000;
let inactivityTimer = null;
let sessionActive = false;
let logoutReason = null;

function resetInactivityTimer() {
  if (!sessionActive) return;
  if (inactivityTimer) clearTimeout(inactivityTimer);
  inactivityTimer = setTimeout(async () => {
    logoutReason = `Tu sesión se cerró por inactividad (${INACTIVITY_LIMIT_MS / 60000} minutos sin uso).`;
    await supabase.auth.signOut();
  }, INACTIVITY_LIMIT_MS);
}
function stopInactivityTimer() {
  if (inactivityTimer) clearTimeout(inactivityTimer);
  inactivityTimer = null;
}
['mousemove', 'mousedown', 'keydown', 'scroll', 'touchstart', 'click'].forEach(evt => {
  document.addEventListener(evt, resetInactivityTimer, { passive: true });
});

let appBooted = false;

async function showApp() {
  loginScreen.style.display = 'none';
  appRoot.style.display = '';
  sessionActive = true;
  resetInactivityTimer();

  if (!appBooted) {
    appBooted = true;
    try {
      if (typeof XLSX === 'undefined') throw new Error('SheetJS (XLSX) no cargó — revisa el <script> de xlsx en index.html.');
      await bootApp();
    } catch (err) {
      console.error(err);
      document.querySelector('.wrap').innerHTML =
        `<div class="empty-state"><div class="empty-title">No se pudo cargar el panel</div><div class="empty-text">${err.message}</div></div>`;
    }
  }
}
function showLogin() {
  appRoot.style.display = 'none';
  loginScreen.style.display = 'flex';
  loginForm.reset();
  sessionActive = false;
  stopInactivityTimer();
  if (logoutReason) {
    loginError.textContent = logoutReason;
    loginError.style.display = 'block';
    logoutReason = null;
  } else {
    loginError.style.display = 'none';
  }
}

// Revisa si ya había sesión al cargar (no aplica aquí, porque la sesión vive
// solo en memoria y se pierde al recargar — pero se deja por si en el futuro
// se cambia a sessionStorage) y reacciona a cualquier cambio después.
supabase.auth.getSession().then(({ data: { session } }) => {
  if (session) { showApp(); } else { showLogin(); }
}).catch(err => {
  console.error('Error revisando sesión existente:', err);
  showLogin();
});

supabase.auth.onAuthStateChange((event, session) => {
  if (event === 'SIGNED_IN' && session) showApp();
  if (event === 'SIGNED_OUT') showLogin();
});

loginForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  loginError.style.display = 'none';
  loginSubmitBtn.disabled = true;
  loginSpinner.style.display = 'inline-block';
  loginBtnText.textContent = 'Ingresando...';

  const email = document.getElementById('loginEmail').value.trim();
  const password = document.getElementById('loginPassword').value;

  try {
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) {
      console.error('Error de login de Supabase:', error);
      loginError.textContent = error.message.includes('Invalid login credentials')
        ? 'Correo o contraseña incorrectos.'
        : 'No se pudo iniciar sesión: ' + error.message;
      loginError.style.display = 'block';
    }
    // Si no hay error, onAuthStateChange se encarga de mostrar la app.
  } catch (err) {
    console.error('Excepción inesperada al iniciar sesión:', err);
    loginError.textContent = 'Error de conexión al intentar iniciar sesión: ' + err.message;
    loginError.style.display = 'block';
  } finally {
    loginSubmitBtn.disabled = false;
    loginSpinner.style.display = 'none';
    loginBtnText.textContent = 'Ingresar';
  }
});

logoutBtn.addEventListener('click', async () => {
  try {
    await supabase.auth.signOut();
  } catch (err) {
    console.error('Error al cerrar sesión:', err);
  }
});
