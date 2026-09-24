/* ============================================================
   Panel de Rechazos — lógica de la app
   Lee los Excel directo de Supabase Storage (sin tablas/SQL),
   igual que el patrón de tu otro proyecto: SheetJS en el navegador.
   ============================================================ */


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

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: { storage: memoryAuthStorage, persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
});

const MES_ABBR = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Set','Oct','Nov','Dic'];

const state = {
  tab: 'general',
  modo: 'mes',            // 'mes' | 'dia'
  mesSel: null,           // 'YYYY-MM'
  diaSel: null,           // 'YYYY-MM-DD'
  joinFrom: null,         // 'YYYY-MM' — desde dónde buscar el chofer hacia atrás
  cond: { search: '', onlyAlerts: false, umbral: 2, sortKey: 'pct', sortDir: 'desc' },
  vend: { search: '', onlyAlerts: false, umbral: 2 },
  doc:  { search: '', vendor: 'all', chofer: 'all', motivo: 'all', montoMin: '', montoMax: '', from: null, to: null, page: 0 },
};

let VENDOR_NAMES = {};                              // código -> nombre (se llena desde Ventas)
let CORTE_DATA = { ventas: [] };
const FILE_CACHE = {};                              // 'Transportistas/Transportistas_2026-09.xlsx' -> filas ya parseadas
let MOTIVO_OVERRIDES = {};                          // PDE -> motivo corregido a mano (persiste en Storage)
const OVERRIDES_PATH = 'Overrides/motivo_overrides.json';
let CHOFER_OVERRIDES = {};                          // PDE -> {codcho,nomcho} corregido a mano (persiste en Storage)
const CHOFER_OVERRIDES_PATH = 'Overrides/chofer_overrides.json';

/* ---------------- utilidades ---------------- */

function fmtMoney(n) {
  n = Number(n) || 0;
  return 'S/ ' + n.toLocaleString('es-PE', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
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
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
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
function buildDocNumber(r) {
  // Ni Transportistas ni Contabilidad NC traen un "Documento" en una sola
  // columna en el Excel sin fórmulas: en ambos se arma con
  // coddoc + sersun + numsun (ej. "BO"+"B002"+"00130369" = "BOB00200130369").
  return `${r.coddoc || ''}${r.sersun || ''}${r.numsun || ''}` || null;
}
function ventasDocNumber(r) {
  // En Ventas el documento se arma distinto: tipo + serie + doc.
  return `${r.tipo || ''}${r.serie || ''}${r.doc || ''}` || null;
}
function cleanRefDoc(raw) {
  if (raw == null) return null;
  const s = String(raw).replace(/\s+/g, '').replace(/-/g, '').trim();
  return s || null;
}
function downloadBlob(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Paleta del panel, en ARGB (formato que pide ExcelJS).
const XL_NAVY = 'FF16356B';
const XL_TEAL = 'FF2E6FB7';   // acento del panel (azul del logo Saga Trans)
const XL_RED = 'FFC1432B';
const XL_CREAM = 'FFF5F3ED';
const XL_STRIPE = 'FFFBF9F4';
const XL_WHITE = 'FFFFFFFF';
const XL_BORDER = { style: 'thin', color: { argb: 'FFE4E0D4' } };

function calcColWidth(col, rows, totalsRow, i) {
  const fmtForWidth = (v) => {
    if (v == null || v === '') return '';
    if (typeof v === 'number') {
      // Aproxima cómo se va a ver ya formateado (miles con coma, % con 2 decimales).
      if (col.numFmt && col.numFmt.indexOf('%') !== -1) return v.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';
      return Math.round(v).toLocaleString('es-PE');
    }
    return String(v);
  };
  let max = String(col.header || '').length;
  rows.forEach(r => { const s = fmtForWidth(r[i]); if (s.length > max) max = s.length; });
  if (totalsRow) { const s = fmtForWidth(totalsRow[i]); if (s.length > max) max = s.length; }
  return Math.min(Math.max(max + 3, col.width ? Math.min(col.width, 10) : 10), 48);
}

async function downloadStyledXlsx({ filename, sheetName, title, subtitle, columns, rows, totalsRow, groupHeader }) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Panel de Rechazos · Saga Trans Confitería';
  const ws = wb.addWorksheet(sheetName || 'Datos', { views: [{ state: 'frozen', ySplit: groupHeader ? 4 : 3 }] });
  const nCols = columns.length;

  // Título
  ws.mergeCells(1, 1, 1, nCols);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = title;
  titleCell.font = { bold: true, size: 15, color: { argb: XL_WHITE }, name: 'Calibri' };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_NAVY } };
  ws.getRow(1).height = 26;

  // Subtítulo
  ws.mergeCells(2, 1, 2, nCols);
  const subCell = ws.getCell(2, 1);
  subCell.value = subtitle || '';
  subCell.font = { italic: true, size: 10, color: { argb: 'FF6B675F' } };
  subCell.alignment = { horizontal: 'center' };
  ws.getRow(2).height = 18;

  let headerRowIdx = 3;

  // Encabezado agrupado opcional (ej. "VENDEDORES" arriba de las columnas de vendedor)
  if (groupHeader) {
    const gRow = ws.getRow(3);
    if (groupHeader.startCol > 1) ws.mergeCells(3, 1, 3, groupHeader.startCol - 1);
    ws.mergeCells(3, groupHeader.startCol, 3, groupHeader.endCol);
    const gCell = ws.getCell(3, groupHeader.startCol);
    gCell.value = groupHeader.label;
    gCell.font = { bold: true, color: { argb: XL_WHITE }, size: 11 };
    gCell.alignment = { horizontal: 'center', vertical: 'middle' };
    gCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_TEAL } };
    gRow.height = 20;
    headerRowIdx = 4;
  }

  // Encabezados de columna
  const headerRow = ws.getRow(headerRowIdx);
  columns.forEach((c, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = c.header;
    cell.font = { bold: true, color: { argb: XL_WHITE }, size: 10 };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_NAVY } };
    cell.border = { top: XL_BORDER, left: XL_BORDER, right: XL_BORDER, bottom: XL_BORDER };
    ws.getColumn(i + 1).width = calcColWidth(c, rows, totalsRow, i);
  });
  headerRow.height = 26;

  // Filas de datos (alternadas)
  rows.forEach((r, ri) => {
    const row = ws.getRow(headerRowIdx + 1 + ri);
    columns.forEach((c, ci) => {
      const cell = row.getCell(ci + 1);
      const raw = r[ci];
      if (c.isCode) {
        cell.value = raw == null || raw === '' ? '' : String(raw);
        cell.numFmt = '@'; // texto exacto: conserva ceros a la izquierda
      } else {
        cell.value = raw === '' || raw == null ? null : raw; // deja vacío en vez de 0
      }
      if (c.numFmt && !c.isCode) cell.numFmt = c.numFmt;
      cell.alignment = { horizontal: c.align || 'left', vertical: 'middle' };
      cell.border = { top: XL_BORDER, left: XL_BORDER, right: XL_BORDER, bottom: XL_BORDER };
      if (ri % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_STRIPE } };
      if (c.red && raw) cell.font = { color: { argb: XL_RED } };
    });
  });

  // Fila de totales
  if (totalsRow) {
    const row = ws.getRow(headerRowIdx + 1 + rows.length);
    columns.forEach((c, ci) => {
      const cell = row.getCell(ci + 1);
      const raw = totalsRow[ci];
      cell.value = raw === '' || raw == null ? null : raw;
      if (c.numFmt) cell.numFmt = c.numFmt;
      cell.font = { bold: true, color: { argb: c.red ? XL_RED : 'FF211F1A' } };
      cell.alignment = { horizontal: c.align || 'left', vertical: 'middle' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_CREAM } };
      cell.border = { top: { style: 'double', color: { argb: 'FF16356B' } }, bottom: XL_BORDER, left: XL_BORDER, right: XL_BORDER };
    });
    row.height = 20;
  }

  const buf = await wb.xlsx.writeBuffer();
  downloadBlob(filename, buf, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
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

  const { data, error } = await supabaseClient.storage.from(STORAGE_BUCKET).download(path);
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

const MES_NOMBRE = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','setiembre','octubre','noviembre','diciembre'];

function getCorteRange() {
  // "Día específico" es un corte ACUMULADO desde el día 1 del mes hasta el
  // día elegido (igual que "Rechazos al [fecha]" del Excel original) — no
  // es solo las transacciones de ese día suelto.
  if (state.modo === 'dia') return { from: state.diaSel.slice(0, 7) + '-01', to: state.diaSel };
  return monthBounds(state.mesSel);
}
function getCorteLabel() {
  if (state.modo === 'dia') {
    const [y, m, d] = state.diaSel.split('-').map(Number);
    return `${d} de ${MES_NOMBRE[m - 1]} del ${y}`;
  }
  const [y, m] = state.mesSel.split('-').map(Number);
  const esMesActual = state.mesSel === isoDate(new Date()).slice(0, 7);
  const dia = esMesActual ? new Date().getDate() : new Date(y, m, 0).getDate();
  return `${dia} de ${MES_NOMBRE[m - 1]} del ${y}`;
}
function syncDocRangeFromCorte() {
  const range = getCorteRange();
  state.doc.from = range.from;
  state.doc.to = range.to;
  document.getElementById('docFechaFrom').value = range.from;
  document.getElementById('docFechaTo').value = range.to;
}

function initCorteControls() {
  const hoy = new Date();
  state.mesSel = isoDate(hoy).slice(0, 7);
  state.diaSel = isoDate(hoy);

  const ventana12 = lastNMonths(12, hoy);
  // Reutiliza el mismo rango de "Cierre de mes" (el más antiguo que ya
  // aparece ahí) como punto de partida para buscar el chofer hacia atrás —
  // así no hay un segundo filtro de fechas por separado.
  state.joinFrom = `${ventana12[0].year}-${pad(ventana12[0].month)}`;

  const sel = document.getElementById('mesSel');
  ventana12.forEach(({ year, month }) => {
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

  syncDocRangeFromCorte();
}

async function onCorteChanged() {
  syncDocRangeFromCorte();
  await loadCorteData();
  await renderDocumentos();
  renderAll();
}

/* ---------------- carga de datos del corte activo ---------------- */

function vendedorKey(r) {
  return r.vendedor || 'OFICINA';
}
function vendedorLabel(cod) {
  if (!cod || cod === 'OFICINA') return 'Vendedor Oficina';
  return VENDOR_NAMES[cod] || ('Vendedor ' + cod);
}

// Una fila "real" es la que se despachó de verdad: tiene número de despacho (nrodsp ≠ 0)
// y monto despachado (totdsp > 0). A veces el Excel trae además una fila vacía del mismo
// documento (INTERNO, despacho 0, S/ 0): esa no debe ganarle a la fila real.
function isDespachoReal(r) {
  const n = String(r.nrodsp == null ? '' : r.nrodsp).replace(/[\r\n]+/g, '').trim();
  return n !== '' && Number(n) !== 0 && (Number(r.totdsp) || 0) > 0;
}
function buildDocToChoferMap(transportistas) {
  const map = {};
  const real = {};   // doc -> true si ya se guardó una fila realmente despachada
  transportistas.forEach(r => {
    const doc = buildDocNumber(r);
    if (doc) {
      const esReal = isDespachoReal(r);
      if (real[doc] && !esReal) return;   // no pisar una fila real con una vacía
      if (esReal) real[doc] = true;
      // Algunas celdas "vacías" en el Excel en realidad traen basura
      // invisible (saltos de línea, espacios) en vez de estar realmente
      // vacías — hay que limpiarlas o JS las trata como "con contenido".
      const desmotClean = (r.desmot == null ? '' : String(r.desmot)).replace(/[\r\n]+/g, '').trim();
      map[doc] = { codcho: r.codcho || null, nomcho: r.nomcho || r.codcho || null, desmot: desmotClean || null, desmotOriginal: desmotClean || null, pde: r.nrodsp || null };
    }
  });
  return map;
}
function buildNcDocToRefMap(nc) {
  // Para las líneas 'D' de Ventas, su propio "documento" (tipo+serie+doc)
  // YA ES el número de la nota de crédito, no el de la venta original.
  // Este mapa va: documento propio de la NC -> su "referencia" (el
  // documento original que sustenta), que es lo que sí puede coincidir
  // con un Documento de Transportistas.
  const map = {};
  nc.forEach(r => {
    const ncDoc = buildDocNumber(r);
    const ref = cleanRefDoc(r.referencia);
    if (ncDoc && ref) map[ncDoc] = ref;
  });
  return map;
}
function resolveChofer(ventaRow, docToChofer, ncDocToRef) {
  const own = ventasDocNumber(ventaRow);
  if (!own) return null;
  // Intento 1: el documento de la venta coincide directo con Transportistas
  // (así funciona para las líneas normales de venta).
  let hit = docToChofer[own];
  if (!hit) {
    // Intento 2: el documento de la venta es en realidad una nota de
    // crédito — se busca su referencia (documento original) y con esa se
    // busca el chofer.
    const ref = ncDocToRef[own];
    if (ref) hit = docToChofer[ref];
  }
  if (hit && hit.pde && MOTIVO_OVERRIDES[hit.pde] != null) {
    hit = { ...hit, desmot: MOTIVO_OVERRIDES[hit.pde] };
  }
  if (hit && hit.pde && CHOFER_OVERRIDES[hit.pde]) {
    const corr = CHOFER_OVERRIDES[hit.pde];
    hit = { ...hit, codcho: corr.codcho, nomcho: corr.nomcho, choferCorregido: true };
  }
  return hit || null;
}

/* ---------------- correcciones de Motivo (persistentes en Storage) ---------------- */

async function loadMotivoOverrides() {
  try {
    const { data, error } = await supabaseClient.storage.from(STORAGE_BUCKET).download(OVERRIDES_PATH);
    if (error) { MOTIVO_OVERRIDES = {}; return; }
    const text = await data.text();
    MOTIVO_OVERRIDES = text ? JSON.parse(text) : {};
  } catch (err) {
    console.warn('No se pudieron cargar las correcciones de motivo (probablemente aún no existen):', err.message);
    MOTIVO_OVERRIDES = {};
  }
}

async function saveMotivoOverride(pde, motivo) {
  if (!pde) return;
  const value = (motivo || '').trim();
  if (value) MOTIVO_OVERRIDES[pde] = value;
  else delete MOTIVO_OVERRIDES[pde];

  // Actualiza en memoria todas las filas ya cargadas que compartan ese PDE,
  // para que se vea el cambio en todas las pestañas sin recargar la página.
  [CORTE_DATA.ventas, DOC_ROWS_CACHE].forEach(list => {
    (list || []).forEach(r => { if (r._chofer && r._chofer.pde === pde) r._chofer.desmot = value || null; });
  });

  try {
    const blob = new Blob([JSON.stringify(MOTIVO_OVERRIDES, null, 2)], { type: 'application/json' });
    const { error } = await supabaseClient.storage.from(STORAGE_BUCKET).upload(OVERRIDES_PATH, blob, { upsert: true, contentType: 'application/json' });
    if (error) throw error;
  } catch (err) {
    console.error('No se pudo guardar la corrección de motivo en Storage:', err);
    alert('No se pudo guardar el cambio en Supabase (revisa la política de escritura de Overrides/). El cambio queda solo en esta sesión por ahora.');
  }
}

async function loadChoferOverrides() {
  try {
    const { data, error } = await supabaseClient.storage.from(STORAGE_BUCKET).download(CHOFER_OVERRIDES_PATH);
    if (error) { CHOFER_OVERRIDES = {}; return; }
    const text = await data.text();
    CHOFER_OVERRIDES = text ? JSON.parse(text) : {};
  } catch (err) {
    console.warn('No se pudieron cargar las correcciones de chofer (probablemente aún no existen):', err.message);
    CHOFER_OVERRIDES = {};
  }
}

async function saveChoferOverride(pde, codcho, nomcho) {
  if (!pde) return;
  if (codcho) CHOFER_OVERRIDES[pde] = { codcho, nomcho };
  else delete CHOFER_OVERRIDES[pde];

  [CORTE_DATA.ventas, DOC_ROWS_CACHE].forEach(list => {
    (list || []).forEach(r => {
      if (r._chofer && r._chofer.pde === pde) {
        if (codcho) { r._chofer.codcho = codcho; r._chofer.nomcho = nomcho; r._chofer.choferCorregido = true; }
        else { r._chofer.choferCorregido = false; }
      }
    });
  });

  try {
    const blob = new Blob([JSON.stringify(CHOFER_OVERRIDES, null, 2)], { type: 'application/json' });
    const { error } = await supabaseClient.storage.from(STORAGE_BUCKET).upload(CHOFER_OVERRIDES_PATH, blob, { upsert: true, contentType: 'application/json' });
    if (error) throw error;
  } catch (err) {
    console.error('No se pudo guardar la corrección de chofer en Storage:', err);
    alert('No se pudo guardar el cambio en Supabase (revisa la política de escritura de Overrides/). El cambio queda solo en esta sesión por ahora.');
  }
}

function monthsInRange(from, to) {
  // Devuelve ['YYYY-MM', ...] desde "from" hasta "to" (inclusive), en orden.
  const [fy, fm] = from.split('-').map(Number);
  const [ty, tm] = to.split('-').map(Number);
  const out = [];
  let y = fy, m = fm;
  while (y < ty || (y === ty && m <= tm)) {
    out.push(`${y}-${pad(m)}`);
    m++; if (m > 12) { m = 1; y++; }
  }
  return out;
}

let LAST_DOC_TO_CHOFER = {};

async function buildJoinMaps(uptoPeriodo) {
  // Un rechazo puede procesarse en Ventas varios meses después del despacho
  // que lo originó, así que el cruce documento->chofer mira hacia atrás
  // (no solo el mes del corte) para encontrarlo igual. Cuánto atrás mirar
  // lo controla el selector "Buscar chofer desde" del encabezado.
  const desde = state.joinFrom && state.joinFrom <= uptoPeriodo ? state.joinFrom : uptoPeriodo;
  const periodos = monthsInRange(desde, uptoPeriodo);
  const files = await Promise.all(periodos.map(p => Promise.all([
    loadModuleFile('Transportistas', 'Transportistas', p),
    loadModuleFile('ContabilidadNC', 'ContabilidadNC', p),
  ])));
  const allTransportistas = files.flatMap(f => f[0]);
  const allNc = files.flatMap(f => f[1]);
  const docToChofer = buildDocToChoferMap(allTransportistas);
  LAST_DOC_TO_CHOFER = docToChofer;
  return {
    docToChofer,
    ncDocToRef: buildNcDocToRefMap(allNc),
  };
}

async function loadCorteData() {
  const periodo = state.modo === 'dia' ? state.diaSel.slice(0, 7) : state.mesSel;
  const [v, { docToChofer, ncDocToRef }] = await Promise.all([
    loadModuleFile('Ventas', 'Ventas', periodo),
    buildJoinMaps(periodo),
  ]);

  let ventas = v;
  if (state.modo === 'dia') {
    const { from, to } = getCorteRange();
    ventas = ventas.filter(r => { const iso = toISO(r.fecha); return iso && iso >= from && iso <= to; });
  }
  // Resuelve y guarda el chofer de cada línea de venta una sola vez.
  ventas.forEach(r => { r._chofer = resolveChofer(r, docToChofer, ncDocToRef); });

  CORTE_DATA = { ventas };

  ventas.forEach(r => { if (r.vendedor && r.nombrevendedor && !VENDOR_NAMES[r.vendedor]) VENDOR_NAMES[r.vendedor] = r.nombrevendedor; });
}

async function loadDocumentosData() {
  // "Documentos rechazados" = líneas de Ventas con vtadvo = 'D' (devolución).
  if (!state.doc.from || !state.doc.to) return [];
  const periodos = monthsBetween(state.doc.from, state.doc.to);
  const { docToChofer, ncDocToRef } = await buildJoinMaps(state.doc.to.slice(0, 7));
  const results = await Promise.all(periodos.map(async p => {
    const v = await loadModuleFile('Ventas', 'Ventas', p);
    v.forEach(r => { r._chofer = resolveChofer(r, docToChofer, ncDocToRef); });
    return v;
  }));
  const all = results.flat();
  return all.filter(r => {
    if (r.vtadvo !== 'D') return false;
    const iso = toISO(r.fecha);
    return iso && iso >= state.doc.from && iso <= state.doc.to;
  });
}

/* ---------------- Vista general (cruce Chofer x Vendedor) ---------------- */

async function exportGeneralCSV() {
  const cross = [];
  CORTE_DATA.ventas.forEach(r => {
    if (r.vtadvo !== 'D') return;
    const monto = -(Number(r.soles) || 0);
    if (monto <= 0.01) return;
    const hit = r._chofer;
    cross.push({
      choferCod: hit && hit.codcho ? hit.codcho : 'SIN_CHOFER',
      chofer: hit && hit.codcho ? (hit.nomcho || hit.codcho) : 'Sin identificar',
      vendCod: vendedorKey(r),
      monto,
    });
  });
  const vendCols = Array.from(new Set(cross.map(c => c.vendCod))).sort();
  const choferNombre = {};
  cross.forEach(c => { choferNombre[c.choferCod] = c.chofer; });
  const choferCods = Array.from(new Set(cross.map(c => c.choferCod))).sort((a, b) => (choferNombre[a] || '').localeCompare(choferNombre[b] || ''));
  const matrix = {};
  choferCods.forEach(c => matrix[c] = {});
  cross.forEach(c => { matrix[c.choferCod][c.vendCod] = (matrix[c.choferCod][c.vendCod] || 0) + c.monto; });

  const columns = [{ header: 'Chofer', width: 26 }]
    .concat(vendCols.map(v => ({ header: vendedorLabel(v).toUpperCase(), width: 16, numFmt: '#,##0', align: 'center', red: true })))
    .concat([{ header: 'Total', width: 16, numFmt: '#,##0', align: 'center', red: true }]);

  const colTotals = vendCols.map(() => 0);
  let grandTotal = 0;
  const rows = choferCods.map(ch => {
    let rowTotal = 0;
    const vals = vendCols.map((v, i) => { const val = matrix[ch][v] || 0; rowTotal += val; colTotals[i] += val; return val || ''; });
    grandTotal += rowTotal;
    return [choferNombre[ch]].concat(vals).concat([rowTotal]);
  });
  const totalsRow = ['Total general'].concat(colTotals).concat([grandTotal]);

  await downloadStyledXlsx({
    filename: `vista_general_${getCorteRange().to}.xlsx`,
    sheetName: 'Vista General',
    title: 'RECHAZOS SAGA TRANS · Confitería — Vista General',
    subtitle: `Cruce Chofer × Vendedor · Corte al ${getCorteLabel()}`,
    groupHeader: { label: 'VENDEDORES', startCol: 2, endCol: vendCols.length + 1 },
    columns,
    rows,
    totalsRow,
  });
}

function renderGeneral() {
  document.querySelectorAll('.corteLabelInline').forEach(e => e.textContent = getCorteLabel());

  const cross = [];
  CORTE_DATA.ventas.forEach(r => {
    if (r.vtadvo !== 'D') return;
    const monto = -(Number(r.soles) || 0);
    if (monto <= 0.01) return;
    const hit = r._chofer;
    cross.push({
      choferCod: hit && hit.codcho ? hit.codcho : 'SIN_CHOFER',
      chofer: hit ? (hit.nomcho || hit.codcho || 'Sin identificar') : 'Sin identificar',
      vendCod: vendedorKey(r),
      monto,
    });
  });

  const kpis = document.getElementById('cross-kpis');
  kpis.innerHTML = '';
  kpis.appendChild(kpiCard('navy', 'Pedidos rechazados', cross.length));
  kpis.appendChild(kpiCard('teal', 'Choferes involucrados', new Set(cross.map(c => c.choferCod)).size));
  kpis.appendChild(kpiCard('red', 'Monto total rechazado', fmtMoney(cross.reduce((s, c) => s + c.monto, 0))));

  const vendCols = Array.from(new Set(cross.map(c => c.vendCod))).sort();
  const choferCods = Array.from(new Set(cross.map(c => c.choferCod)));
  const choferNombre = {};
  cross.forEach(c => { choferNombre[c.choferCod] = c.chofer; });
  choferCods.sort((a, b) => (choferNombre[a] || '').localeCompare(choferNombre[b] || ''));

  const matrix = {};
  choferCods.forEach(c => matrix[c] = {});
  cross.forEach(c => { matrix[c.choferCod][c.vendCod] = (matrix[c.choferCod][c.vendCod] || 0) + c.monto; });

  const table = document.getElementById('cross-table');
  const thead = table.querySelector('thead'), tbody = table.querySelector('tbody'), tfoot = table.querySelector('tfoot');
  thead.innerHTML = ''; tbody.innerHTML = ''; tfoot.innerHTML = '';

  if (!choferCods.length) {
    return;
  }

  const trTop = el('tr'); trTop.style.background = '#F5F3ED';
  trTop.appendChild(el('th'));
  const thVend = el('th', null, 'Vendedores'); thVend.colSpan = vendCols.length; thVend.style.textAlign = 'center';
  trTop.appendChild(thVend);
  const thTotal = el('th', null, 'Total<br>general'); thTotal.rowSpan = 2; thTotal.style.textAlign = 'center'; thTotal.style.verticalAlign = 'middle'; thTotal.style.background = '#F5F3ED';
  trTop.appendChild(thTotal);
  thead.appendChild(trTop);

  const trHead = el('tr'); trHead.style.background = '#F5F3ED';
  trHead.appendChild(el('th', null, 'Chofer'));
  vendCols.forEach(v => {
    const th = el('th', null, vendedorLabel(v).toUpperCase());
    th.style.textAlign = 'center';
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  const colTotals = vendCols.map(() => 0);
  let grandTotal = 0;
  choferCods.forEach(ch => {
    const tr = el('tr');
    tr.appendChild(el('td', null, (choferNombre[ch] || ch).toUpperCase()));
    let rowTotal = 0;
    vendCols.forEach((v, i) => {
      const val = matrix[ch][v];
      const td = el('td', 'num rej', val ? val.toLocaleString('es-PE', { minimumFractionDigits: 0, maximumFractionDigits: 0 }) : '');
      td.style.textAlign = 'center';
      if (val) {
        td.classList.add('clickable');
        td.title = 'Ver documentos rechazados';
        td.addEventListener('click', () => goToDocs({ chofer: ch, vendedor: v }));
      }
      tr.appendChild(td);
      if (val) { rowTotal += val; colTotals[i] += val; }
    });
    grandTotal += rowTotal;
    const tdTotal = el('td', 'num rej clickable', rowTotal.toLocaleString('es-PE', { minimumFractionDigits: 0, maximumFractionDigits: 0 }));
    tdTotal.title = 'Ver documentos rechazados de este chofer';
    tdTotal.addEventListener('click', () => goToDocs({ chofer: ch }));
    tr.appendChild(tdTotal);
    tbody.appendChild(tr);
  });

  const trFoot = el('tr');
  trFoot.appendChild(el('td', null, 'Total general'));
  colTotals.forEach((t, i) => {
    const td = el('td', 'num rej', t.toLocaleString('es-PE', { minimumFractionDigits: 0, maximumFractionDigits: 0 }));
    td.style.textAlign = 'center';
    if (t) { td.classList.add('clickable'); td.title = 'Ver documentos rechazados'; td.addEventListener('click', () => goToDocs({ vendedor: vendCols[i] })); }
    trFoot.appendChild(td);
  });
  trFoot.appendChild(el('td', 'num rej', grandTotal.toLocaleString('es-PE', { minimumFractionDigits: 0, maximumFractionDigits: 0 })));
  tfoot.appendChild(trFoot);

}

/* ---------------- Por conductor ---------------- */

async function exportConductorCSV() {
  const rows = aggregateConductor().sort((a, b) => b.monto - a.monto);
  const totals = rows.reduce((a, r) => ({ facturado: a.facturado + r.facturado, ventaReal: a.ventaReal + r.ventaReal, monto: a.monto + r.monto, pedidos: a.pedidos + r.pedidos }), { facturado: 0, ventaReal: 0, monto: 0, pedidos: 0 });
  const totalPct = totals.facturado ? totals.monto / totals.facturado : 0;

  const columns = [
    { header: 'Código', width: 12, isCode: true, align: 'center' },
    { header: 'Chofer', width: 30 },
    { header: 'Facturado', width: 16, numFmt: '#,##0', align: 'right' },
    { header: 'Venta real', width: 16, numFmt: '#,##0', align: 'right' },
    { header: 'Monto rechazado', width: 18, numFmt: '#,##0', align: 'right', red: true },
    { header: 'Pedidos', width: 12, numFmt: '#,##0', align: 'center' },
    { header: '% Rechazo', width: 13, numFmt: '0.00"%"', align: 'right', red: true },
  ];
  const dataRows = rows.map(r => [r.cod, r.nombre, r.facturado, r.ventaReal, r.monto, r.pedidos, r.pct * 100]);
  const totalsRow = ['', 'Total general', totals.facturado, totals.ventaReal, totals.monto, totals.pedidos, totalPct * 100];

  await downloadStyledXlsx({
    filename: `por_conductor_${getCorteRange().to}.xlsx`,
    sheetName: 'Por Conductor',
    title: 'RECHAZOS SAGA TRANS · Confitería — Por Conductor',
    subtitle: `Corte al ${getCorteLabel()}`,
    columns,
    rows: dataRows,
    totalsRow,
  });
}

function aggregateConductor() {
  const map = {};
  CORTE_DATA.ventas.forEach(r => {
    const hit = r._chofer;
    const key = hit && hit.codcho ? hit.codcho : 'SIN_CHOFER';
    if (!map[key]) map[key] = { cod: key, nombre: hit ? (hit.nomcho || hit.codcho) : 'Sin identificar', facturado: 0, monto: 0, pedidos: 0 };
    const soles = Number(r.soles) || 0;
    if (r.vtadvo === 'V') {
      map[key].facturado += soles;
    } else if (r.vtadvo === 'D') {
      map[key].monto += -soles;
      map[key].pedidos += 1;
    }
  });
  // Venta real = facturado - rechazado (igual que la fórmula del Excel: E=D-F).
  return Object.values(map).map(r => ({ ...r, ventaReal: r.facturado - r.monto, pct: r.facturado ? r.monto / r.facturado : 0 }));
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
    <th><button data-sort="nombre">Chofer</button></th>
    <th style="text-align:right"><button data-sort="facturado">Facturado</button></th>
    <th style="text-align:right">Venta real</th>
    <th style="text-align:right"><button data-sort="monto">Monto rech.</button></th>
    <th style="text-align:right"><button data-sort="pedidos">Pedidos</button></th>
    <th style="text-align:center"><button data-sort="pct">% Rechazo</button></th>
    <th style="width:20px"></th>
  </tr>`;
  thead.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
    const k = b.dataset.sort;
    state.cond.sortDir = (state.cond.sortKey === k && state.cond.sortDir === 'desc') ? 'asc' : 'desc';
    state.cond.sortKey = k;
    renderConductor();
  }));

  tbody.innerHTML = '';
  rows.forEach(r => {
    const tr = el('tr', 'clickable');
    tr.addEventListener('click', () => goToDocs({ chofer: r.cod }));
    tr.appendChild(el('td', null, `${r.nombre}<div class="sub">${r.cod}</div>`));
    tr.appendChild(el('td', 'num', fmtMoney(r.facturado)));
    tr.appendChild(el('td', 'num muted', fmtMoney(r.ventaReal)));
    tr.appendChild(el('td', 'num rej', fmtMoney(r.monto)));
    tr.appendChild(el('td', 'num', r.pedidos));
    tr.appendChild(pctCell(r.pct, state.cond.umbral));
    tr.appendChild(el('td', 'muted', '&rsaquo;'));
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
  trT.appendChild(el('td'));
  tfoot.appendChild(trT);

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

function aggregateVendedor() {
  const map = {};
  CORTE_DATA.ventas.forEach(r => {
    const key = vendedorKey(r);
    if (!map[key]) map[key] = { cod: key, facturado: 0, monto: 0, pedidos: 0 };
    const soles = Number(r.soles) || 0;
    if (r.vtadvo === 'V') {
      map[key].facturado += soles;
    } else if (r.vtadvo === 'D') {
      map[key].monto += -soles;
      map[key].pedidos += 1;
    }
  });
  return Object.values(map).map(r => ({ ...r, ventaReal: r.facturado - r.monto, pct: r.facturado ? r.monto / r.facturado : 0 }));
}

async function exportVendedorCSV() {
  const rows = aggregateVendedor().sort((a, b) => b.monto - a.monto);
  const totals = rows.reduce((a, r) => ({ facturado: a.facturado + r.facturado, ventaReal: a.ventaReal + r.ventaReal, monto: a.monto + r.monto, pedidos: a.pedidos + r.pedidos }), { facturado: 0, ventaReal: 0, monto: 0, pedidos: 0 });
  const totalPct = totals.facturado ? totals.monto / totals.facturado : 0;

  const columns = [
    { header: 'Código', width: 12, isCode: true, align: 'center' },
    { header: 'Vendedor', width: 30 },
    { header: 'Facturado', width: 16, numFmt: '#,##0', align: 'right' },
    { header: 'Venta real', width: 16, numFmt: '#,##0', align: 'right' },
    { header: 'Monto rechazado', width: 18, numFmt: '#,##0', align: 'right', red: true },
    { header: 'Pedidos', width: 12, numFmt: '#,##0', align: 'center' },
    { header: '% Rechazo', width: 13, numFmt: '0.00"%"', align: 'right', red: true },
  ];
  const dataRows = rows.map(r => [r.cod, vendedorLabel(r.cod), r.facturado, r.ventaReal, r.monto, r.pedidos, r.pct * 100]);
  const totalsRow = ['', 'Total general', totals.facturado, totals.ventaReal, totals.monto, totals.pedidos, totalPct * 100];

  await downloadStyledXlsx({
    filename: `por_vendedor_${getCorteRange().to}.xlsx`,
    sheetName: 'Por Vendedor',
    title: 'RECHAZOS SAGA TRANS · Confitería — Por Vendedor',
    subtitle: `Corte al ${getCorteLabel()}`,
    columns,
    rows: dataRows,
    totalsRow,
  });
}

function renderVendedor() {
  document.querySelectorAll('.corteLabelInline').forEach(e => e.textContent = getCorteLabel());

  const all = aggregateVendedor();

  const totals = all.reduce((a, r) => ({ facturado: a.facturado + r.facturado, ventaReal: a.ventaReal + r.ventaReal, monto: a.monto + r.monto, pedidos: a.pedidos + r.pedidos }), { facturado: 0, ventaReal: 0, monto: 0, pedidos: 0 });
  const totalPct = totals.facturado ? totals.monto / totals.facturado : 0;

  const kpis = document.getElementById('vend-kpis');
  kpis.innerHTML = '';
  kpis.appendChild(kpiCard('navy', 'Facturado', fmtMoney(totals.facturado)));
  kpis.appendChild(kpiCard('teal', 'Venta real', fmtMoney(totals.ventaReal)));
  kpis.appendChild(kpiCard('red', 'Monto rechazado', fmtMoney(totals.monto)));
  kpis.appendChild(kpiCard('navy', '% Rechazo global', fmtPct(totalPct)));

  const q = state.vend.search.toLowerCase();
  let rows = all.filter(v => {
    const nm = vendedorLabel(v.cod).toLowerCase();
    return !q || String(v.cod).toLowerCase().includes(q) || nm.includes(q);
  }).sort((a, b) => b.monto - a.monto);
  if (state.vend.onlyAlerts) rows = rows.filter(v => v.pct * 100 >= state.vend.umbral);

  const table = document.getElementById('vend-table');
  const thead = table.querySelector('thead'), tbody = table.querySelector('tbody'), tfoot = table.querySelector('tfoot');
  thead.innerHTML = `<tr>
    <th>Vendedor</th>
    <th style="text-align:right">Facturado</th>
    <th style="text-align:right">Venta real</th>
    <th style="text-align:right">Monto rech.</th>
    <th style="text-align:right">Pedidos</th>
    <th style="text-align:center">% Rechazo</th>
    <th style="width:20px"></th>
  </tr>`;

  tbody.innerHTML = '';
  rows.forEach(v => {
    const tr = el('tr', 'clickable');
    tr.addEventListener('click', () => goToDocs({ vendedor: v.cod }));
    const nombre = v.cod === 'OFICINA' ? 'Vendedor Oficina' : VENDOR_NAMES[v.cod];
    tr.appendChild(el('td', null, nombre ? `${nombre}${v.cod === 'OFICINA' ? '' : `<div class="sub">Vendedor ${v.cod}</div>`}` : `Vendedor ${v.cod}<div class="sub" style="font-style:italic">nombre pendiente</div>`));
    tr.appendChild(el('td', 'num', fmtMoney(v.facturado)));
    tr.appendChild(el('td', 'num muted', fmtMoney(v.ventaReal)));
    tr.appendChild(el('td', 'num rej', fmtMoney(v.monto)));
    tr.appendChild(el('td', 'num', v.pedidos));
    tr.appendChild(pctCell(v.pct, state.vend.umbral));
    tr.appendChild(el('td', 'muted', '&rsaquo;'));
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = '';
  const trT = el('tr');
  trT.appendChild(el('td', null, 'Total general'));
  trT.appendChild(el('td', 'num', fmtMoney(totals.facturado)));
  trT.appendChild(el('td', 'num', fmtMoney(totals.ventaReal)));
  trT.appendChild(el('td', 'num rej', fmtMoney(totals.monto)));
  trT.appendChild(el('td', 'num', totals.pedidos));
  trT.appendChild(pctCell(totalPct, state.vend.umbral));
  trT.appendChild(el('td'));
  tfoot.appendChild(trT);

}

/* ---------------- Rechazos por motivo / reparto ---------------- */

async function exportMotivoCSV() {
  const rows = [];
  CORTE_DATA.ventas.forEach(r => {
    if (r.vtadvo !== 'D') return;
    const monto = -(Number(r.soles) || 0);
    const hit = r._chofer;
    rows.push({
      chofer: hit && hit.codcho ? (hit.nomcho || hit.codcho) : 'Venta Oficina',
      motivo: (hit && hit.desmot) || 'Sin motivo',
      monto,
    });
  });
  const motivos = Array.from(new Set(rows.map(r => r.motivo))).sort();
  const choferes = Array.from(new Set(rows.map(r => r.chofer))).sort((a, b) => {
    if (a === 'Venta Oficina') return 1;
    if (b === 'Venta Oficina') return -1;
    return a.localeCompare(b);
  });
  const matrix = {};
  choferes.forEach(c => matrix[c] = {});
  rows.forEach(r => { matrix[r.chofer][r.motivo] = (matrix[r.chofer][r.motivo] || 0) + r.monto; });

  const columns = [{ header: 'Chofer', width: 26 }]
    .concat(motivos.map(m => ({ header: m.toUpperCase(), width: 15, numFmt: '#,##0', align: 'center', red: true })))
    .concat([{ header: 'Total', width: 15, numFmt: '#,##0', align: 'center', red: true }]);

  const colTotals = motivos.map(() => 0);
  let grandTotal = 0;
  const dataRows = choferes.map(ch => {
    let rowTotal = 0;
    const vals = motivos.map((m, i) => { const val = matrix[ch][m] || 0; rowTotal += val; colTotals[i] += val; return val || ''; });
    grandTotal += rowTotal;
    return [ch].concat(vals).concat([rowTotal]);
  });
  const totalsRow = ['Total'].concat(colTotals).concat([grandTotal]);

  await downloadStyledXlsx({
    filename: `rechazos_por_motivo_${getCorteRange().to}.xlsx`,
    sheetName: 'Rechazos por Motivo',
    title: 'RECHAZOS POR MOTIVO / REPARTO',
    subtitle: `Saga Trans Confitería · Corte al ${getCorteLabel()}`,
    columns,
    rows: dataRows,
    totalsRow,
  });
}

function renderMotivo() {
  document.querySelectorAll('.corteLabelInline').forEach(e => e.textContent = getCorteLabel());

  const rows = [];
  CORTE_DATA.ventas.forEach(r => {
    if (r.vtadvo !== 'D') return;
    const monto = Number(r.soles) || 0; // ya viene negativo
    const hit = r._chofer;
    rows.push({
      chofer: hit && hit.codcho ? (hit.nomcho || hit.codcho) : 'Venta Oficina',
      motivo: (hit && hit.desmot) || 'Sin motivo',
      monto,
    });
  });

  const kpis = document.getElementById('motivo-kpis');
  kpis.innerHTML = '';
  kpis.appendChild(kpiCard('navy', 'Líneas rechazadas', rows.length));
  kpis.appendChild(kpiCard('teal', 'Motivos distintos', new Set(rows.map(r => r.motivo)).size));
  kpis.appendChild(kpiCard('red', 'Monto total', fmtMoneyNeg(rows.reduce((s, r) => s + r.monto, 0))));

  const motivos = Array.from(new Set(rows.map(r => r.motivo))).sort();
  const choferes = Array.from(new Set(rows.map(r => r.chofer))).sort((a, b) => {
    if (a === 'Venta Oficina') return 1;
    if (b === 'Venta Oficina') return -1;
    return a.localeCompare(b);
  });

  const matrix = {};
  choferes.forEach(c => matrix[c] = {});
  rows.forEach(r => { matrix[r.chofer][r.motivo] = (matrix[r.chofer][r.motivo] || 0) + r.monto; });

  const table = document.getElementById('motivo-table');
  const thead = table.querySelector('thead'), tbody = table.querySelector('tbody'), tfoot = table.querySelector('tfoot');
  thead.innerHTML = ''; tbody.innerHTML = ''; tfoot.innerHTML = '';

  if (!choferes.length) {
    return;
  }

  const trHead = el('tr'); trHead.style.background = '#F5F3ED';
  trHead.appendChild(el('th', null, 'Chofer'));
  motivos.forEach(m => {
    const th = el('th', null, m.toUpperCase());
    th.style.textAlign = 'center';
    trHead.appendChild(th);
  });
  const thTotal = el('th', null, 'Total');
  thTotal.style.textAlign = 'center';
  trHead.appendChild(thTotal);
  thead.appendChild(trHead);

  const colTotals = motivos.map(() => 0);
  let grandTotal = 0;
  choferes.forEach(ch => {
    const tr = el('tr');
    tr.appendChild(el('td', null, ch.toUpperCase()));
    let rowTotal = 0;
    motivos.forEach((m, i) => {
      const val = matrix[ch][m];
      const td = el('td', 'num rej', val ? fmtMoneyNeg(val) : '');
      td.style.textAlign = 'center';
      if (val) {
        td.classList.add('clickable');
        td.title = 'Ver documentos rechazados';
        td.addEventListener('click', () => goToDocs({ chofer: ch === 'Venta Oficina' ? 'SIN_CHOFER' : findChoferCod(ch), motivo: m === 'Sin motivo' ? null : m }));
      }
      tr.appendChild(td);
      if (val) { rowTotal += val; colTotals[i] += val; }
    });
    grandTotal += rowTotal;
    tr.appendChild(el('td', 'num rej', fmtMoneyNeg(rowTotal)));
    tbody.appendChild(tr);
  });

  const trFoot = el('tr');
  trFoot.appendChild(el('td', null, 'Total'));
  colTotals.forEach(t => {
    const td = el('td', 'num rej', fmtMoneyNeg(t));
    td.style.textAlign = 'center';
    trFoot.appendChild(td);
  });
  trFoot.appendChild(el('td', 'num rej', fmtMoneyNeg(grandTotal)));
  tfoot.appendChild(trFoot);

}

function findChoferCod(nombre) {
  for (const r of CORTE_DATA.ventas) {
    if (r._chofer && r._chofer.codcho && (r._chofer.nomcho || r._chofer.codcho) === nombre) return r._chofer.codcho;
  }
  return 'all';
}

function goToDocs({ chofer, vendedor, motivo } = {}) {
  state.doc.chofer = chofer || 'all';
  state.doc.vendor = vendedor || 'all';
  state.doc.motivo = motivo || 'all';
  state.doc.page = 0;
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
  const choferSel = document.getElementById('docChofer');
  const motivoSel = document.getElementById('docMotivo');
  const prevVendor = state.doc.vendor, prevChofer = state.doc.chofer, prevMotivo = state.doc.motivo;

  const vendors = Array.from(new Set(rows.map(r => vendedorKey(r)))).sort();
  vendorSel.innerHTML = '';
  vendorSel.appendChild(new Option('Todos', 'all'));
  vendors.forEach(v => vendorSel.appendChild(new Option(vendedorLabel(v), v)));
  vendorSel.value = vendors.includes(prevVendor) ? prevVendor : 'all';
  state.doc.vendor = vendorSel.value;

  const choferMap = {};
  let hayNoIdentificado = false;
  rows.forEach(r => {
    if (r._chofer && r._chofer.codcho) choferMap[r._chofer.codcho] = r._chofer.nomcho || r._chofer.codcho;
    else hayNoIdentificado = true;
  });
  const choferCods = Object.keys(choferMap).sort((a, b) => choferMap[a].localeCompare(choferMap[b]));
  choferSel.innerHTML = '';
  choferSel.appendChild(new Option('Todos', 'all'));
  choferCods.forEach(c => choferSel.appendChild(new Option(choferMap[c], c)));
  if (hayNoIdentificado) choferSel.appendChild(new Option('Sin identificar', 'SIN_CHOFER'));
  const choferValidos = choferCods.concat(hayNoIdentificado ? ['SIN_CHOFER'] : []);
  choferSel.value = choferValidos.includes(prevChofer) ? prevChofer : 'all';
  state.doc.chofer = choferSel.value;

  const motivos = Array.from(new Set(rows.map(r => r._chofer && r._chofer.desmot).filter(Boolean))).sort();
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
    const choferCod = r._chofer ? r._chofer.codcho : null;
    if (q && !((String(ventasDocNumber(r) || '')).toLowerCase().includes(q) || (String((r._chofer && r._chofer.pde) || '')).toLowerCase().includes(q) || (String(codCliente(r) || '')).toLowerCase().includes(q) || (String(r.nombrecliente || '')).toLowerCase().includes(q))) return false;
    if (state.doc.vendor !== 'all' && vendedorKey(r) !== state.doc.vendor) return false;
    if (state.doc.chofer === 'SIN_CHOFER') { if (choferCod) return false; }
    else if (state.doc.chofer !== 'all' && choferCod !== state.doc.chofer) return false;
    if (state.doc.motivo !== 'all' && (r._chofer && r._chofer.desmot) !== state.doc.motivo) return false;
    const monto = -(Number(r.soles) || 0);
    if (min !== null && monto < min) return false;
    if (max !== null && monto > max) return false;
    return true;
  }).sort((a, b) => {
    const da = toISO(a.fecha) || '';
    const db = toISO(b.fecha) || '';
    return da.localeCompare(db);
  });
}

function fmtMoneyNeg(n) {
  // El monto viene negativo en el Excel (líneas de devolución), pero se
  // muestra en positivo (en rojo vía la clase "rej") — igual que en el
  // resto del panel, así no hay que leer signos.
  return fmtMoney(-(Number(n) || 0));
}
function codCliente(r) {
  return `${r.codclte || ''}${r.domic || ''}` || '';
}

function getKnownChoferes() {
  const map = {};
  DOC_ROWS_CACHE.forEach(r => { if (r._chofer && r._chofer.codcho) map[r._chofer.codcho] = r._chofer.nomcho || r._chofer.codcho; });
  Object.values(LAST_DOC_TO_CHOFER).forEach(h => { if (h.codcho && !map[h.codcho]) map[h.codcho] = h.nomcho || h.codcho; });
  return Object.keys(map).sort((a, b) => map[a].localeCompare(map[b])).map(cod => ({ cod, nombre: map[cod] }));
}

function buildChoferEditor(pde, currentCod, currentNombre, onSaved, choferesList) {
  const select = document.createElement('select');
  select.className = 'motivo-select';
  select.appendChild(new Option('— Sin identificar —', ''));
  const choferes = (choferesList || getKnownChoferes()).slice();
  if (currentCod && !choferes.some(c => c.cod === currentCod)) choferes.push({ cod: currentCod, nombre: currentNombre });
  choferes.forEach(c => select.appendChild(new Option(c.nombre, c.cod)));
  select.value = currentCod || '';
  select.addEventListener('change', async () => {
    const chosen = choferes.find(c => c.cod === select.value);
    await saveChoferOverride(pde, select.value || null, chosen ? chosen.nombre : null);
    (onSaved || drawDocumentos)();
  });
  return select;
}

function getKnownMotivos() {
  return Array.from(new Set(DOC_ROWS_CACHE.map(r => r._chofer && r._chofer.desmot).filter(Boolean))).sort();
}

function buildMotivoEditor(pde, currentValue) {
  const wrap = el('div');
  const select = document.createElement('select');
  select.className = 'motivo-select';
  select.appendChild(new Option('— Sin motivo —', ''));
  getKnownMotivos().forEach(m => select.appendChild(new Option(m, m)));
  const esConocido = !currentValue || getKnownMotivos().includes(currentValue);
  select.appendChild(new Option('Otro (escribir)...', '__otro__'));
  select.value = esConocido ? currentValue : '__otro__';

  const otroInput = document.createElement('input');
  otroInput.type = 'text';
  otroInput.className = 'motivo-input';
  otroInput.placeholder = 'Escribe el motivo...';
  otroInput.value = esConocido ? '' : currentValue;
  otroInput.style.display = esConocido ? 'none' : '';

  select.addEventListener('change', async () => {
    if (select.value === '__otro__') {
      otroInput.style.display = '';
      otroInput.focus();
      return;
    }
    otroInput.style.display = 'none';
    await saveMotivoOverride(pde, select.value);
    drawDocumentos();
  });
  otroInput.addEventListener('change', async (e) => {
    await saveMotivoOverride(pde, e.target.value);
    drawDocumentos();
  });

  wrap.appendChild(select);
  wrap.appendChild(otroInput);
  return wrap;
}

function drawDocumentos() {
  const rows = filteredDocRows();
  const sum = rows.reduce((s, r) => s + (Number(r.soles) || 0), 0);
  const avg = rows.length ? sum / rows.length : 0;
  const clientes = new Set(rows.map(r => r.nombrecliente)).size;

  const kpis = document.getElementById('doc-kpis');
  kpis.innerHTML = '';
  kpis.appendChild(kpiCard('navy', 'Documentos (filtro)', rows.length));
  kpis.appendChild(kpiCard('red', 'Suma total', fmtMoneyNeg(sum)));
  kpis.appendChild(kpiCard('teal', 'Promedio por documento', fmtMoneyNeg(avg)));
  kpis.appendChild(kpiCard('navy', 'Clientes únicos', clientes));

  document.getElementById('doc-result-label').textContent = `${rows.length} documento(s) encontrado(s) · Ventas (devoluciones)`;

  const pageSize = 8;
  const maxPage = Math.max(0, Math.ceil(rows.length / pageSize) - 1);
  state.doc.page = Math.min(state.doc.page, maxPage);
  const pageRows = rows.slice(state.doc.page * pageSize, state.doc.page * pageSize + pageSize);

  const table = document.getElementById('doc-table');
  table.querySelector('thead').innerHTML = `<tr>
    <th>Fecha</th><th>PDE</th><th>Código</th><th>Cliente</th>
    <th>Vendedor</th><th>Chofer</th><th>Motivo</th><th style="text-align:right">Monto</th>
  </tr>`;
  const tbody = table.querySelector('tbody');
  tbody.innerHTML = '';
  pageRows.forEach(r => {
    const chofer = r._chofer ? (r._chofer.nomcho || r._chofer.codcho) : null;
    const tr = el('tr');
    tr.appendChild(el('td', 'mono muted', fmtFecha(r.fecha)));
    tr.appendChild(el('td', 'mono', (r._chofer && r._chofer.pde) || ''));
    tr.appendChild(el('td', 'mono muted', codCliente(r)));
    tr.appendChild(el('td', null, r.nombrecliente || ''));
    tr.appendChild(el('td', 'mono muted', vendedorKey(r) === 'OFICINA' ? 'Vendedor Oficina' : (VENDOR_NAMES[r.vendedor] || r.nombrevendedor || ('Vendedor ' + r.vendedor))));
    const choferTd = el('td', 'mono muted', chofer || '<span class="muted" style="font-style:italic">sin identificar</span>');
    tr.appendChild(choferTd);
    const motivoTd = el('td');
    const pde = r._chofer && r._chofer.pde;
    const yaTieneMotivoReal = !!(r._chofer && r._chofer.desmotOriginal);
    const yaCorregido = !!(pde && MOTIVO_OVERRIDES[pde]);
    if (pde && !yaTieneMotivoReal && !yaCorregido) {
      motivoTd.appendChild(buildMotivoEditor(pde, (r._chofer && r._chofer.desmot) || ''));
    } else {
      motivoTd.className = 'muted';
      motivoTd.textContent = (r._chofer && r._chofer.desmot) || '';
      if (yaCorregido) motivoTd.title = 'Corregido a mano — ya no se puede editar desde aquí.';
    }
    tr.appendChild(motivoTd);
    tr.appendChild(el('td', 'num rej', fmtMoneyNeg(Number(r.soles) || 0)));
    tbody.appendChild(tr);
  });

  document.getElementById('docPageLabel').textContent = `${state.doc.page + 1} / ${maxPage + 1}`;
  document.getElementById('docPrev').disabled = state.doc.page === 0;
  document.getElementById('docNext').disabled = state.doc.page === maxPage;
}

async function exportDocumentosCSV() {
  const rows = filteredDocRows();
  const columns = [
    { header: 'Fecha', width: 13 },
    { header: 'PDE', width: 10, isCode: true, align: 'center' },
    { header: 'Código', width: 14, isCode: true, align: 'center' },
    { header: 'Cliente', width: 32 },
    { header: 'Vendedor', width: 26 },
    { header: 'Chofer', width: 26 },
    { header: 'Motivo', width: 20 },
    { header: 'Monto', width: 14, numFmt: '#,##0', align: 'right', red: true },
  ];
  const dataRows = rows.map(r => {
    const chofer = r._chofer ? (r._chofer.nomcho || r._chofer.codcho) : 'Sin identificar';
    return [
      fmtFecha(r.fecha),
      (r._chofer && r._chofer.pde) || '',
      codCliente(r),
      r.nombrecliente,
      vendedorKey(r) === 'OFICINA' ? 'Vendedor Oficina' : (VENDOR_NAMES[r.vendedor] || r.vendedor),
      chofer,
      (r._chofer && r._chofer.desmot) || '',
      -(Number(r.soles) || 0),
    ];
  });
  const sum = rows.reduce((s, r) => s + -(Number(r.soles) || 0), 0);
  const totalsRow = ['Total', '', '', '', '', '', `${rows.length} documento(s)`, sum];

  await downloadStyledXlsx({
    filename: `documentos_rechazados_${state.doc.from || 'todos'}.xlsx`,
    sheetName: 'Documentos',
    title: 'DOCUMENTOS RECHAZADOS',
    subtitle: `Saga Trans Confitería · ${state.doc.from} al ${state.doc.to}`,
    columns,
    rows: dataRows,
    totalsRow,
  });
}

/* ---------------- tabs / render general ---------------- */

function setTab(tab) {
  state.tab = tab;
  document.querySelectorAll('.tabbar .tab-btn').forEach(b => b.classList.toggle('tab-active', b.dataset.tab === tab));
  ['general', 'conductor', 'vendedor', 'motivo', 'documentos', 'pdes'].forEach(t => {
    document.getElementById('panel-' + t).style.display = (t === tab) ? '' : 'none';
  });
  renderAll();
}

function renderAll() {
  const bannerHayDatos = CORTE_DATA.ventas.length;
  const banner = document.getElementById('pendienteBanner');
  document.getElementById('corteLabel').textContent = getCorteLabel();

  if (state.tab !== 'documentos' && state.tab !== 'motivo' && state.tab !== 'pdes' && !bannerHayDatos) {
    banner.style.display = 'flex';
    banner.textContent = `Todavía no hay datos cargados para ${getCorteLabel()}. Sube los 3 Excel de ese mes a Storage y este panel se completa solo.`;
  } else {
    banner.style.display = 'none';
  }

  if (state.tab === 'general') renderGeneral();
  else if (state.tab === 'conductor') renderConductor();
  else if (state.tab === 'vendedor') renderVendedor();
  else if (state.tab === 'motivo') renderMotivo();
  else if (state.tab === 'documentos') drawDocumentos();
  else if (state.tab === 'pdes') renderPdes();
}

/* ---------------- Choferes por PDE (mantenimiento) ---------------- */

let PDE_SEARCH = '';
let PDE_PAGE = 0;

function renderPdes() {
  // Uno por PDE (varios documentos comparten el mismo PDE = mismo reparto),
  // así se puede corregir el chofer aunque ese PDE no tenga ningún rechazo.
  const byPde = {};
  Object.values(LAST_DOC_TO_CHOFER).forEach(hit => {
    if (!hit.pde) return;
    if (!byPde[hit.pde]) byPde[hit.pde] = hit;
  });
  // Aplica correcciones ya guardadas.
  Object.keys(byPde).forEach(pde => {
    if (CHOFER_OVERRIDES[pde]) {
      byPde[pde] = { ...byPde[pde], codcho: CHOFER_OVERRIDES[pde].codcho, nomcho: CHOFER_OVERRIDES[pde].nomcho, choferCorregido: true };
    }
  });

  const q = PDE_SEARCH.toLowerCase();
  const allEntries = Object.values(byPde);

  // Vista fija: solo el chofer que sabemos que está mal asignado. Si más
  // adelante hay otro caso, se cambia aquí (HUAMANI_TARGET_REGEX).
  const choferMap = {};
  allEntries.forEach(h => { if (h.codcho) choferMap[h.codcho] = h.nomcho || h.codcho; });
  const huamaniCod = Object.keys(choferMap).find(c => /HUAMANI/i.test(choferMap[c]) && /JUAN/i.test(choferMap[c]));
  const badge = document.getElementById('pdeChoferFixed');
  badge.textContent = huamaniCod ? choferMap[huamaniCod] : 'No encontrado en este corte';

  const rows = allEntries
    .filter(h => huamaniCod ? h.codcho === huamaniCod : false)
    .filter(h => !q || String(h.pde).toLowerCase().includes(q))
    .sort((a, b) => String(a.pde).localeCompare(String(b.pde)));

  const pageSize = 30;
  const maxPage = Math.max(0, Math.ceil(rows.length / pageSize) - 1);
  PDE_PAGE = Math.min(PDE_PAGE, maxPage);
  const pageRows = rows.slice(PDE_PAGE * pageSize, PDE_PAGE * pageSize + pageSize);

  document.getElementById('pdeResultLabel').textContent = `${rows.length} PDE encontrado(s) de ${Object.keys(byPde).length}`;
  document.getElementById('pdePageLabel').textContent = `${PDE_PAGE + 1} / ${maxPage + 1}`;
  document.getElementById('pdePrev').disabled = PDE_PAGE === 0;
  document.getElementById('pdeNext').disabled = PDE_PAGE === maxPage;

  const choferesList = getKnownChoferes(); // se calcula UNA vez para toda la página, no por fila

  const table = document.getElementById('pde-table');
  table.querySelector('thead').innerHTML = `<tr><th>PDE</th><th>Chofer</th></tr>`;
  const tbody = table.querySelector('tbody');
  tbody.innerHTML = '';
  pageRows.forEach(h => {
    const tr = el('tr');
    tr.appendChild(el('td', 'mono', h.pde));
    const td = el('td');
    if (h.choferCorregido) {
      td.className = 'mono muted';
      td.textContent = h.nomcho || h.codcho || '';
      td.title = 'Corregido a mano — ya no se puede editar desde aquí.';
    } else {
      td.appendChild(buildChoferEditor(h.pde, h.codcho, h.nomcho, renderPdes, choferesList));
    }
    tr.appendChild(td);
    tbody.appendChild(tr);
  });
}

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
  document.getElementById('vendUmbral').addEventListener('input', e => { state.vend.umbral = parseFloat(e.target.value) || 0; document.getElementById('vendAlertLabel').textContent = state.vend.onlyAlerts ? 'Ver todos' : `Solo alertas ≥${state.vend.umbral}%`; renderVendedor(); });
  document.getElementById('vendAlertToggle').addEventListener('click', () => {
    state.vend.onlyAlerts = !state.vend.onlyAlerts;
    document.getElementById('vendAlertLabel').textContent = state.vend.onlyAlerts ? 'Ver todos' : `Solo alertas ≥${state.vend.umbral}%`;
    renderVendedor();
  });

  document.getElementById('docSearch').addEventListener('input', e => { state.doc.search = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docVendor').addEventListener('change', e => { state.doc.vendor = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docChofer').addEventListener('change', e => { state.doc.chofer = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docMotivo').addEventListener('change', e => { state.doc.motivo = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docMontoMin').addEventListener('input', e => { state.doc.montoMin = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docMontoMax').addEventListener('input', e => { state.doc.montoMax = e.target.value; state.doc.page = 0; drawDocumentos(); });
  document.getElementById('docFechaFrom').addEventListener('change', async e => { state.doc.from = e.target.value; await renderDocumentos(); });
  document.getElementById('docFechaTo').addEventListener('change', async e => { state.doc.to = e.target.value; await renderDocumentos(); });
  document.getElementById('docClear').addEventListener('click', async () => {
    state.doc.search = ''; state.doc.vendor = 'all'; state.doc.chofer = 'all'; state.doc.motivo = 'all';
    state.doc.montoMin = ''; state.doc.montoMax = ''; state.doc.page = 0;
    document.getElementById('docSearch').value = '';
    document.getElementById('docMontoMin').value = '';
    document.getElementById('docMontoMax').value = '';
    await renderDocumentos();
  });
  document.getElementById('docPrev').addEventListener('click', () => { state.doc.page = Math.max(0, state.doc.page - 1); drawDocumentos(); });
  document.getElementById('docNext').addEventListener('click', () => { state.doc.page += 1; drawDocumentos(); });
  document.getElementById('docExportCSV').addEventListener('click', exportDocumentosCSV);
  document.getElementById('exportGeneralCSV').addEventListener('click', exportGeneralCSV);
  document.getElementById('exportMotivoCSV').addEventListener('click', exportMotivoCSV);
  document.getElementById('exportConductorCSV').addEventListener('click', exportConductorCSV);
  document.getElementById('exportVendedorCSV').addEventListener('click', exportVendedorCSV);
  document.getElementById('pdeSearch').addEventListener('input', e => { PDE_SEARCH = e.target.value; PDE_PAGE = 0; renderPdes(); });
  document.getElementById('pdePrev').addEventListener('click', () => { PDE_PAGE = Math.max(0, PDE_PAGE - 1); renderPdes(); });
  document.getElementById('pdeNext').addEventListener('click', () => { PDE_PAGE += 1; renderPdes(); });
}

/* ---------------- arranque ---------------- */

async function bootApp() {
  initCorteControls();
  wireEvents();
  await loadMotivoOverrides();
  await loadChoferOverrides();
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
const userEmailLabel = document.getElementById('userEmailLabel');
const userAvatar = document.getElementById('userAvatar');

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
    await supabaseClient.auth.signOut();
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

async function showApp(session) {
  loginScreen.style.display = 'none';
  appRoot.style.display = '';
  if (userEmailLabel) userEmailLabel.textContent = session?.user?.email || '';
  if (userAvatar) userAvatar.textContent = (session?.user?.email || '?').trim().charAt(0).toUpperCase();
  sessionActive = true;
  resetInactivityTimer();

  if (!appBooted) {
    appBooted = true;
    const overlay = document.getElementById('loadingOverlay');
    try {
      if (typeof XLSX === 'undefined') throw new Error('SheetJS (XLSX) no cargó — revisa el <script> de xlsx en index.html.');
      await bootApp();
    } catch (err) {
      console.error(err);
      document.querySelector('.wrap').innerHTML =
        `<div class="empty-state"><div class="empty-title">No se pudo cargar el panel</div><div class="empty-text">${err.message}</div></div>`;
    } finally {
      if (overlay) overlay.style.display = 'none';
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
supabaseClient.auth.getSession().then(({ data: { session } }) => {
  if (session) { showApp(session); } else { showLogin(); }
}).catch(err => {
  console.error('Error revisando sesión existente:', err);
  showLogin();
});

supabaseClient.auth.onAuthStateChange((event, session) => {
  if (event === 'SIGNED_IN' && session) showApp(session);
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
    const { error } = await supabaseClient.auth.signInWithPassword({ email, password });
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
    await supabaseClient.auth.signOut();
  } catch (err) {
    console.error('Error al cerrar sesión:', err);
  }
});
