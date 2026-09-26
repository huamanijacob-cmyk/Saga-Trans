/* =====================================================================
   Módulo Notas de crédito — motor de cálculo (js/nc-calc.js)
   Sin pantalla ni Supabase: recibe filas ya leídas y devuelve resultados.
   Reglas validadas con los cierres de agosto y setiembre 2026:
   · DPR = descuento por ítem · CO = vigente · AN = anulada (monto 0)
   · Serie con N (BN01, FN01) la asume Nestlé; con S (BS01, FS01), Saga Trans (ST)
   · Sell out = venta neta de Chocolate + Galleta, sin IGV (mismo "Avance" de Ventas)
   · %DCTOS ST = descuentos ST con IGV ÷ sell out (sin IGV y con IGV)
   ===================================================================== */
(function (root) {
  'use strict';
  const IGV = 1.18;
  const CAT_SELLOUT = ['00001', '00002'];
  const CAT_PANETON = '00003';

  const txt = v => (v === null || v === undefined ? '' : String(v).replace(/[\r\n\t]+/g, ' ').trim());
  const num = v => { if (typeof v === 'number') return isFinite(v) ? v : 0; const n = parseFloat(txt(v).replace(/,/g, '')); return isFinite(n) ? n : 0; };
  const pad = (s, n) => (/^\d+$/.test(s) && s.length < n ? s.padStart(n, '0') : s);
  function fecha(v) {
    if (v === null || v === undefined || v === '') return '';
    const p = n => String(n).padStart(2, '0');
    if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}/${p(v.getMonth() + 1)}/${p(v.getDate())}`;
    if (typeof v === 'number') { const d = new Date(Math.round((v - 25569) * 86400000)); return `${d.getUTCFullYear()}/${p(d.getUTCMonth() + 1)}/${p(d.getUTCDate())}`; }
    const s = txt(v);
    let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/); if (m) return `${m[1]}/${p(m[2])}/${p(m[3])}`;
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/); if (m) return `${m[3]}/${p(m[2])}/${p(m[1])}`;
    return s;
  }
  const r2 = n => Math.round(n * 100) / 100;

  // "BO    B002-00130268" → { tipo:'BO', serie:'B002', num:'00130268', doc:'BOB00200130268', n:130268 }
  function parseRef(ref) {
    const m = txt(ref).match(/^(\w{2})\s+(\w{4})-(\d+)/);
    if (!m) return null;
    return { tipo: m[1], serie: m[2], num: m[3], doc: `${m[1]}${m[2]}${m[3]}`, n: parseInt(m[3], 10) };
  }

  // ---------- Notas de crédito ----------
  function parseNC(rows) {
    return rows.map(r => {
      const sersun = txt(r.sersun), numsun = txt(r.numsun);
      const ref = parseRef(r.referencia);
      const c = sersun.charAt(1).toUpperCase();
      return {
        tipped: txt(r.tipped).toUpperCase(), estdoc: txt(r.estdoc).toUpperCase(),
        fecha: fecha(r.emision), fecanu: fecha(r.fecanu),
        sersun, numsun, doc: `${txt(r.coddoc) || 'NC'}${sersun}${numsun}`,
        codcli: pad(txt(r.codcli), 6), razsoc: txt(r.razsoc),
        presub: num(r.presub), mondoc: num(r.mondoc), vendedor: txt(r.vendedor),
        motivo: txt(r.motivo), ref, refDoc: ref ? ref.doc : '',
        cargo: c === 'N' ? 'N' : c === 'S' ? 'S' : '',
      };
    });
  }

  // ---------- Facturas (de Ventas): monto, categoría principal, cliente ----------
  function indexFacturas(ventas) {
    const fac = {};      // doc → {doc, tipo, n, fecha, cli, nombre, venta (s/IGV), cats:{cat:soles}}
    const porNumero = {}; // "BO|130341" → doc (el archivo de Alex no trae serie)
    const conDevolucion = new Set();
    ventas.forEach(r => {
      const doc = `${txt(r.tipo)}${txt(r.serie)}${txt(r.doc)}`;
      if (!doc) return;
      if (txt(r.vtadvo).toUpperCase() === 'D') { conDevolucion.add(doc); return; }
      const f = (fac[doc] = fac[doc] || { doc, tipo: txt(r.tipo), n: parseInt(txt(r.doc), 10), fecha: fecha(r.fecha), cli: pad(txt(r.codclte), 6) + pad(txt(r.domic), 3), nombre: txt(r.nombrecliente), venta: 0, cats: {} });
      const s = num(r.soles), cat = pad(txt(r.codigocategoria), 5);
      f.venta += s; f.cats[cat] = (f.cats[cat] || 0) + s;
      porNumero[`${f.tipo}|${f.n}`] = doc;
    });
    Object.values(fac).forEach(f => { f.cat = Object.entries(f.cats).sort((a, b) => b[1] - a[1])[0]?.[0] || ''; });
    return { fac, porNumero, conDevolucion };
  }

  function sellOut(ventas, desde, hasta) {
    const catSet = new Set(CAT_SELLOUT); const porDia = {};
    ventas.forEach(r => {
      const f = fecha(r.fecha), cat = pad(txt(r.codigocategoria), 5);
      if (!catSet.has(cat) || !f || f < desde || f > hasta) return;
      porDia[f] = (porDia[f] || 0) + num(r.soles);
    });
    return porDia;
  }

  // ---------- Archivo de Alex ----------
  // Recibe { nombreHoja: filas (arrays) }. Usa las hojas con encabezado "CodClie".
  // Si existe una hoja con columna "Observacion" (p. ej. "X DIA"), de ahí salen RECHAZO/PARCIAL.
  // Las hojas por día aportan el producto/promoción (columna O). Se evita contar dos veces.
  function parseAlex(hojas) {
    const filas = {}; const obs = {}; const prod = {};
    Object.entries(hojas).forEach(([nombre, rows]) => {
      if (!rows.length || txt(rows[0][0]).toLowerCase() !== 'codclie') return;
      // Columnas por su título: "Observación" (RECHAZO/PARCIAL) y "Producto". En las hojas por día
      // de Alex la columna O no tiene título y trae el producto.
      const head = rows[0].map(h => txt(h).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''));
      let colObs = head.findIndex(h => h.startsWith('observ'));
      let colProd = head.findIndex(h => h.startsWith('producto'));
      if (colObs < 0 && colProd < 0) colProd = 14;
      const esConsolidado = colObs >= 0 && colProd < 0;
      for (let i = 2; i < rows.length; i++) {
        const r = rows[i]; const tipo = txt(r[4]).toUpperCase(); const n = parseInt(txt(r[5]), 10);
        if (!tipo || !n) continue;
        const key = `${tipo}|${n}`;
        const o = colObs >= 0 ? txt(r[colObs]).toUpperCase() : '';
        const pr = colProd >= 0 ? txt(r[colProd]) : '';
        if (o) obs[key] = o;
        if (pr) prod[key] = pr;
        const cliRaw = txt(r[0]).replace(/\D/g, '');
        const fila = {
          key, tipo, n, cli: cliRaw ? pad(cliRaw, 9) : '', razon: txt(r[1]), fecha: fecha(r[2]), vendedor: txt(r[3]),
          monto: num(r[6]), dpct: num(r[7]), dmonto: num(r[8]), npct: num(r[10]), nmonto: num(r[11]),
          spct: r[12] === null || r[12] === undefined || r[12] === '' ? num(r[7]) - num(r[10]) : num(r[12]),
          smonto: r[13] === null || r[13] === undefined || r[13] === '' ? num(r[8]) - num(r[11]) : num(r[13]),
          hoja: nombre, consolidado: esConsolidado,
        };
        // Una sola fila por factura: la de la hoja diaria tiene prioridad (trae producto).
        if (!filas[key] || (filas[key].consolidado && !esConsolidado)) filas[key] = fila;
      }
    });
    const lista = Object.values(filas);
    lista.forEach(f => { f.obs = obs[f.key] || ''; f.producto = prod[f.key] || ''; f.smonto = r2(f.smonto); });
    return lista;
  }

  // ---------- Cálculo principal del mes ----------
  // opts: { nc, ventas, alex, ajustes, registro, corte, ym }
  //   ajustes = { marcas: {docNC: {tipo:'adicional'|'regular'|'rechazo'|'parcial'|'error', etiqueta, monto?}}, manuales: [{concepto, cargo, monto}], correo: {fecha: monto} }
  //   registro = { docNC: { m, s, ref, cli, f, cargo } } montos vistos mientras la NC estaba vigente
  function calcular({ nc, ventas, alex, ajustes, registro, corte, ym }) {
    const pref = ym.replace('-', '/');
    const desde = `${pref}/01`;
    const marcas = (ajustes && ajustes.marcas) || {};
    const manuales = (ajustes && ajustes.manuales) || [];
    const correo = (ajustes && ajustes.correo) || {};
    const correoS = (ajustes && ajustes.correoST) || {};
    const { fac, porNumero, conDevolucion } = indexFacturas(ventas);
    const alexKey = {};
    (alex || []).forEach(a => { a.doc = porNumero[a.key] || `${a.tipo}${a.n}`; alexKey[a.doc] = a; });

    const dpr = nc.filter(r => r.tipped === 'DPR' && r.fecha >= desde && r.fecha <= corte);
    const vig = dpr.filter(r => r.estdoc === 'CO');
    const anu = dpr.filter(r => r.estdoc === 'AN');

    // Clasificación de cada NC vigente: regular / adicional / panetón
    vig.forEach(r => {
      const m = marcas[r.doc];
      const f = fac[r.refDoc];
      r.factura = f || null;
      r.producto = (alexKey[r.refDoc] && alexKey[r.refDoc].producto) || '';
      if (m && m.tipo === 'adicional') { r.grupo = 'adicional'; r.etiqueta = m.etiqueta || ''; }
      else if (f && f.cat === CAT_PANETON) r.grupo = 'paneton';
      else r.grupo = 'regular';
    });

    const suma = (arr, cargo, campo) => arr.filter(r => !cargo || r.cargo === cargo).reduce((s, r) => s + r[campo], 0);
    const reg = vig.filter(r => r.grupo === 'regular');
    const ajN = manuales.filter(a => a.cargo === 'N').reduce((s, a) => s + num(a.monto), 0);
    const ajS = manuales.filter(a => a.cargo === 'S').reduce((s, a) => s + num(a.monto), 0);
    const tot = {
      N: suma(reg, 'N', 'mondoc') + ajN, S: suma(reg, 'S', 'mondoc') + ajS,
      Ns: suma(reg, 'N', 'presub') + ajN / IGV, Ss: suma(reg, 'S', 'presub') + ajS / IGV,
      ajN, ajS,
      adicional: suma(vig.filter(r => r.grupo === 'adicional'), null, 'mondoc'),
      paneton: suma(vig.filter(r => r.grupo === 'paneton'), null, 'mondoc'),
    };
    tot.total = tot.N + tot.S; tot.totals = tot.Ns + tot.Ss;
    const soDia = sellOut(ventas, desde, corte);
    tot.sellout = Object.values(soDia).reduce((a, b) => a + b, 0);
    tot.pctST = tot.sellout ? tot.S / tot.sellout : null;
    tot.pctSTc = tot.sellout ? tot.S / (tot.sellout * IGV) : null;
    tot.shareN = tot.total ? tot.N / tot.total : null; tot.shareS = tot.total ? tot.S / tot.total : null;

    // Curva diaria: % ST acumulado (misma fórmula que la tarjeta) en cada día con venta
    const stDia = {};
    reg.filter(r => r.cargo === 'S').forEach(r => { stDia[r.fecha] = (stDia[r.fecha] || 0) + r.mondoc; });
    let aS = 0, aV = 0; const curva = [];
    Object.keys({ ...soDia, ...stDia }).sort().forEach(f => {
      aS += stDia[f] || 0; aV += soDia[f] || 0;
      if (soDia[f]) curva.push({ fecha: f, pct: aV ? aS / aV : null });
    });

    // Diario por serie
    const diario = {};
    vig.forEach(r => {
      const d = (diario[r.fecha] = diario[r.fecha] || { fecha: r.fecha, c: {}, s: {} });
      d.c[r.sersun] = (d.c[r.sersun] || 0) + r.mondoc; d.s[r.sersun] = (d.s[r.sersun] || 0) + r.presub;
    });
    const series = [...new Set(vig.map(r => r.sersun))].sort((a, b) => (a.charAt(1) + a).localeCompare(b.charAt(1) + b));

    // Estado de cada factura de Alex frente a las NC vigentes. Si Alex no la marcó, igual se detecta
    // rechazo/parcial porque la factura tuvo devolución (D en Ventas o nota de devolución PDP).
    nc.filter(r => r.tipped === 'PDP' && r.refDoc).forEach(r => conDevolucion.add(r.refDoc));
    const sis = {};
    vig.forEach(r => { if (!r.refDoc) return; const x = (sis[r.refDoc] = sis[r.refDoc] || { N: 0, S: 0, ncs: [] }); x[r.cargo] += r.mondoc; x.ncs.push(r.doc); });
    (alex || []).forEach(a => {
      const x = sis[a.doc], dev = conDevolucion.has(a.doc);
      const iguales = x && Math.abs(a.nmonto - x.N) <= 0.05 && Math.abs(a.smonto - x.S) <= 0.05;
      a.auto = false;
      if (a.obs === 'RECHAZO') a.estado = 'rechazo';
      else if (a.obs === 'PARCIAL') a.estado = 'parcial';
      else if (!x && !fac[a.doc]) { a.estado = 'facanulada'; a.auto = true; }   // la factura ya no está en Ventas: se anuló
      else if (!x) { a.estado = dev ? 'rechazo' : 'falta'; a.auto = dev; }
      else if (iguales) a.estado = 'coincide';
      else if (Math.abs((a.nmonto + a.smonto) - (x.N + x.S)) <= 0.05) a.estado = 'reparto';   // mismo total, distinto quién asume
      else { a.estado = dev ? 'parcial' : 'distinto'; a.auto = dev; }
    });

    // Anuladas: monto y clasificación
    const alexRech = (alex || []).filter(a => a.estado === 'rechazo' || a.estado === 'parcial' || a.estado === 'facanulada')
      .map(a => ({ ...a, obs: a.estado === 'parcial' ? 'PARCIAL' : 'RECHAZO' }));
    const grupos = {};
    anu.forEach(r => { const k = `${r.codcli}|${r.fecha}|${r.cargo}`; (grupos[k] = grupos[k] || []).push(r); });
    Object.entries(grupos).forEach(([k, arr]) => {
      const [cli, f, cargo] = k.split('|');
      const cand = alexRech.filter(a => a.cli.slice(0, 6) === cli && a.fecha === f && (cargo === 'N' ? a.nmonto : a.smonto) > 0)
        .map(a => ({ obs: a.obs, monto: cargo === 'N' ? a.nmonto : a.smonto, doc: a.doc })).sort((x, y) => y.monto - x.monto);
      arr.forEach((r, i) => { r.alexCand = cand[i] || null; });
    });
    anu.forEach(r => {
      const reg0 = registro && registro[r.doc];
      const man = marcas[r.doc];
      r.montoRegistro = reg0 ? reg0.m : null;
      r.refDocRegistro = reg0 ? reg0.ref : '';
      let clase = 'error', monto = 0, origen = '';
      if (r.alexCand) { clase = r.alexCand.obs === 'PARCIAL' ? 'parcial' : 'rechazo'; monto = r.alexCand.monto; origen = 'alex'; }
      else if (reg0 && reg0.ref && conDevolucion.has(reg0.ref)) { clase = 'rechazo'; monto = reg0.m; origen = 'registro'; }
      if (reg0 && clase !== 'error' && !monto) { monto = reg0.m; origen = 'registro'; }
      if (man && ['rechazo', 'parcial', 'error'].includes(man.tipo)) {
        clase = man.tipo; origen = 'manual';
        if (man.tipo === 'error') monto = 0;
        else if (man.monto !== undefined && man.monto !== null && man.monto !== '') monto = num(man.monto);
        else if (!monto && reg0) monto = reg0.m;
      }
      r.clase = clase; r.montoAnulada = r2(monto); r.origen = origen;
    });
    const rech = { N: 0, S: 0 };
    anu.forEach(r => { if (r.clase === 'rechazo' || r.clase === 'parcial') rech[r.cargo] += r.montoAnulada; });
    tot.rechN = r2(rech.N); tot.rechS = r2(rech.S);

    // Conciliación con Alex (por factura)
    let concil = null;
    if (alex && alex.length) {
      const filas = []; const vistos = new Set();
      const finAlex = alex.map(a => a.fecha).sort().slice(-1)[0];   // último día que cubre el archivo de Alex
      alex.filter(a => a.fecha >= desde && a.fecha <= corte).forEach(a => {
        vistos.add(a.doc);
        const s = sis[a.doc];
        filas.push({ doc: a.doc, fecha: a.fecha, cli: a.cli, razon: a.razon, producto: a.producto, alexN: a.nmonto, alexS: a.smonto, sisN: s ? r2(s.N) : 0, sisS: s ? r2(s.S) : 0, dN: r2(a.nmonto - (s ? s.N : 0)), dS: r2(a.smonto - (s ? s.S : 0)), estado: a.estado, auto: a.auto, ncs: s ? s.ncs : [] });
      });
      Object.entries(sis).forEach(([doc, s]) => {
        if (vistos.has(doc)) return;
        const f = fac[doc];
        const fch = f ? f.fecha : (vig.find(r => r.refDoc === doc) || {}).fecha || '';
        if (fch > finAlex) return;   // todavía no llega en el archivo de Alex: no es una diferencia
        filas.push({ doc, fecha: fch, cli: f ? f.cli : '', razon: f ? f.nombre : '', producto: '', alexN: 0, alexS: 0, sisN: r2(s.N), sisS: r2(s.S), dN: r2(-s.N), dS: r2(-s.S), estado: 'sinalex', ncs: s.ncs });
      });
      const cuenta = {}; filas.forEach(f => { cuenta[f.estado] = (cuenta[f.estado] || 0) + 1; });
      // Tabla "SR. ALEX": Nestlé diario según Alex, acumulado, lo que informa por correo y diferencia
      const porDia = {}, porDiaS = {};
      alex.filter(a => a.fecha >= desde && a.fecha <= corte).forEach(a => { porDia[a.fecha] = (porDia[a.fecha] || 0) + a.nmonto; porDiaS[a.fecha] = (porDiaS[a.fecha] || 0) + a.smonto; });
      let acc = 0, accS = 0;
      const vacio = v => v === undefined || v === null || v === '';
      const srAlex = Object.keys(porDia).sort().map(f => {
        acc += porDia[f]; accS += porDiaS[f] || 0; const c = correo[f], cs = correoS[f];
        return { fecha: f, diario: r2(porDia[f]), acumulado: r2(acc), correo: vacio(c) ? null : num(c), dif: vacio(c) ? null : r2(num(c) - acc),
          diarioS: r2(porDiaS[f] || 0), acumuladoS: r2(accS), correoS: vacio(cs) ? null : num(cs), difS: vacio(cs) ? null : r2(num(cs) - accS) };
      });
      concil = { filas, cuenta, srAlex, finAlex };
    }

    // Por producto (de las hojas diarias de Alex)
    const prod = {};
    (alex || []).filter(a => a.producto && a.fecha >= desde && a.fecha <= corte).forEach(a => {
      const p = (prod[a.producto] = prod[a.producto] || { producto: a.producto, facturas: 0, N: 0, S: 0 });
      p.facturas++; p.N += a.nmonto; p.S += a.smonto;
    });
    const productos = Object.values(prod).map(p => ({ ...p, N: r2(p.N), S: r2(p.S), total: r2(p.N + p.S) })).sort((a, b) => b.total - a.total);

    // Documentos: por factura
    const docs = {};
    vig.forEach(r => {
      const key = r.refDoc || r.doc;
      const f = r.factura;
      const d = (docs[key] = docs[key] || { factura: r.refDoc, fecha: f ? f.fecha : r.fecha, cli: f ? f.cli : r.codcli, cliente: f ? f.nombre : r.razsoc, venta: f ? f.venta : 0, N: 0, S: 0, Ns: 0, Ss: 0, ncs: [], grupo: r.grupo, etiqueta: r.etiqueta || '', producto: r.producto, vendedor: r.vendedor });
      d[r.cargo] += r.mondoc; d[r.cargo + 's'] += r.presub; d.ncs.push(r.doc);
      if (r.grupo !== 'regular') { d.grupo = r.grupo; d.etiqueta = r.etiqueta || d.etiqueta; }
    });
    const documentos = Object.values(docs).map(d => ({
      ...d, ventaIgv: r2(d.venta * IGV),
      pctN: d.venta ? d.Ns / d.venta : null, pctS: d.venta ? d.Ss / d.venta : null, pctT: d.venta ? (d.Ns + d.Ss) / d.venta : null,
    })).sort((a, b) => b.fecha.localeCompare(a.fecha) || a.factura.localeCompare(b.factura));

    return { tot, curva, diario: Object.values(diario).sort((a, b) => a.fecha.localeCompare(b.fecha)), series, anuladas: anu, concil, productos, documentos, vigentes: vig, facturas: fac, alexPorDoc: alexKey };
  }

  // Montos de NC vigentes que todavía no están en el registro (para recordarlos si luego se anulan).
  function nuevosParaRegistro(nc, registro) {
    const out = {};
    nc.forEach(r => {
      if (r.tipped !== 'DPR' || r.estdoc !== 'CO' || !r.mondoc) return;
      const prev = registro[r.doc];
      if (!prev || prev.m !== r.mondoc) out[r.doc] = { m: r2(r.mondoc), s: r2(r.presub), ref: r.refDoc, cli: r.codcli, f: r.fecha, cargo: r.cargo };
    });
    return out;
  }

  function fechasNC(nc) { return [...new Set(nc.filter(r => r.tipped === 'DPR').map(r => r.fecha).filter(Boolean))].sort(); }

  const NCCalc = { IGV, parseNC, parseAlex, calcular, nuevosParaRegistro, fechasNC, fecha, parseRef };
  if (typeof module !== 'undefined' && module.exports) module.exports = NCCalc;
  else root.NCCalc = NCCalc;
})(typeof window !== 'undefined' ? window : this);
