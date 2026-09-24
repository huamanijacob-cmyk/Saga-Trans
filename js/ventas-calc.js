// =====================================================================
// Módulo Ventas — motor de cálculo (js/ventas-calc.js) (sin DOM, sin Supabase).
// Todas las funciones reciben datos ya leídos del Excel y devuelven
// objetos listos para dibujar. Se puede probar en Node.
// =====================================================================
(function (root) {
  'use strict';
  if (typeof console !== 'undefined') console.log('Módulo Ventas — ventas-calc.js version 6 (neto > 0 · meta del día = pendientes)');

  // ---------------- Saneo de datos crudos ----------------
  // Quita \r\n, espacios y convierte a texto. NUNCA usar !!valor para
  // decidir si un campo está vacío: algunos traen "\r\n   ".
  function txt(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/[\r\n\t]+/g, ' ').trim();
  }
  // Código con ceros a la izquierda ('1' → '00001'), por si Excel lo guardó como número.
  function code(v, len) {
    const s = txt(v);
    if (!s) return '';
    return /^\d+$/.test(s) && s.length < len ? s.padStart(len, '0') : s;
  }
  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const s = txt(v).replace(/,/g, '');
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }
  // Normaliza cualquier fecha a 'AAAA/MM/DD' (texto 2026/09/23, 2026-09-23, 23/09/2026, Date o serial de Excel).
  function fecha(v) {
    if (v === null || v === undefined || v === '') return '';
    const pad = n => String(n).padStart(2, '0');
    if (v instanceof Date && !isNaN(v)) {
      return `${v.getFullYear()}/${pad(v.getMonth() + 1)}/${pad(v.getDate())}`;
    }
    if (typeof v === 'number') { // serial de Excel (días desde 1899-12-30)
      const d = new Date(Math.round((v - 25569) * 86400000));
      return `${d.getUTCFullYear()}/${pad(d.getUTCMonth() + 1)}/${pad(d.getUTCDate())}`;
    }
    const s = txt(v);
    let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m) return `${m[1]}/${pad(m[2])}/${pad(m[3])}`;
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) return `${m[3]}/${pad(m[2])}/${pad(m[1])}`;
    return s;
  }
  // Busca una columna sin importar mayúsculas/espacios en el encabezado.
  function pick(row, name) {
    if (name in row) return row[name];
    const want = name.toLowerCase().replace(/\s+/g, '');
    for (const k of Object.keys(row)) {
      if (k.toLowerCase().replace(/\s+/g, '') === want) return row[k];
    }
    return undefined;
  }
  function cliKey(codclte, domic) { return `${codclte}-${domic}`; }
  function docNumber(r) { return `${r.tipo}${r.serie}${r.doc}`; }

  // ---------------- Lectura de cada archivo ----------------
  function parseVentas(rows) {
    return rows.map(r => {
      const o = {
        tipo: txt(pick(r, 'tipo')),
        serie: txt(pick(r, 'serie')),
        doc: txt(pick(r, 'doc')),
        codclte: code(pick(r, 'codclte'), 6),
        domic: code(pick(r, 'domic'), 3),
        nombrecliente: txt(pick(r, 'nombrecliente')),
        vendedor: txt(pick(r, 'vendedor')),
        nombrevendedor: txt(pick(r, 'nombrevendedor')),
        cat: code(pick(r, 'codigocategoria'), 5),
        catNombre: txt(pick(r, 'DescripcionCategoria')),
        producto: txt(pick(r, 'descrpcionproducto')),
        soles: num(pick(r, 'soles')),
        vtadvo: txt(pick(r, 'vtadvo')).toUpperCase(),
        fecha: fecha(pick(r, 'fecha')),
      };
      o.cli = cliKey(o.codclte, o.domic);
      o.documento = docNumber(o);
      return o;
    });
  }

  // Cuotas en formato ancho: Codigo | Vendedor | Cuota Chocolate | Cuota Galleta | Cuota Paneton | Cuota Total
  function parseCuotas(rows, categorias) {
    const out = {};
    const avisos = [];
    rows.forEach((r, i) => {
      const cod = txt(pick(r, 'Codigo'));
      if (!cod) return;
      const porCat = {};
      let suma = 0;
      categorias.forEach(c => {
        const v = pick(r, c.columnaCuota);
        if (v === undefined) avisos.push(`Cuotas: no existe la columna "${c.columnaCuota}".`);
        porCat[c.codigo] = num(v);
        suma += porCat[c.codigo];
      });
      const totalExcel = pick(r, 'Cuota Total');
      if (totalExcel !== undefined && Math.abs(num(totalExcel) - suma) > 0.5) {
        avisos.push(`Cuotas fila ${i + 2} (${cod}): "Cuota Total" (${num(totalExcel)}) no es la suma de las categorías (${suma}). Se usa la suma.`);
      }
      out[cod] = { vendedor: cod, nombre: txt(pick(r, 'Vendedor')), porCat };
    });
    return { cuotas: out, avisos: [...new Set(avisos)] };
  }

  // Cartera: vendedor viene como "99999 VENTAS EN OFICINA".
  // Un cliente-local cuenta solo si estado_cliente = AC y estado_domicilio = AC.
  function parseCartera(rows) {
    const todos = {};           // cli → {vendedor, activo, ...} (incluye IN)
    const activosPorVend = {};  // vendedor → Set(cli)
    rows.forEach(r => {
      const cod = code(pick(r, 'codigo_cliente'), 6);
      if (!cod) return;
      const dom = code(pick(r, 'domicilio'), 3);
      const vendRaw = txt(pick(r, 'vendedor'));
      const vendedor = vendRaw.split(/\s+/)[0] || '';
      const estCli = txt(pick(r, 'estado_cliente')).toUpperCase();
      const estDom = txt(pick(r, 'estado_domicilio')).toUpperCase();
      const k = cliKey(cod, dom);
      const activo = estCli === 'AC' && estDom === 'AC';
      todos[k] = { cli: k, vendedor, nombre: txt(pick(r, 'nombre_cliente')), estCli, estDom, activo };
      if (activo) (activosPorVend[vendedor] = activosPorVend[vendedor] || new Set()).add(k);
    });
    return { todos, activosPorVend };
  }

  // Calendario: solo excepciones. fecha | trabaja (SI/NO) | motivo
  function parseCalendario(rows) {
    const ex = {};
    rows.forEach(r => {
      const f = fecha(pick(r, 'fecha'));
      const t = txt(pick(r, 'trabaja')).toUpperCase().replace('SÍ', 'SI');
      if (!f || (t !== 'SI' && t !== 'NO')) return;
      ex[f] = { trabaja: t === 'SI', motivo: txt(pick(r, 'motivo')) };
    });
    return ex;
  }

  // ---------------- Días hábiles ----------------
  // Regla normal: lunes a sábado se trabaja, domingo no. El calendario manda si tiene la fecha.
  function diasDelMes(ym, excepciones) {
    const [y, m] = ym.split('-').map(Number);
    const n = new Date(y, m, 0).getDate();
    const pad = x => String(x).padStart(2, '0');
    const dias = [];
    for (let d = 1; d <= n; d++) {
      const f = `${y}/${pad(m)}/${pad(d)}`;
      const dow = new Date(y, m - 1, d).getDay();
      const ex = excepciones && excepciones[f];
      const trabaja = ex ? ex.trabaja : dow !== 0;
      dias.push({ fecha: f, dia: d, dow, trabaja, motivo: ex ? ex.motivo : (dow === 0 ? 'Domingo' : ''), excepcion: !!ex });
    }
    return dias;
  }
  // Transcurridos: hábiles hasta la fecha de corte INCLUIDA.
  // Restantes: hábiles DESPUÉS del corte (preventa: lo que se vende hoy se factura mañana).
  function habiles(ym, excepciones, corte) {
    const dias = diasDelMes(ym, excepciones);
    const hab = dias.filter(d => d.trabaja);
    const trans = hab.filter(d => d.fecha <= corte).length;
    return { dias, total: hab.length, transcurridos: trans, restantes: hab.length - trans };
  }

  function estado(proy, cfg) {
    if (proy === null || !isFinite(proy)) return null;
    if (proy >= cfg.UMBRAL_EN_RITMO) return 'ritmo';
    if (proy >= cfg.UMBRAL_ATENCION) return 'atencion';
    return 'riesgo';
  }
  function div(a, b) { return b > 0 ? a / b : null; }

  // ---------------- Agregados base ----------------
  // Índice: vendedor → cat → {mes (≤corte), antes (<corte), hoy (=corte)} en soles netos (V + D).
  function indexar(ventas, corte, catCodes) {
    const catSet = new Set(catCodes);
    const idx = {};
    const nombres = {};
    ventas.forEach(r => {
      if (!catSet.has(r.cat) || !r.fecha || r.fecha > corte) return;
      nombres[r.vendedor] = nombres[r.vendedor] || r.nombrevendedor;
      const v = (idx[r.vendedor] = idx[r.vendedor] || {});
      const c = (v[r.cat] = v[r.cat] || { mes: 0, antes: 0, hoy: 0 });
      c.mes += r.soles;
      if (r.fecha < corte) c.antes += r.soles; else c.hoy += r.soles;
    });
    return { idx, nombres };
  }

  // ---------------- 1. Avance general ----------------
  function avanceGeneral({ ventas, cuotas, corte, hab, cats, meta, cfg }) {
    const { idx, nombres } = indexar(ventas, corte, cats);
    const vends = new Set([...Object.keys(cuotas), ...Object.keys(idx)]);
    const factor = hab.total > 0 ? hab.transcurridos / hab.total : 0;
    const filas = [...vends].map(v => {
      const q100 = cats.reduce((s, c) => s + ((cuotas[v] && cuotas[v].porCat[c]) || 0), 0);
      const fact = cats.reduce((s, c) => s + ((idx[v] && idx[v][c] && idx[v][c].mes) || 0), 0);
      const cuota = q100 * meta;
      const pct = div(fact, cuota);
      const proy = pct !== null && factor > 0 ? pct / factor : null;
      const sinRk = cfg.VENDEDORES_SIN_RANKING.includes(v) || !(cuota > 0);
      return {
        vendedor: v,
        nombre: (cuotas[v] && cuotas[v].nombre) || nombres[v] || v,
        cuota100: q100, cuota, facturado: fact,
        pct: sinRk ? null : pct, proy: sinRk ? null : proy,
        estado: sinRk ? null : estado(proy, cfg), sinRanking: sinRk,
      };
    });
    const rk = filas.filter(f => !f.sinRanking).sort((a, b) => b.pct - a.pct);
    rk.forEach((f, i) => { f.ranking = i + 1; });
    const otros = filas.filter(f => f.sinRanking).sort((a, b) => a.vendedor.localeCompare(b.vendedor));
    const tot = filas.reduce((t, f) => ({ cuota100: t.cuota100 + f.cuota100, cuota: t.cuota + f.cuota, facturado: t.facturado + f.facturado }),
      { cuota100: 0, cuota: 0, facturado: 0 });
    tot.pct = div(tot.facturado, tot.cuota);
    tot.proy = tot.pct !== null && factor > 0 ? tot.pct / factor : null;
    tot.estado = estado(tot.proy, cfg);
    return { filas: [...rk, ...otros], total: tot };
  }

  // ---------------- 2. Cobertura por categoría ----------------
  function cobertura({ ventas, cuotas, cartera, corte, hab, cat, meta, cfg }) {
    const { idx, nombres } = indexar(ventas, corte, [cat]);
    // Regla de cobertura: NETO > 0. Un cliente está cubierto si (ventas − devoluciones)
    // de la categoría, acumulado hasta la fecha, es mayor que cero. Si le devolvieron todo, no cuenta.
    // netoAntes = hasta el día anterior al corte · netoMes = hasta el corte (incluido).
    const netoAntes = {}, netoMes = {};
    ventas.forEach(r => {
      if (r.cat !== cat || !r.fecha || r.fecha > corte) return;
      const vm = (netoMes[r.vendedor] = netoMes[r.vendedor] || {});
      vm[r.cli] = (vm[r.cli] || 0) + r.soles;
      if (r.fecha < corte) {
        const va = (netoAntes[r.vendedor] = netoAntes[r.vendedor] || {});
        va[r.cli] = (va[r.cli] || 0) + r.soles;
      }
    });
    const positivos = m => new Set(Object.keys(m || {}).filter(k => m[k] > 0.01));
    const vends = new Set([...Object.keys(cuotas), ...Object.keys(cartera.activosPorVend)]);
    const filas = [];
    [...vends].sort().forEach(v => {
      if (cfg.VENDEDORES_SIN_RANKING.includes(v)) return;
      const cart = cartera.activosPorVend[v] || new Set();
      const enCart = s => new Set([...s].filter(k => cart.has(k)));
      const cubAntes = enCart(positivos(netoAntes[v]));   // cubiertos hasta ayer
      const cubMes = enCart(positivos(netoMes[v]));       // cubiertos hasta el corte
      const nuevosHoy = [...cubMes].filter(k => !cubAntes.has(k));
      const q = ((cuotas[v] && cuotas[v].porCat[cat]) || 0) * meta;
      const q100 = (cuotas[v] && cuotas[v].porCat[cat]) || 0;
      const c = (idx[v] && idx[v][cat]) || { mes: 0, antes: 0, hoy: 0 };
      const pendientes = cart.size - cubAntes.size;
      // Cuota diaria = lo que falta ÷ días hábiles restantes (después del corte).
      // Si ya cumplió la cuota del mes, se muestra la cuota diaria "normal"
      // (cuota ÷ días hábiles del mes) para que igual se vea cuánto hizo hoy.
      // El último día (0 restantes) se divide entre 1: lo que falta es para ese día.
      let cuotaDiaria = null, cuotaCumplida = false;
      if (q > 0) {
        const falta = q - c.antes;
        if (falta <= 0) { cuotaCumplida = true; cuotaDiaria = hab.total > 0 ? q / hab.total : null; }
        else cuotaDiaria = falta / Math.max(hab.restantes, 1);
      }
      filas.push({
        vendedor: v, nombre: (cuotas[v] && cuotas[v].nombre) || nombres[v] || v,
        mes: {
          clientes: cart.size, conVenta: cubMes.size, pctCob: div(cubMes.size, cart.size),
          cuota100: q100, cuota: q, avance: c.mes, pct: div(c.mes, q),
        },
        dia: {
          // Meta de clientes del día = los pendientes. % cobertura del día = venta ÷ pendientes.
          pendientes, nuevos: nuevosHoy.length, pctCob: div(nuevosHoy.length, pendientes),
          cuotaDiaria, cuotaCumplida, avance: c.hoy,
          pct: cuotaDiaria ? c.hoy / cuotaDiaria : null,
        },
      });
    });
    const s = (f) => filas.reduce((a, x) => a + f(x), 0);
    const totMes = { clientes: s(x => x.mes.clientes), conVenta: s(x => x.mes.conVenta), cuota100: s(x => x.mes.cuota100), cuota: s(x => x.mes.cuota), avance: s(x => x.mes.avance) };
    totMes.pctCob = div(totMes.conVenta, totMes.clientes); totMes.pct = div(totMes.avance, totMes.cuota);
    const cdVals = filas.map(x => x.dia.cuotaDiaria);
    const totDia = { pendientes: s(x => x.dia.pendientes), nuevos: s(x => x.dia.nuevos), avance: s(x => x.dia.avance),
      cuotaDiaria: cdVals.some(v => v === null) ? null : cdVals.reduce((a, b) => a + b, 0) };
    totDia.pctCob = div(totDia.nuevos, totDia.pendientes);
    totDia.pct = totDia.cuotaDiaria ? totDia.avance / totDia.cuotaDiaria : null;
    return { filas, total: { mes: totMes, dia: totDia } };
  }

  // ---------------- 3. Alertas: ventas fuera de cartera ----------------
  function fueraDeCartera({ ventas, cartera, corte, catCodes, cfg }) {
    const catSet = new Set(catCodes);
    const out = {};
    ventas.forEach(r => {
      if (!catSet.has(r.cat) || r.vtadvo !== 'V' || !r.fecha || r.fecha > corte) return;
      if (cfg.VENDEDORES_SIN_RANKING.includes(r.vendedor)) return;
      const cart = cartera.activosPorVend[r.vendedor];
      if (cart && cart.has(r.cli)) return;
      const info = cartera.todos[r.cli];
      let motivo = 'No está en la cartera';
      if (info && !info.activo) motivo = `Inactivo (cliente ${info.estCli || '—'} / local ${info.estDom || '—'})`;
      else if (info && info.vendedor !== r.vendedor) motivo = `En cartera del vendedor ${info.vendedor}`;
      const k = `${r.vendedor}|${r.cli}`;
      const o = (out[k] = out[k] || { vendedor: r.vendedor, nombreVendedor: r.nombrevendedor, cli: r.cli, nombre: r.nombrecliente, motivo, soles: 0, docs: new Set(), ultima: '' });
      o.soles += r.soles; o.docs.add(r.documento); if (r.fecha > o.ultima) o.ultima = r.fecha;
    });
    return Object.values(out).map(o => ({ ...o, docs: o.docs.size })).sort((a, b) => a.vendedor.localeCompare(b.vendedor) || b.soles - a.soles);
  }

  // ---------------- 4. Drill-down: documentos detrás de una cifra ----------------
  // filtro: { vendedor, cats:[...], desde, hasta } (fechas 'AAAA/MM/DD', inclusivas)
  // filtro.vendedor = null → todos · filtro.cats = lista de códigos permitidos (obligatoria)
  // filtro.buscar = texto libre (documento, código de cliente o nombre)
  function documentos(ventas, filtro) {
    const catSet = new Set(filtro.cats);
    const q = txt(filtro.buscar).toLowerCase();
    const out = {};
    ventas.forEach(r => {
      if (filtro.vendedor && r.vendedor !== filtro.vendedor) return;
      if (!catSet.has(r.cat)) return;
      if (filtro.cli && r.cli !== filtro.cli) return;
      if (!r.fecha || r.fecha < filtro.desde || r.fecha > filtro.hasta) return;
      if (q && !(r.documento.toLowerCase().includes(q) || r.cli.includes(q) || r.nombrecliente.toLowerCase().includes(q))) return;
      const k = `${r.documento}|${r.fecha}`;
      const o = (out[k] = out[k] || { fecha: r.fecha, documento: r.documento, tipo: r.tipo, vendedor: r.vendedor, nombreVendedor: r.nombrevendedor, codigo: r.cli, cliente: r.nombrecliente, cats: new Set(), venta: 0, devolucion: 0, lineas: 0 });
      o.cats.add(r.catNombre);
      if (r.vtadvo === 'D') o.devolucion += r.soles; else o.venta += r.soles;
      o.lineas++;
    });
    return Object.values(out)
      .map(o => ({ ...o, cats: [...o.cats].join(', '), neto: o.venta + o.devolucion }))
      .sort((a, b) => b.fecha.localeCompare(a.fecha) || a.documento.localeCompare(b.documento));
  }

  function fechasDisponibles(ventas) {
    return [...new Set(ventas.map(r => r.fecha).filter(Boolean))].sort();
  }

  // ---------------- 5. Comparativo por DÍA DE VENTA ----------------
  // Día de venta N = N-ésima fecha con facturación en el Excel del mes (de toda la
  // empresa, sin filtros), así domingos, feriados y días no trabajados se saltan solos.
  // actual se corta en la fecha de corte; previo se toma completo.
  function comparativo({ actual, previo, cats, vendedor, corte }) {
    const cs = new Set(cats);
    const serie = (rows, fechas) => {
      const m = Object.fromEntries(fechas.map(f => [f, 0]));
      rows.forEach(r => {
        if (!cs.has(r.cat) || (vendedor && r.vendedor !== vendedor) || !(r.fecha in m)) return;
        m[r.fecha] += r.soles;
      });
      return fechas.map(f => m[f]);
    };
    const fa = fechasDisponibles(actual).filter(f => !corte || f <= corte);
    const fp = fechasDisponibles(previo);
    const va = serie(actual, fa), vp = serie(previo, fp);
    const dias = [];
    let sa = 0, sp = 0;
    for (let i = 0; i < Math.max(fa.length, fp.length); i++) {
      const hayA = i < fa.length, hayP = i < fp.length;
      if (hayA) sa += va[i];
      if (hayP) sp += vp[i];
      dias.push({
        n: i + 1,
        fechaAct: hayA ? fa[i] : null, ventaAct: hayA ? va[i] : null, acumAct: hayA ? sa : null,
        fechaPrev: hayP ? fp[i] : null, ventaPrev: hayP ? vp[i] : null, acumPrev: hayP ? sp : null,
      });
    }
    const nAct = fa.length;
    const prevMismoDia = nAct && nAct <= fp.length ? dias[nAct - 1].acumPrev : (fp.length ? sp : null);
    return {
      dias, nAct, nPrev: fp.length, acumAct: sa, prevMismoDia, prevMes: sp,
      variacion: prevMismoDia ? sa / prevMismoDia - 1 : null,
    };
  }

  const VentasCalc = { txt, code, num, fecha, parseVentas, parseCuotas, parseCartera, parseCalendario,
    diasDelMes, habiles, avanceGeneral, cobertura, fueraDeCartera, documentos, fechasDisponibles, comparativo, estado };
  if (typeof module !== 'undefined' && module.exports) module.exports = VentasCalc;
  else root.VentasCalc = VentasCalc;
})(typeof window !== 'undefined' ? window : this);
