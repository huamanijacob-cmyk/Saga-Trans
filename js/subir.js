/* =====================================================================
   Subir archivos desde el panel (js/subir.js)
   Reconoce cada Excel por su nombre, revisa que tenga las columnas
   esperadas, avisa si reemplaza uno existente y lo sube a su carpeta.
   Reutiliza: supabaseClient · FILE_CACHE · onCorteChanged (Rechazos)
   ===================================================================== */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  if (!$('subirModal')) { console.error('Subir archivos: falta el bloque #subirModal en index.html'); return; }

  // Tipos de archivo que se pueden subir: patrón del nombre, destino y columnas obligatorias.
  const TIPOS = [
    { id: 'Ventas', nombre: 'Ventas', re: /^ventas_(\d{4})-(\d{2})/i, bucket: 'rechazos-data', carpeta: 'Ventas', cols: ['vtadvo', 'soles', 'vendedor', 'fecha', 'codclte'] },
    { id: 'ContabilidadNC', nombre: 'Notas de crédito', re: /^contabilidadnc_(\d{4})-(\d{2})/i, bucket: 'rechazos-data', carpeta: 'ContabilidadNC', cols: ['tipped', 'sersun', 'numsun', 'estdoc', 'mondoc'] },
    { id: 'Transportistas', nombre: 'Transportistas', re: /^transportistas_(\d{4})-(\d{2})/i, bucket: 'rechazos-data', carpeta: 'Transportistas', cols: ['pednum', 'codcli', 'candsp', 'totdsp'] },
    { id: 'Cuotas', nombre: 'Cuotas', re: /^cuotas_(\d{4})-(\d{2})/i, bucket: 'ventas-data', carpeta: 'Cuotas', cols: ['codigo', 'cuota chocolate', 'cuota galleta'] },
    { id: 'Cartera', nombre: 'Cartera', re: /^cartera_(\d{4})-(\d{2})/i, bucket: 'ventas-data', carpeta: 'Cartera', cols: ['codigo_cliente', 'domicilio', 'vendedor', 'estado_cliente', 'estado_domicilio'] },
    { id: 'Calendario', nombre: 'Calendario', re: /^calendario_(\d{4})(?![-\d])/i, bucket: 'ventas-data', carpeta: 'Calendario', cols: ['fecha', 'trabaja'], anual: true },
  ];
  const MES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'setiembre', 'octubre', 'noviembre', 'diciembre'];
  const LOG = 'Overrides/subidas_log.json';
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const norm = s => String(s ?? '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[\r\n]+/g, ' ').trim();
  const nf = new Intl.NumberFormat('es-PE');
  let items = [];   // archivos elegidos, ya analizados

  function tipoDe(nombre) {
    for (const t of TIPOS) {
      const m = nombre.match(t.re);
      if (m) return { t, periodo: t.anual ? m[1] : `${m[1]}-${m[2]}`, destino: t.anual ? `${t.carpeta}/${t.id}_${m[1]}.xlsx` : `${t.carpeta}/${t.id}_${m[1]}-${m[2]}.xlsx` };
    }
    return null;
  }
  const periodoTxt = (t, p) => (t.anual ? `año ${p}` : `${MES[+p.slice(5) - 1]} ${p.slice(0, 4)}`);

  async function existentes(bucket, carpeta) {
    const { data, error } = await supabaseClient.storage.from(bucket).list(carpeta, { limit: 1000 });
    if (error || !Array.isArray(data)) return new Set();
    return new Set(data.map(o => String(o.name || '').toLowerCase()));
  }

  async function analizar(files) {
    const cache = {};
    const res = [];
    for (const file of files) {
      const it = { file, estado: 'error', msg: '' };
      const tp = tipoDe(file.name);
      if (!/\.xlsx$/i.test(file.name)) { it.msg = 'Solo archivos .xlsx'; res.push(it); continue; }
      if (!tp) { it.msg = 'Nombre no reconocido'; res.push(it); continue; }
      Object.assign(it, tp);
      try {
        const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const head = (XLSX.utils.sheet_to_json(ws, { header: 1, range: 0, defval: '' })[0] || []).map(norm);
        const faltan = tp.t.cols.filter(c => !head.includes(c));
        if (faltan.length) { it.msg = `Faltan columnas: ${faltan.join(', ')}`; res.push(it); continue; }
        const rango = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        it.filas = Math.max(0, rango.e.r);
        const key = `${tp.t.bucket}/${tp.t.carpeta}`;
        cache[key] = cache[key] || await existentes(tp.t.bucket, tp.t.carpeta);
        it.reemplaza = cache[key].has(tp.destino.split('/')[1].toLowerCase());
        it.estado = 'ok';
      } catch (e) { it.msg = 'No se pudo leer el archivo'; console.error(e); }
      res.push(it);
    }
    // Si se eligió dos veces el mismo destino, solo vale el último.
    const vistos = new Set();
    for (let i = res.length - 1; i >= 0; i--) {
      const it = res[i]; if (it.estado !== 'ok') continue;
      if (vistos.has(it.t.bucket + it.destino)) { it.estado = 'error'; it.msg = 'Repetido: se usa el último elegido'; }
      vistos.add(it.t.bucket + it.destino);
    }
    return res;
  }

  function dibujar() {
    const ok = items.filter(i => i.estado === 'ok');
    $('subirTabla').innerHTML = items.length ? `<table class="v-table"><thead><tr><th>Archivo</th><th>Tipo</th><th>Periodo</th><th>Se guarda en</th><th class="r">Filas</th><th class="c">Estado</th></tr></thead><tbody>${items.map(i => `<tr>
      <td class="mono">${esc(i.file.name)}</td><td>${i.t ? esc(i.t.nombre) : '—'}</td><td>${i.t ? esc(periodoTxt(i.t, i.periodo)) : '—'}</td>
      <td class="mono">${i.t ? esc(`${i.t.bucket}/${i.t.carpeta}/`) : '—'}</td><td class="num">${i.filas !== undefined ? nf.format(i.filas) : '—'}</td>
      <td class="c">${i.estado === 'ok' ? `<span class="sub-chip ${i.reemplaza ? 'rep' : 'new'}">${i.reemplaza ? 'Reemplaza el existente' : 'Nuevo'}</span>`
        : i.estado === 'subido' ? '<span class="sub-chip new">Subido ✓</span>' : `<span class="sub-chip bad">${esc(i.msg)}</span>`}</td></tr>`).join('')}</tbody></table>` : '';
    const b = $('subirConfirmar');
    b.style.display = ok.length ? '' : 'none';
    b.textContent = `Subir ${ok.length} archivo${ok.length === 1 ? '' : 's'}`;
  }

  async function elegir(files) {
    if (!files || !files.length) return;
    $('subirMsg').textContent = 'Revisando archivos...';
    items = await analizar([...files]);
    $('subirMsg').textContent = '';
    dibujar();
  }

  async function registrar(entradas) {
    try {
      const { data } = await supabaseClient.auth.getSession();
      const quien = data?.session?.user?.email || '';
      const s = await supabaseClient.storage.from('ventas-data').download(LOG).catch(() => null);
      let log = [];
      if (s && s.data) { try { log = JSON.parse(await s.data.text()); } catch { log = []; } }
      const ahora = new Date().toISOString();
      entradas.forEach(e => log.push({ ...e, usuario: quien, fecha: ahora }));
      await supabaseClient.storage.from('ventas-data').upload(LOG, new Blob([JSON.stringify(log.slice(-500))], { type: 'application/json' }), { upsert: true, contentType: 'application/json', cacheControl: '0' });
    } catch (e) { console.warn('No se pudo registrar la subida:', e); }
  }

  async function subir() {
    const ok = items.filter(i => i.estado === 'ok');
    if (!ok.length) return;
    const b = $('subirConfirmar'); b.disabled = true;
    const hechos = [];
    for (const it of ok) {
      $('subirMsg').textContent = `Subiendo ${it.file.name}...`;
      const tipo = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      const cuerpo = new Blob([await it.file.arrayBuffer()], { type: tipo });   // el archivo completo en memoria antes de enviarlo
      const { error } = await supabaseClient.storage.from(it.t.bucket).upload(it.destino, cuerpo, { upsert: true, contentType: tipo, cacheControl: '0' });
      if (error) { it.estado = 'error'; it.msg = `No se pudo subir: ${error.message}`; console.error(error); }
      else {
        it.estado = 'subido'; hechos.push({ archivo: it.file.name, destino: `${it.t.bucket}/${it.destino}`, filas: it.filas });
        if (typeof FILE_CACHE !== 'undefined') delete FILE_CACHE[it.destino];
      }
      dibujar();
    }
    b.disabled = false;
    if (!hechos.length) { $('subirMsg').textContent = 'No se subió ningún archivo. Revisa los permisos (SQL de subida).'; return; }
    await registrar(hechos);
    $('subirMsg').textContent = 'Actualizando los datos del panel...';
    try {
      if (typeof onCorteChanged === 'function') await onCorteChanged();
      if (window.VENTAS && VENTAS.state && VENTAS.state.booted && $('vBtnRecargar')) $('vBtnRecargar').click();
      if (window.NCMOD && NCMOD.state && NCMOD.state.booted && $('ncBtnRecargar')) $('ncBtnRecargar').click();
    } catch (e) { console.error(e); }
    $('subirMsg').textContent = `Listo: ${hechos.length} archivo${hechos.length === 1 ? '' : 's'} subido${hechos.length === 1 ? '' : 's'}.`;
  }

  function abrir() { items = []; dibujar(); $('subirMsg').textContent = ''; $('subirModal').style.display = 'flex'; }
  function cerrar() { $('subirModal').style.display = 'none'; }

  document.querySelectorAll('[data-abrir-subir]').forEach(b => b.addEventListener('click', abrir));
  $('subirCerrar').addEventListener('click', cerrar);
  $('subirCancelar').addEventListener('click', cerrar);
  $('subirModal').addEventListener('click', e => { if (e.target.id === 'subirModal') cerrar(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('subirModal').style.display === 'flex') cerrar(); });
  $('subirConfirmar').addEventListener('click', subir);
  const drop = $('subirDrop'), inp = $('subirInput');
  drop.addEventListener('click', () => inp.click());
  inp.addEventListener('change', e => { elegir(e.target.files); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('on'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('on'); }));
  drop.addEventListener('drop', e => elegir(e.dataTransfer.files));
})();
