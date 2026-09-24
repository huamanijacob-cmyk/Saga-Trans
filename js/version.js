// =====================================================================
// Panel de Gestión — Saga Trans Confitería
// VERSIÓN ÚNICA de todo el panel (Rechazos + Ventas).
// En cada actualización: sube este número (1.0 → 1.1 … 1.9 → 2.0)
// y el mismo número en los ?v= de index.html.
// =====================================================================
const PANEL_VERSION = '1.0';
console.log('Panel de Gestión — versión ' + PANEL_VERSION);

// "jacob.huamani@empresa.com" → { nombre: 'Jacob Huamani', iniciales: 'JH' }
function nombreDesdeCorreo(email) {
  const user = String(email || '').split('@')[0];
  const partes = user.split(/[._\-]+/).filter(Boolean);
  const cap = w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  const nombre = partes.map(cap).join(' ') || email || '';
  const iniciales = (partes.length > 1 ? partes[0].charAt(0) + partes[partes.length - 1].charAt(0) : (partes[0] || '?').charAt(0)).toUpperCase();
  return { nombre, iniciales };
}
