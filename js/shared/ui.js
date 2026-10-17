/**
 * NEXOVA Suite · shared/ui.js
 * Utilidades de interfaz compartidas por todas las herramientas:
 * formato, descargas, avisos (toast), confirmación modal e íconos SVG.
 * Reemplaza alert()/confirm() nativos: no bloquean la página y respetan la marca.
 */

export const $ = (id) => document.getElementById(id);
export const num = (n) => (+n || 0).toLocaleString('es-PE');
export const pct = (n) => (+n || 0).toFixed(1) + '%';
export const esc = (t) => String(t ?? '').replace(/[<>&"']/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&#39;' }[c]));

// ── Íconos (trazo 1.8, estilo línea; se colorean con currentColor) ──
const P = {
  camara: '<rect x="3" y="6.5" width="18" height="13" rx="2.5"/><circle cx="12" cy="13" r="3.6"/><path d="M8.5 6.5 10 4h4l1.5 2.5"/>',
  datos: '<ellipse cx="12" cy="5.5" rx="7.5" ry="2.8"/><path d="M4.5 5.5v6c0 1.55 3.36 2.8 7.5 2.8s7.5-1.25 7.5-2.8v-6"/><path d="M4.5 11.5v6c0 1.55 3.36 2.8 7.5 2.8s7.5-1.25 7.5-2.8v-6"/>',
  proyecto: '<rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 3.5h6a1 1 0 0 1 1 1V6a1 1 0 0 1-1 1H9a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1z"/><path d="M8.5 13l2 2 4-4.2"/>',
  entrega: '<path d="M9 3.5h5l4 4V17a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2V5.5a2 2 0 0 1 2-2z"/><path d="M14 3.5V8h4"/><path d="M6 7.5V19a2.5 2.5 0 0 0 2.5 2.5H16"/>',
  correo: '<rect x="3" y="5.5" width="18" height="13" rx="2.5"/><path d="M3.5 7l8.5 5.5L20.5 7"/>',
  formato: '<rect x="4.5" y="3.5" width="15" height="17" rx="2"/><path d="M8 8.5h8M8 12h8M8 15.5h5"/>',
  excel: '<path d="M14 3.5H7a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8.5z"/><path d="M14 3.5v5h5"/><path d="M9.5 12l5 5M14.5 12l-5 5"/>',
  archivo: '<path d="M14 3.5H7a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8.5z"/><path d="M14 3.5v5h5"/><path d="M9 13h6M9 16.5h4"/>',
  carpeta: '<path d="M3.5 7a2 2 0 0 1 2-2h4l2 2.5h7a2 2 0 0 1 2 2V17a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2z"/>',
  carpetaDest: '<path d="M3.5 7a2 2 0 0 1 2-2h4l2 2.5h7a2 2 0 0 1 2 2V17a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2z"/><path d="M12 10.5v5M9.5 13l2.5 2.5 2.5-2.5"/>',
  subir: '<path d="M12 15.5V4.5M7.5 9 12 4.5 16.5 9"/><path d="M4.5 15v3a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2v-3"/>',
  bajar: '<path d="M12 4.5v11M7.5 11l4.5 4.5 4.5-4.5"/><path d="M4.5 15v3a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2v-3"/>',
  ok: '<circle cx="12" cy="12" r="9"/><path d="M8 12.3l2.7 2.7L16 9.6"/>',
  alerta: '<path d="M12 4 2.8 19.5h18.4z"/><path d="M12 10v4.2M12 17h.01"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5.5M12 7.8h.01"/>',
  error: '<circle cx="12" cy="12" r="9"/><path d="M9 9l6 6M15 9l-6 6"/>',
  ojo: '<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="3"/>',
  rayo: '<path d="M13 3 5 13.5h6L10 21l8-10.5h-6z"/>',
  flecha: '<path d="M5 12h14M13 6l6 6-6 6"/>',
  atras: '<path d="M19 12H5M11 6l-6 6 6 6"/>',
  abajo: '<path d="M6 9l6 6 6-6"/>',
  candado: '<rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
  cerrar: '<path d="M6 6l12 12M18 6 6 18"/>',
  limpiar: '<path d="M4 20h16"/><path d="M14.5 4.5l5 5-8 8H6.5v-5z"/>',
  reiniciar: '<path d="M4.5 12a7.5 7.5 0 1 0 2.2-5.3"/><path d="M4.5 4.5v4h4"/>',
  presentacion: '<rect x="3" y="4" width="18" height="12" rx="2"/><path d="M12 16v4M8 20h8"/><path d="M7.5 12.5 10.5 9.5l2.5 2.5 3.5-3.5"/>',
  pdf: '<path d="M14 3.5H7a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8.5z"/><path d="M14 3.5v5h5"/><path d="M8.5 16.5v-4h1.5a1.25 1.25 0 0 1 0 2.5H8.5"/>',
  word: '<path d="M14 3.5H7a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8.5z"/><path d="M14 3.5v5h5"/><path d="M8.5 12l1.2 5 1.8-4 1.8 4 1.2-5"/>',
  // Dimensiones del depurador
  familia: '<rect x="3.5" y="4" width="7" height="7" rx="1.5"/><rect x="13.5" y="4" width="7" height="7" rx="1.5"/><rect x="3.5" y="14" width="7" height="7" rx="1.5"/><rect x="13.5" y="14" width="7" height="7" rx="1.5"/>',
  subfamilia: '<path d="M3.5 12.2V5a1.5 1.5 0 0 1 1.5-1.5h7.2l8.3 8.3-8.7 8.7z"/><circle cx="8" cy="8" r="1.4"/>',
  area: '<rect x="4" y="3.5" width="16" height="17" rx="1.5"/><path d="M8 7.5h2M14 7.5h2M8 11.5h2M14 11.5h2M10 20.5v-4h4v4"/>',
  sede: '<path d="M12 21s-6.5-5.6-6.5-11a6.5 6.5 0 0 1 13 0c0 5.4-6.5 11-6.5 11z"/><circle cx="12" cy="10" r="2.4"/>',
  centro: '<circle cx="12" cy="12" r="8.5"/><path d="M14.8 9.2c-.5-.9-1.6-1.5-2.8-1.5-1.6 0-2.8.9-2.8 2.1 0 2.8 5.6 1.4 5.6 4.3 0 1.2-1.2 2.1-2.8 2.1-1.3 0-2.4-.6-2.9-1.6M12 6v1.7M12 16.3V18"/>',
  responsable: '<circle cx="12" cy="8" r="3.8"/><path d="M4.5 20.5c.8-3.8 3.9-6 7.5-6s6.7 2.2 7.5 6"/>',
  estado: '<path d="M3.5 12h4l2-5 4 10 2-5h5"/>',
  ubicacion: '<path d="M9 4 3.5 6v14L9 18l6 2 5.5-2V4L15 6z"/><path d="M9 4v14M15 6v14"/>',
  linea: '<path d="M3.5 20.5V10l5 3V10l5 3V6.5h7v14z"/><path d="M7 17h2M12 17h2M17 17h1"/>',
  activos: '<path d="M12 3 20 7.5v9L12 21 4 16.5v-9z"/><path d="M4 7.5 12 12l8-4.5M12 12v9"/>',
  filtro: '<path d="M4 5h16l-6.2 7.4V19l-3.6 1.5v-8.1z"/>',
  paleta: '<path d="M12 3.5a8.5 8.5 0 1 0 0 17c1.1 0 1.7-.8 1.7-1.7 0-1.1-.9-1.4-.9-2.4 0-.9.7-1.6 1.6-1.6h2a4.1 4.1 0 0 0 4.1-4.1c0-4-3.8-7.2-8.5-7.2z"/><circle cx="7.8" cy="11" r="1.1"/><circle cx="10.5" cy="7.4" r="1.1"/><circle cx="15" cy="7.8" r="1.1"/>',
  trofeo: '<path d="M8 4h8v5a4 4 0 0 1-8 0z"/><path d="M8 5.5H4.5a3 3 0 0 0 3.6 3.9M16 5.5h3.5a3 3 0 0 1-3.6 3.9M12 13v4M8.5 20.5h7M9.5 17h5v3.5h-5z"/>',
  hex: '<path d="M12 2.8 20 7.4v9.2l-8 4.6-8-4.6V7.4z"/>',
};

/** SVG en línea de un ícono del set (tamaño por CSS: .ico). */
export function icono(nombre, clase = '') {
  const d = P[nombre] || P.hex;
  return `<svg class="ico ${clase}" viewBox="0 0 24 24" aria-hidden="true" focusable="false">${d}</svg>`;
}

// ── Avisos (toasts) ────────────────────────────────────────
function pilaToasts() {
  let el = document.getElementById('toasts');
  if (!el) {
    el = document.createElement('div');
    el.id = 'toasts';
    el.className = 'toasts';
    el.setAttribute('role', 'status');
    el.setAttribute('aria-live', 'polite');
    document.body.appendChild(el);
  }
  return el;
}

/**
 * Muestra un aviso no bloqueante.
 * @param {string} texto  Mensaje (texto plano).
 * @param {'info'|'ok'|'warn'|'error'} tipo
 * @param {{titulo?:string, ms?:number}} opts
 */
export function toast(texto, tipo = 'info', opts = {}) {
  const ic = { info: 'info', ok: 'ok', warn: 'alerta', error: 'error' }[tipo] || 'info';
  const t = document.createElement('div');
  t.className = `toast ${tipo}`;
  t.innerHTML = `${icono(ic)}<div class="toast-txt">${opts.titulo ? `<b>${esc(opts.titulo)}</b>` : ''}<span>${esc(texto).replace(/\n/g, '<br>')}</span></div>
    <button class="toast-x" aria-label="Cerrar aviso">${icono('cerrar')}</button>`;
  const cerrar = () => { t.classList.add('sale'); setTimeout(() => t.remove(), 220); };
  t.querySelector('.toast-x').onclick = cerrar;
  pilaToasts().appendChild(t);
  const ms = opts.ms ?? (tipo === 'error' ? 9000 : tipo === 'warn' ? 7000 : 4500);
  if (ms > 0) setTimeout(cerrar, ms);
  return cerrar;
}

/**
 * Confirmación modal accesible (reemplaza confirm()).
 * @returns {Promise<boolean>}
 */
export function confirmar({ titulo, texto, aceptar = 'Continuar', cancelar = 'Cancelar', peligro = false }) {
  return new Promise((resolver) => {
    const previo = document.activeElement;
    const m = document.createElement('div');
    m.className = 'nx-modal';
    m.innerHTML = `<div class="nx-dialog" role="alertdialog" aria-modal="true" aria-labelledby="nxdT" aria-describedby="nxdD">
      <div class="nx-dialog-ic ${peligro ? 'warn' : ''}">${icono(peligro ? 'alerta' : 'info')}</div>
      <h3 id="nxdT">${esc(titulo)}</h3>
      <p id="nxdD">${esc(texto).replace(/\n/g, '<br>')}</p>
      <div class="nx-dialog-btns">
        <button class="b-sec" data-r="0">${esc(cancelar)}</button>
        <button class="b-pri" data-r="1">${esc(aceptar)}</button>
      </div></div>`;
    const fin = (r) => {
      m.remove();
      document.removeEventListener('keydown', tecla, true);
      previo?.focus?.();
      resolver(r);
    };
    const tecla = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); fin(false); }
      if (e.key === 'Tab') {   // foco atrapado en el diálogo
        const bs = [...m.querySelectorAll('button')];
        const i = bs.indexOf(document.activeElement);
        e.preventDefault();
        bs[(i + (e.shiftKey ? -1 : 1) + bs.length) % bs.length].focus();
      }
    };
    m.addEventListener('click', (e) => {
      const b = e.target.closest('button[data-r]');
      if (b) fin(b.dataset.r === '1');
      else if (e.target === m) fin(false);
    });
    document.addEventListener('keydown', tecla, true);
    document.body.appendChild(m);
    m.querySelector('[data-r="1"]').focus();
  });
}

/** Descarga un Blob con nombre de archivo. */
export function descargar(blob, nombre) {
  if (!blob) { toast('Aún no hay nada que descargar. Corra la simulación primero.', 'warn'); return; }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = nombre; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Zona de carga: clic, teclado (Enter/Espacio) y arrastrar-soltar. */
export function zonaArchivo(id, alSoltar, alActivar) {
  const z = $(id);
  if (!z.hasAttribute('tabindex')) z.tabIndex = 0;
  z.setAttribute('role', 'button');
  z.onclick = alActivar;
  z.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); alActivar(); } });
  if (!alSoltar) return;
  z.addEventListener('dragover', (e) => { e.preventDefault(); z.classList.add('drag'); });
  z.addEventListener('dragleave', () => z.classList.remove('drag'));
  z.addEventListener('drop', (e) => {
    e.preventDefault(); z.classList.remove('drag');
    const f = e.dataTransfer.files[0]; if (f) alSoltar(f);
  });
}

/** Marca una zona de carga como lista (ícono de check + título + detalle). */
export function zonaLista(id, titulo, detalle, aviso = false) {
  const z = $(id);
  z.classList.add('ok');
  z.classList.toggle('aviso', aviso);
  z.querySelector('.ic').innerHTML = icono(aviso ? 'alerta' : 'ok');
  z.querySelector('.zt').textContent = titulo;
  z.querySelector('.zs').textContent = detalle;
}

/** Menú desplegable simple (botón + panel). Cierra con Escape o clic fuera. */
export function menu(botonId, panelId) {
  const b = $(botonId), p = $(panelId);
  const abrir = (si) => { p.hidden = !si; b.setAttribute('aria-expanded', String(si)); };
  b.setAttribute('aria-haspopup', 'true');
  b.setAttribute('aria-expanded', 'false');
  b.onclick = (e) => { e.stopPropagation(); abrir(p.hidden); if (!p.hidden) p.querySelector('button:not([disabled])')?.focus(); };
  p.addEventListener('click', (e) => { if (e.target.closest('button')) abrir(false); });
  document.addEventListener('click', (e) => { if (!p.hidden && !p.contains(e.target)) abrir(false); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !p.hidden) { abrir(false); b.focus(); } });
}
