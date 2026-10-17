/**
 * ============================================================================
 * shared/excel-io.js — Lectura robusta de Excel/CSV para todas las herramientas
 * ----------------------------------------------------------------------------
 * · Recorta el rango de cada hoja al área con datos reales. Muchos Excel de
 *   cliente traen formato aplicado hasta la fila 1.048.576: sin este recorte
 *   SheetJS genera cientos de miles de filas vacías y la pestaña se cuelga.
 * · Elige la hoja correcta cuando el libro trae portada u hojas auxiliares
 *   (la de mejor puntaje según la herramienta; empate → la primera).
 * · Detecta la fila de encabezado y desduplica nombres de columna.
 *
 * Usa SheetJS como global (window.XLSX / globalThis.XLSX).
 * ============================================================================
 */
import { celdaATexto, detectarFilaEncabezado } from '../motor.js';

const X = () => globalThis.XLSX;

/** Ajusta !ref al área con datos reales (siempre desde A1, para no correr índices). */
export function recortarRango(ws) {
  let r2 = -1, c2 = -1;
  for (const k in ws) {
    if (k.charCodeAt(0) === 33) continue;            // claves internas "!ref", "!merges"…
    const cel = ws[k];
    if (!cel || cel.v === undefined || cel.v === null || cel.v === '') continue;
    const { r, c } = X().utils.decode_cell(k);
    if (r > r2) r2 = r;
    if (c > c2) c2 = c;
  }
  ws['!ref'] = r2 < 0 ? 'A1:A1' : X().utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r2, c: c2 } });
  return ws;
}

/** Nombres de columna limpios y únicos ("Col3" si está vacío, "Marca(2)" si se repite). */
export function encabezados(fila = []) {
  const vistos = {};
  return fila.map((c, j) => {
    let n = celdaATexto(c) || `Col${j + 1}`;
    if (vistos[n]) { vistos[n]++; n = `${n}(${vistos[n]})`; } else vistos[n] = 1;
    return n;
  });
}

/**
 * Lee un libro y devuelve la tabla de la hoja más adecuada.
 * @param {ArrayBuffer} buffer
 * @param {(headers:string[]) => number} [puntuar]  puntaje de una hoja según sus encabezados
 * @returns {{hoja:string, hojas:string[], headers:string[], registros:Object[]}}
 */
export function leerTabla(buffer, puntuar) {
  const wb = X().read(buffer, { type: 'array', cellDates: true });
  let mejor = null;
  for (const hoja of wb.SheetNames) {
    const ws = recortarRango(wb.Sheets[hoja]);
    const filas = X().utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    if (!filas.some((f) => f && f.some((c) => celdaATexto(c) !== ''))) continue;
    const idx = detectarFilaEncabezado(filas);
    const headers = encabezados(filas[idx] || []);
    const score = puntuar ? puntuar(headers) : 0;
    if (!mejor || score > mejor.score) mejor = { hoja, filas, idx, headers, score };
    if (!puntuar) break;                              // sin criterio: la primera hoja con datos
  }
  if (!mejor) return { hoja: wb.SheetNames[0] || '', hojas: wb.SheetNames, headers: [], registros: [] };

  const { filas, idx, headers } = mejor;
  const registros = [];
  for (let i = idx + 1; i < filas.length; i++) {
    const f = filas[i] || [];
    if (!f.some((c) => celdaATexto(c) !== '')) continue;
    const o = {};
    headers.forEach((h, j) => { o[h] = f[j] ?? null; });
    registros.push(o);
  }
  return { hoja: mejor.hoja, hojas: wb.SheetNames, headers, registros };
}
