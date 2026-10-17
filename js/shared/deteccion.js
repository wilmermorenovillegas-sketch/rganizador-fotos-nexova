/**
 * ============================================================================
 * shared/deteccion.js — Detección de columnas por puntaje (una sola fuente)
 * ----------------------------------------------------------------------------
 * Para cada par (campo, encabezado) se calcula un puntaje:
 *   exacto 1000 · encabezado empieza con la pista 650 · la pista empieza con el
 *   encabezado 480 · el encabezado contiene la pista 300   (− orden de la pista)
 * y se asigna de forma global y greedy: cada columna va a un solo campo, así
 * "codigo" no se apropia de "CODIGO DE UBICACIÓN" si existe "BARRA NUEVA".
 *
 * Encabezados muy cortos ("N°", "C", "E") solo cuentan por coincidencia exacta:
 * antes "N°" caía en Observaciones (pista "nota") y "E" en Estado.
 * ============================================================================
 */
import { normalizar } from '../motor.js';

/** Compara ignorando acentos, espacios y símbolos ("Cód. Ubicación" → "CODUBICACION"). */
export const clave = (s) => normalizar(s).replace(/[^A-Z0-9]/g, '');

const MIN_PARCIAL = 4;   // largo mínimo de encabezado para coincidencias parciales

/** Puntaje de un encabezado frente a una pista (0 = no coincide). */
export function puntaje(hKey, hk, ordenPista) {
  if (!hKey || !hk) return 0;
  if (hKey === hk) return 1000 - ordenPista;
  if (hKey.length < MIN_PARCIAL) return 0;
  if (hKey.startsWith(hk)) return 650 - ordenPista - (hKey.length - hk.length);
  if (hk.startsWith(hKey)) return 480 - ordenPista - (hk.length - hKey.length);
  if (hk.length >= 3 && hKey.includes(hk)) return 300 - ordenPista - (hKey.length - hk.length);
  return 0;
}

/**
 * @param {string[]} headers                 encabezados del Excel
 * @param {{k:string, hints:string[]}[]} campos  campos canónicos con sus pistas
 * @returns {Object<string,string>}           campo → encabezado ('' si no se detectó)
 */
export function detectarColumnas(headers, campos) {
  const H = headers.map((h) => ({ raw: h, key: clave(h) }));
  const cands = [];
  for (const campo of campos) {
    campo.hints.forEach((hint, hi) => {
      const hk = clave(hint);
      for (const h of H) {
        const s = puntaje(h.key, hk, hi);
        if (s > 0) cands.push({ k: campo.k, header: h.raw, score: s });
      }
    });
  }
  cands.sort((a, b) => b.score - a.score);
  const cols = {}, usadaH = new Set(), usadaK = new Set();
  for (const c of cands) {
    if (usadaK.has(c.k) || usadaH.has(c.header)) continue;
    cols[c.k] = c.header; usadaK.add(c.k); usadaH.add(c.header);
  }
  for (const c of campos) if (!(c.k in cols)) cols[c.k] = '';
  return cols;
}

/** Cuántos campos reconoce un juego de encabezados (sirve para elegir la hoja). */
export function puntuarHoja(headers, campos, clavePrincipal = 'codigo') {
  const cols = detectarColumnas(headers, campos);
  const n = Object.values(cols).filter(Boolean).length;
  return n + (cols[clavePrincipal] ? 10 : 0);
}
