/**
 * NEXOVA Suite · shared/maestro-nube.js
 * Maestro de marcas y modelos alojado en Supabase (proyecto "nexova-suite").
 *
 * Privacidad: solo se DESCARGA el catálogo (consulta GET de lectura pública).
 * Ningún dato de la base del cliente se envía: la validación ocurre en el navegador.
 * La clave publicable es pública por diseño; RLS impide escribir sin autenticación.
 */
import { normalizar } from '../motor.js';

export const SUPABASE_URL = 'https://nqgytziglkbffcfnvkwy.supabase.co';
const CLAVE_PUBLICABLE = 'sb_publishable_Su6fOYl4OaHTZTA-rErlBw_sXq-SeCo';
const PAGINA = 1000;   // tope de filas por respuesta de PostgREST

/** Convierte pares {marca, modelo} al mapa que usa el depurador. */
export function construirMaestro(pares, base = new Map()) {
  for (const p of pares) {
    const mN = normalizar(p.marca); if (!mN) continue;
    if (!base.has(mN)) base.set(mN, { nombre: String(p.marca).trim(), modelos: new Map() });
    const moN = normalizar(p.modelo);
    if (moN) base.get(mN).modelos.set(moN, String(p.modelo).trim());
  }
  return base;
}

/**
 * Descarga el maestro completo (paginado). Lanza error si no hay conexión
 * o si tarda más de `ms`; el depurador sigue funcionando sin él.
 * @returns {Promise<Array<{marca:string, modelo:string|null}>>}
 */
export async function descargarMaestro(ms = 8000) {
  const ctl = new AbortController();
  const t = setTimeout(() => ctl.abort(), ms);
  const filas = [];
  try {
    for (let desde = 0; ; desde += PAGINA) {
      const r = await fetch(`${SUPABASE_URL}/rest/v1/maestro?select=marca,modelo&order=marca.asc,modelo.asc&limit=${PAGINA}&offset=${desde}`, {
        headers: { apikey: CLAVE_PUBLICABLE },
        signal: ctl.signal,
        referrerPolicy: 'no-referrer',
      });
      if (!r.ok) throw new Error(`El servidor respondió ${r.status}`);
      const lote = await r.json();
      filas.push(...lote);
      if (lote.length < PAGINA) break;
    }
    return filas;
  } finally {
    clearTimeout(t);
  }
}
