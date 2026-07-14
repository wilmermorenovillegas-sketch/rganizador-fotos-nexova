/**
 * ============================================================================
 * archivos.js — ACCESO AL DISCO (File System Access API)
 * ============================================================================
 *
 * Todo ocurre en la maquina del usuario. Ningun byte sale al servidor.
 *
 * Decision clave: NO usamos Web Worker. El trabajo pesado aqui es I/O de
 * disco, no calculo. Cada operacion es un `await` que cede el control al
 * event loop, asi que la interfaz nunca se congela — aun con 10,000 fotos.
 * Un worker solo agregaria complejidad y una clase entera de bugs.
 * ============================================================================
 */

import { esFoto } from './motor.js';

/** El navegador soporta la API? (Chrome/Edge/Opera de escritorio, y solo en HTTPS) */
export const soportado = () =>
  typeof window !== 'undefined' && 'showDirectoryPicker' in window;

/** Abre el selector de carpetas. OBLIGATORIO llamarlo desde un click del usuario. */
export async function pedirCarpeta() {
  return window.showDirectoryPicker({ mode: 'readwrite' });
}

/**
 * Recorre la carpeta de origen y devuelve las fotos.
 * Ignora la carpeta destino si esta anidada dentro del origen: sin esto,
 * al ejecutar dos veces se reprocesarian las fotos ya organizadas.
 */
export async function escanearFotos(origen, recursivo, destino, onAvance) {
  const fotos = [];

  // resolve() devuelve la ruta relativa si `destino` esta dentro de `origen`.
  let rutaDestinoDentro = null;
  if (destino) {
    try { rutaDestinoDentro = await origen.resolve(destino); } catch { /* distinto arbol */ }
  }
  const rutaDestino = rutaDestinoDentro?.length ? rutaDestinoDentro.join('/') : null;

  async function recorrer(dir, prefijo) {
    for await (const [nombre, handle] of dir.entries()) {
      const rel = prefijo ? `${prefijo}/${nombre}` : nombre;
      if (handle.kind === 'directory') {
        if (!recursivo) continue;
        if (rutaDestino && rel === rutaDestino) continue;   // saltar exactamente el destino
        await recorrer(handle, rel);
      } else if (esFoto(nombre)) {
        fotos.push({ nombre, ruta: rel, handle, carpetaOrigen: dir });
        if (fotos.length % 500 === 0) onAvance?.(fotos.length);
      }
    }
  }

  await recorrer(origen, '');
  onAvance?.(fotos.length);
  return fotos.sort((a, b) => a.ruta.localeCompare(b.ruta, 'es'));
}

/**
 * Obtiene (creando si hace falta) una carpeta anidada.
 * Cachea la PROMESA, no el handle: si ocho operaciones en paralelo piden la
 * misma carpeta a la vez, todas esperan la misma creacion en lugar de competir.
 * Con 10,000 fotos repartidas en miles de codigos, este cache es la diferencia
 * entre segundos y minutos.
 */
async function obtenerCarpeta(raiz, partes, cache) {
  let actual = raiz, ruta = '';
  for (const parte of partes) {
    ruta = ruta ? `${ruta}/${parte}` : parte;
    if (!cache.has(ruta)) {
      const padre = actual;
      cache.set(ruta, padre.getDirectoryHandle(parte, { create: true }));
    }
    actual = await cache.get(ruta);
  }
  return actual;
}

/** Ejecuta N tareas con un limite de concurrencia. Secuencial desperdicia el disco. */
async function conPool(items, limite, tarea, onAvance) {
  let siguiente = 0, hechos = 0;
  const trabajador = async () => {
    while (siguiente < items.length) {
      const i = siguiente++;
      await tarea(items[i]);
      hechos++;
      // Avisar cada 25: un postMessage/render por archivo cuesta mas que el trabajo.
      if (hechos % 25 === 0 || hechos === items.length) onAvance?.(hechos, items.length, items[i]);
    }
  };
  await Promise.all(
    Array.from({ length: Math.min(limite, items.length) }, trabajador)
  );
}

/**
 * Mueve o copia fisicamente las fotos segun el plan.
 * Nunca aborta todo el proceso por un archivo fallido: lo marca como ERROR,
 * lo registra en el reporte y sigue.
 *
 * @param {'mover'|'copiar'} modo
 */
export async function ejecutarPlan(plan, destinoRaiz, modo, onAvance) {
  const cache = new Map();
  const pendientes = plan.items.filter((i) => i.carpetas);

  await conPool(pendientes, 8, async (item) => {
    try {
      const carpeta = await obtenerCarpeta(destinoRaiz, item.carpetas, cache);

      if (modo === 'mover') {
        try {
          // move() es un RENAME a nivel de sistema de archivos: no copia bytes.
          // Miles de fotos en segundos.
          await item.handle.move(carpeta, item.nombreDestino);
        } catch {
          // Falla si origen y destino estan en unidades distintas (p.ej. una
          // unidad de red mapeada). Respaldo: copiar y borrar.
          await copiar(item, carpeta);
          try {
            await item.handle.remove();
          } catch {
            await item.carpetaOrigen.removeEntry(item.nombre);
          }
        }
      } else {
        await copiar(item, carpeta);
      }
    } catch (e) {
      item.estado = 'ERROR';
      item.detalle = String(e?.message ?? e).slice(0, 200);
    }
  }, onAvance);

  return plan;
}

async function copiar(item, carpetaDestino) {
  const archivo = await item.handle.getFile();
  const destino = await carpetaDestino.getFileHandle(item.nombreDestino, { create: true });
  const w = await destino.createWritable();
  await w.write(archivo);   // stream directo disco->disco, no carga en memoria
  await w.close();
}

/** Guarda un Blob generado (Excel/PDF) dentro de la carpeta destino. */
export async function guardarEnDestino(destinoRaiz, nombre, blob) {
  const h = await destinoRaiz.getFileHandle(nombre, { create: true });
  const w = await h.createWritable();
  await w.write(blob);
  await w.close();
}
