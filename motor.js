/**
 * ============================================================================
 * NEXOVA · AssetControl — Organizador Fotografico de Inventarios
 * motor.js — LOGICA DE NEGOCIO PURA
 * ============================================================================
 *
 * Este modulo NO toca el DOM ni el sistema de archivos. Solo transforma datos.
 * Gracias a eso se puede testear en Node sin navegador (ver test-motor.mjs).
 *
 * Portado desde nucleo.py, ya depurado. La logica de emparejamiento respeta
 * exactamente la regla de DOS PASADAS: la coincidencia exacta manda siempre
 * sobre la flexible, incluso cuando exige recortar el sufijo de foto multiple.
 * ============================================================================
 */

export const EXTENSIONES_FOTO = new Set([
  'jpg', 'jpeg', 'png', 'webp', 'heic', 'heif',
  'bmp', 'tif', 'tiff', 'gif', 'avif',
]);

export const CARPETA_HUERFANAS = '_SIN_CLASIFICAR';
export const SIN_FAMILIA = 'SIN FAMILIA';
export const SIN_SUBFAMILIA = 'SIN SUBFAMILIA';

// Sufijos de foto multiple: COD-1 / COD_2 / COD.3 / COD (4)
// Se EXIGE un separador antes del numero. Sin esa exigencia, "EC001" se
// recortaria a "EC00" y el emparejamiento se rompe por completo.
const SUFIJO_SEPARADOR = /^(.+?)[\s._-]+\(?\d{1,3}\)?$/;
const SUFIJO_PARENTESIS = /^(.+?)\s*\(\d{1,3}\)$/;

const CARACTERES_INVALIDOS = /[<>:"/\\|?*\u0000-\u001f]/g;

// ---------------------------------------------------------------------------
// NORMALIZACION
// ---------------------------------------------------------------------------

/**
 * Convierte cualquier celda de Excel a texto limpio.
 * CRITICO: Excel entrega los codigos numericos como Number (1001, o incluso
 * 1001.0). Sin esta conversion, el 90% de los inventarios no empareja nunca.
 */
export function celdaATexto(valor) {
  if (valor === null || valor === undefined) return '';
  if (typeof valor === 'boolean') return valor ? 'SI' : 'NO';
  if (typeof valor === 'number') {
    return Number.isInteger(valor) ? String(valor) : String(valor);
  }
  if (valor instanceof Date) return valor.toLocaleDateString('es-PE');
  return String(valor).trim();
}

/** Nivel 1: mayusculas, sin tildes, espacios colapsados. CONSERVA los guiones. */
export function normalizar(valor) {
  return celdaATexto(valor)
    .normalize('NFKD')                 // separa la letra de su tilde
    .replace(/[\u0300-\u036f]/g, '')   // elimina la tilde
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Nivel 2: solo alfanumericos + se eliminan los ceros a la izquierda de cada
 * bloque numerico. Hace que EC-001, "ec 001" y EC1 emparejen entre si.
 */
export function claveFlexible(valor) {
  return normalizar(valor)
    .replace(/[^A-Z0-9]/g, '')
    .replace(/0*(\d+)/g, '$1');
}

/** Convierte un texto de familia/subfamilia en un nombre de carpeta valido. */
export function sanearCarpeta(nombre, porDefecto = 'SIN DATO') {
  const limpio = celdaATexto(nombre)
    .replace(CARACTERES_INVALIDOS, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^[.\s]+|[.\s]+$/g, '')   // Windows no admite punto/espacio final
    .slice(0, 80)
    .trim();
  return limpio || porDefecto;
}

/** Separa "EC-001-2.jpg" en { base: "EC-001-2", ext: "jpg" }. */
export function partirNombre(nombreArchivo) {
  const i = nombreArchivo.lastIndexOf('.');
  if (i <= 0) return { base: nombreArchivo, ext: '' };
  return {
    base: nombreArchivo.slice(0, i),
    ext: nombreArchivo.slice(i + 1).toLowerCase(),
  };
}

export const esFoto = (nombre) => EXTENSIONES_FOTO.has(partirNombre(nombre).ext);

// ---------------------------------------------------------------------------
// DETECCION DE COLUMNAS
// ---------------------------------------------------------------------------

/**
 * Detecta la fila de encabezado. Los Excel de cliente casi nunca empiezan en
 * A1: traen logo, titulo y fechas. Gana la fila con mas celdas de TEXTO
 * (los encabezados son palabras; los datos, cifras).
 */
export function detectarFilaEncabezado(filas, ventana = 25) {
  let mejorIdx = 0, mejorPuntaje = -1;
  for (let i = 0; i < Math.min(filas.length, ventana); i++) {
    const llenas = (filas[i] || []).filter((c) => celdaATexto(c) !== '');
    if (llenas.length < 2) continue;
    const textuales = llenas.filter((c) => typeof c !== 'number').length;
    const puntaje = llenas.length + textuales * 2;
    if (puntaje > mejorPuntaje) { mejorIdx = i; mejorPuntaje = puntaje; }
  }
  return mejorIdx;
}

/**
 * Preselecciona la columna probable. LA PRIORIDAD LA DEFINE EL ORDEN DE LAS
 * PISTAS, no el tipo de coincidencia: para cada pista se prueba igualdad y
 * luego "contiene", antes de pasar a la pista siguiente.
 *
 * Por que importa: con encabezados "Cod. Catalogo" y "Descripcion", si primero
 * barrieramos todas las pistas por igualdad, "descripcion" ganaria por
 * coincidencia exacta y el enlace al catalogo se mapearia mal. Ese bug ocurrio
 * de verdad durante las pruebas.
 */
export function adivinarColumna(encabezados, pistas) {
  const norm = encabezados.map((h) => [h, normalizar(h)]);
  for (const pista of pistas) {
    const p = normalizar(pista);
    for (const [orig, n] of norm) if (n === p) return orig;
    for (const [orig, n] of norm) if (n.includes(p)) return orig;
  }
  return '';
}

export const PISTAS = {
  codigo: ['codigo activo', 'codigo de activo', 'cod activo', 'codigo interno',
           'codigo', 'cod', 'placa', 'etiqueta', 'id activo', 'correlativo'],
  enlaceInv: ['codigo catalogo', 'cod catalogo', 'catalogo', 'descripcion',
              'denominacion', 'bien', 'detalle', 'articulo'],
  enlaceCat: ['codigo catalogo', 'cod catalogo', 'catalogo', 'codigo', 'cod',
              'descripcion', 'denominacion', 'bien'],
  familia: ['familia', 'grupo', 'rubro', 'clase', 'categoria'],
  subfamilia: ['subfamilia', 'sub familia', 'sub-familia', 'subgrupo',
               'sub clase', 'subcategoria', 'tipo'],
  descripcion: ['descripcion', 'denominacion', 'detalle', 'bien', 'nombre',
                'articulo'],
};

// ---------------------------------------------------------------------------
// MOTOR DE EMPAREJAMIENTO  (el corazon del sistema)
// ---------------------------------------------------------------------------

export class IndiceCodigos {
  /** @param {Array<{codigo:string}>} activos */
  constructor(activos) {
    this.exacto = new Map();      // clave normalizada -> activo
    this.flexible = new Map();    // clave flexible    -> activo
    this.ambiguos = new Set();    // claves flexibles que colisionan
    this.duplicados = [];         // codigos repetidos en el inventario

    for (const activo of activos) {
      const clave = normalizar(activo.codigo);
      if (!clave) continue;
      if (this.exacto.has(clave)) { this.duplicados.push(activo.codigo); continue; }
      this.exacto.set(clave, activo);

      const flex = claveFlexible(activo.codigo);
      if (!flex) continue;
      const previo = this.flexible.get(flex);
      if (previo && previo.codigo !== activo.codigo) {
        this.ambiguos.add(flex);   // colision -> jamas adivinamos
      } else {
        this.flexible.set(flex, activo);
      }
    }
  }

  /** Recorta un sufijo de foto multiple. Devuelve '' si no hay nada que recortar. */
  static recortarSufijo(texto) {
    for (const patron of [SUFIJO_PARENTESIS, SUFIJO_SEPARADOR]) {
      const m = patron.exec(texto);
      if (m) {
        const base = m[1].replace(/^[\s._-]+|[\s._-]+$/g, '');
        if (base) return base;
      }
    }
    return '';
  }

  /**
   * DOS PASADAS. El orden es lo que hace correcto a este motor.
   *
   *   Pasada 1 (ESTRICTA): coincidencia exacta sobre el nombre completo y
   *                        luego sobre el nombre sin el sufijo de foto.
   *   Pasada 2 (FLEXIBLE): recien ahora, la clave sin guiones ni ceros.
   *
   * Con los activos MOB-10 y MOB-102 en el inventario, el archivo MOB-10-2.jpg
   * es la foto 2 de MOB-10. Si la clave flexible actuara primero, colapsaria a
   * "MOB102" y asignaria la foto al activo equivocado, en silencio.
   * La exactitud manda siempre.
   */
  buscar(nombreSinExtension) {
    const candidatos = [normalizar(nombreSinExtension)];
    while (candidatos.length < 3) {                    // maximo 2 recortes
      const recorte = IndiceCodigos.recortarSufijo(candidatos[candidatos.length - 1]);
      if (!recorte || recorte === candidatos[candidatos.length - 1]) break;
      candidatos.push(recorte);
    }

    for (let i = 0; i < candidatos.length; i++) {      // --- Pasada 1: exacta
      const a = this.exacto.get(candidatos[i]);
      if (a) return { activo: a, metodo: i === 0 ? 'exacto' : 'exacto+sufijo' };
    }

    for (let i = 0; i < candidatos.length; i++) {      // --- Pasada 2: flexible
      const flex = claveFlexible(candidatos[i]);
      if (!flex || this.ambiguos.has(flex)) continue;
      const a = this.flexible.get(flex);
      if (a) return { activo: a, metodo: i === 0 ? 'normalizado' : 'normalizado+sufijo' };
    }

    return { activo: null, metodo: 'SIN MATCH' };
  }
}

// ---------------------------------------------------------------------------
// CRUCE INVENTARIO x CATALOGO
// ---------------------------------------------------------------------------

/**
 * @param {Array<Object>} filasInv   filas del inventario (objetos por encabezado)
 * @param {Array<Object>} filasCat   filas del catalogo (o [] si no se usa)
 * @param {Object} cfg               columnas mapeadas por el usuario
 * @returns {{activos: Map<string,Object>, alertas: Array}}
 */
export function cargarActivos(filasInv, filasCat, cfg) {
  const alertas = [];
  const mapaCatalogo = new Map();

  if (cfg.usarCatalogo) {
    for (const fila of filasCat) {
      const enlace = normalizar(fila[cfg.colEnlaceCat]);
      if (!enlace) continue;
      const familia = celdaATexto(fila[cfg.colFamilia]);
      const sub = cfg.colSubfamilia ? celdaATexto(fila[cfg.colSubfamilia]) : '';
      const previo = mapaCatalogo.get(enlace);
      if (previo && (previo.familia !== familia || previo.subfamilia !== sub)) {
        alertas.push(['Catalogo duplicado con familias distintas', enlace]);
      }
      mapaCatalogo.set(enlace, { familia, subfamilia: sub });
      const flex = claveFlexible(enlace);                     // respaldo flexible
      if (flex && !mapaCatalogo.has('~' + flex)) {
        mapaCatalogo.set('~' + flex, { familia, subfamilia: sub });
      }
    }
  }

  const activos = new Map();
  for (const fila of filasInv) {
    const codigo = celdaATexto(fila[cfg.colCodigo]);
    if (!codigo) continue;
    const clave = normalizar(codigo);
    if (activos.has(clave)) {
      alertas.push(['Codigo duplicado en el inventario', codigo]);
      continue;
    }

    const enlace = cfg.colEnlaceInv ? celdaATexto(fila[cfg.colEnlaceInv]) : '';
    const descripcion = cfg.colDescripcion ? celdaATexto(fila[cfg.colDescripcion]) : '';

    let familia = '', subfamilia = '';
    if (cfg.usarCatalogo) {
      let hit = mapaCatalogo.get(normalizar(enlace));
      if (!hit) hit = mapaCatalogo.get('~' + claveFlexible(enlace));
      if (hit) {
        familia = hit.familia; subfamilia = hit.subfamilia;
      } else {
        alertas.push(['Enlace de catalogo no encontrado', `${codigo} -> "${enlace}"`]);
      }
    } else {
      familia = cfg.colFamilia ? celdaATexto(fila[cfg.colFamilia]) : '';
      subfamilia = cfg.colSubfamilia ? celdaATexto(fila[cfg.colSubfamilia]) : '';
    }

    activos.set(clave, { codigo, enlace, descripcion, familia, subfamilia });
  }

  return { activos, alertas };
}

// ---------------------------------------------------------------------------
// PLAN (SIMULACION)
// ---------------------------------------------------------------------------

/**
 * Decide, para cada foto, a que activo pertenece y a que carpeta va.
 * NO toca el disco: solo produce el plan. La ejecucion vive en fs.js.
 *
 * @param {Array<{nombre:string, ruta:string, handle:any}>} fotos
 * @param {Map<string,Object>} activos
 * @param {Array} alertas
 * @param {Object} cfg  { niveles, moverHuerfanas }
 */
export function construirPlan(fotos, activos, alertas, cfg) {
  const indice = new IndiceCodigos(activos.values());
  const alertasTotales = [...alertas];

  for (const dup of indice.duplicados) {
    alertasTotales.push(['Codigo repetido (se usa el primero)', dup]);
  }
  for (const amb of [...indice.ambiguos].sort()) {
    alertasTotales.push(['Clave flexible ambigua (solo match exacto)', amb]);
  }

  const items = [];
  const codigosConFoto = new Set();
  const destinosUsados = new Set();

  for (const foto of fotos) {
    const { base } = partirNombre(foto.nombre);
    const { activo, metodo } = indice.buscar(base);

    let carpetas, estado;
    if (activo) {
      codigosConFoto.add(normalizar(activo.codigo));
      carpetas = [sanearCarpeta(activo.familia, SIN_FAMILIA)];
      if (cfg.niveles >= 2) carpetas.push(sanearCarpeta(activo.subfamilia, SIN_SUBFAMILIA));
      if (cfg.niveles >= 3) carpetas.push(sanearCarpeta(activo.codigo, 'SIN CODIGO'));
      estado = 'OK';
    } else {
      estado = 'SIN CLASIFICAR';
      if (!cfg.moverHuerfanas) {
        items.push({ ...foto, activo: null, metodo, carpetas: null,
                     nombreDestino: null, estado, detalle: 'No se movio (opcion desactivada)' });
        continue;
      }
      carpetas = [CARPETA_HUERFANAS];
    }

    // Evitar pisar archivos homonimos (frecuente al escanear subcarpetas)
    let nombreDestino = foto.nombre;
    let clave = (carpetas.join('/') + '/' + nombreDestino).toLowerCase();
    let n = 2;
    while (destinosUsados.has(clave)) {
      const { base: b, ext } = partirNombre(foto.nombre);
      nombreDestino = `${b}__${n}.${ext}`;
      clave = (carpetas.join('/') + '/' + nombreDestino).toLowerCase();
      n++;
    }
    destinosUsados.add(clave);

    items.push({
      ...foto,
      activo,
      metodo,
      carpetas,
      nombreDestino,
      estado,
      detalle: '',
    });
  }

  return { items, activos, codigosConFoto, alertas: alertasTotales };
}

// ---------------------------------------------------------------------------
// METRICAS
// ---------------------------------------------------------------------------

export function calcularMetricas(plan) {
  const emparejadas = plan.items.filter((i) => i.activo).length;
  const huerfanas = plan.items.filter((i) => !i.activo);
  const activosSinFoto = [...plan.activos.values()]
    .filter((a) => !plan.codigosConFoto.has(normalizar(a.codigo)));
  const sinFamilia = [...plan.activos.values()]
    .filter((a) => !String(a.familia).trim()).length;
  const errores = plan.items.filter((i) => i.estado === 'ERROR').length;
  const total = plan.activos.size;

  return {
    totalActivos: total,
    totalFotos: plan.items.length,
    emparejadas,
    huerfanas: huerfanas.length,
    listaHuerfanas: huerfanas,
    activosConFoto: plan.codigosConFoto.size,
    activosSinFoto: activosSinFoto.length,
    listaActivosSinFoto: activosSinFoto,
    cobertura: total ? (plan.codigosConFoto.size / total) * 100 : 0,
    sinFamilia,
    pctSinFamilia: total ? (sinFamilia / total) * 100 : 0,
    errores,
  };
}

/** Agrega por familia y subfamilia. Alimenta los graficos y el PDF. */
export function agruparPorFamilia(plan) {
  const mapa = new Map();

  for (const activo of plan.activos.values()) {
    const fam = String(activo.familia).trim() || SIN_FAMILIA;
    if (!mapa.has(fam)) {
      mapa.set(fam, { familia: fam, activos: 0, conFoto: 0, fotos: 0, subfamilias: new Map() });
    }
    const f = mapa.get(fam);
    f.activos++;
    const tieneFoto = plan.codigosConFoto.has(normalizar(activo.codigo));
    if (tieneFoto) f.conFoto++;

    const sub = String(activo.subfamilia).trim() || SIN_SUBFAMILIA;
    if (!f.subfamilias.has(sub)) {
      f.subfamilias.set(sub, { subfamilia: sub, activos: 0, conFoto: 0, fotos: 0 });
    }
    const s = f.subfamilias.get(sub);
    s.activos++;
    if (tieneFoto) s.conFoto++;
  }

  for (const item of plan.items) {
    if (!item.activo) continue;
    const fam = String(item.activo.familia).trim() || SIN_FAMILIA;
    const sub = String(item.activo.subfamilia).trim() || SIN_SUBFAMILIA;
    const f = mapa.get(fam);
    if (!f) continue;
    f.fotos++;
    const s = f.subfamilias.get(sub);
    if (s) s.fotos++;
  }

  return [...mapa.values()]
    .map((f) => ({
      ...f,
      cobertura: f.activos ? (f.conFoto / f.activos) * 100 : 0,
      sinFoto: f.activos - f.conFoto,
      subfamilias: [...f.subfamilias.values()]
        .map((s) => ({ ...s, cobertura: s.activos ? (s.conFoto / s.activos) * 100 : 0,
                       sinFoto: s.activos - s.conFoto }))
        .sort((a, b) => b.activos - a.activos),
    }))
    .sort((a, b) => b.activos - a.activos);
}
