/**
 * ============================================================================
 * NEXOVA Suite · depurador.js — "Depurador de Bases"
 * ----------------------------------------------------------------------------
 * Limpia y normaliza bases operativas (genérico, configurable), evalúa la
 * calidad de los datos y valida marcas/modelos contra un maestro.
 *
 * · Genérico: detección de columnas por pistas + remapeo manual del usuario.
 * · Maestro de marcas/modelos: por ahora LOCAL (Excel/JSON subido). Cuando se
 *   conecte Supabase, solo cambia la fuente de `D.maestro` — la lógica queda.
 * · No se muestra el listado de "sin marca/modelo": se descarga en Excel a pedido.
 * · Cuadros resumen por dimensión con gráfico de dona (total al centro) +
 *   ranking numerado; etiquetas de cantidad siempre visibles (pantalla y PDF).
 * · PDF vectorial (jsPDF + autoTable + gráficos en canvas ocultos), no captura.
 * · Reutiliza los helpers ya probados de motor.js (no se modifica motor.js).
 * ============================================================================
 */
import { celdaATexto, normalizar, detectarFilaEncabezado, adivinarColumna } from './motor.js';
import { LOGO_AQUARIUS } from './logo-aquarius.js';

const $ = (id) => document.getElementById(id);
const num = (n) => (+n || 0).toLocaleString('es-PE');
const pct1 = (n) => (+n || 0).toFixed(1) + '%';
const esc = (t) => String(t ?? '').replace(/[<>&]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c]));

// Paletas globales seleccionables por el usuario (pocas, para no saturar).
const PALETAS = {
  teal:    { nombre: 'Teal NEXOVA', colores: ['#0F766E', '#14B8A6', '#0D9488', '#2DD4BF', '#0891B2', '#0EA5E9', '#6366F1', '#7C3AED', '#DB2777', '#F59E0B', '#DC2626', '#65A30D'] },
  azul:    { nombre: 'Azul corporativo', colores: ['#1E3A8A', '#2563EB', '#3B82F6', '#60A5FA', '#0EA5E9', '#38BDF8', '#0891B2', '#155E75', '#1D4ED8', '#93C5FD', '#0C4A6E', '#7DD3FC'] },
  naranja: { nombre: 'Naranja / Gris', colores: ['#EA7317', '#3B82F6', '#64748B', '#F59E0B', '#2563EB', '#94A3B8', '#FB923C', '#1D4ED8', '#475569', '#FDBA74', '#0891B2', '#CBD5E1'] },
  multi:   { nombre: 'Multicolor', colores: ['#0F766E', '#F59E0B', '#2563EB', '#DC2626', '#7C3AED', '#65A30D', '#DB2777', '#0891B2', '#EA580C', '#4F46E5', '#16A34A', '#0EA5E9'] },
  tierra:  { nombre: 'Verde / Tierra', colores: ['#166534', '#4D7C0F', '#CA8A04', '#B45309', '#78350F', '#15803D', '#A16207', '#3F6212', '#92400E', '#65A30D', '#854D0E', '#22C55E'] },
};
let paletaActual = 'teal';
const paleta = () => PALETAS[paletaActual].colores;
const GRIS = '#CBD5E1';

/** Campos canónicos y sus pistas de detección (genéricos, multi-base). */
const CAMPOS = [
  { k: 'codigo',       l: 'Código (barra nueva)', key: true, hints: ['barnueva', 'barra nueva', 'codigo de barras', 'codigo de barra', 'codbar', 'cod barra', 'codigo barra', 'barcode', 'barnue', 'codigo inventario', 'codigo del activo', 'codigo activo', 'codigo del bien', 'codigo bien', 'codigo patrimonial', 'codigo interno', 'placa', 'etiqueta', 'correlativo', 'id activo', 'codigo', 'cod'] },
  { k: 'bar_antigua',  l: 'Barra antigua',   hints: ['barantigua', 'barra antigua', 'codigo antiguo', 'cod antiguo', 'barra anterior', 'codigo anterior', 'codigo antiguo de barra'] },
  { k: 'bar_padre',    l: 'Barra padre',     hints: ['barpadre', 'barra padre', 'codigo padre', 'cod padre', 'barra del padre', 'codigo de barra padre'] },
  { k: 'descripcion',  l: 'Descripción',     hints: ['descripcion del activo', 'descripcion del bien', 'descripcion catalogo', 'desccatalogo', 'descripcion bien', 'denominacion', 'descripcion', 'nombre del activo', 'articulo'] },
  { k: 'sede',         l: 'Sede',            hints: ['sede', 'sucursal', 'local', 'establecimiento'] },
  { k: 'area',         l: 'Área',            hints: ['area', 'área'] },
  { k: 'cod_ubic',     l: 'Cód. Ubicación',  hints: ['codigo ubicacion', 'cod ubicacion', 'codubicacion'] },
  { k: 'ubicacion',    l: 'Ubicación',       hints: ['descripcion ubicacion', 'desc ubicacion', 'descubicacion', 'ubicacion', 'ambiente'] },
  { k: 'cod_centro',   l: 'Cód. Centro Costo', hints: ['codigo centro costo', 'codigo ccosto', 'cod centro costo', 'cod ccosto', 'codccosto', 'codcentrocosto'] },
  { k: 'centro',       l: 'Centro de costo', hints: ['descripcion centro costo', 'descripcion ccosto', 'desc centro costo', 'desccentrocosto', 'descccosto', 'centro de costo', 'centro costo', 'ccosto', 'cco'] },
  { k: 'cod_resp',     l: 'Cód. Responsable', hints: ['codigo responsable', 'cod responsable', 'codresponsable'] },
  { k: 'responsable',  l: 'Responsable',     hints: ['descripcion responsable', 'desc responsable', 'descresponsable', 'responsable', 'custodio', 'usuario asignado', 'asignado'] },
  { k: 'cod_familia',  l: 'Cód. Familia',    hints: ['codigo familia', 'cod familia', 'codfamilia'] },
  { k: 'familia',      l: 'Familia',         hints: ['descripcion familia', 'descripcion de familia', 'desc familia', 'descfamilia', 'familia', 'grupo', 'rubro', 'clase', 'categoria'] },
  { k: 'cod_subfam',   l: 'Cód. Subfamilia', hints: ['codigo subfamilia', 'codigo sub familia', 'cod subfamilia', 'codsubfamilia'] },
  { k: 'subfamilia',   l: 'Subfamilia (tipo)', hints: ['descripcion subfamilia', 'descripcion sub familia', 'desc subfamilia', 'descsubfamilia', 'subfamilia', 'sub familia', 'subgrupo', 'sub grupo', 'subclase', 'sub clase', 'subcategoria', 'sub categoria', 'tipo de activo', 'tipo activo', 'tipo de bien', 'tipo bien', 'tipo'] },
  { k: 'cod_catalogo', l: 'Cód. Catálogo',   hints: ['codigo catalogo', 'codigo de catalogo', 'cod catalogo', 'codcatalogo'] },
  { k: 'cod_linea',    l: 'Cód. Línea',      hints: ['codigo linea produccion', 'cod linea produccion', 'codigo linea', 'cod linea', 'codlineaproduccion'] },
  { k: 'linea',        l: 'Línea',           hints: ['descripcion linea produccion', 'desc linea produccion', 'descripcion linea', 'linea de produccion', 'linea produccion', 'linea', 'proceso', 'planta'] },
  { k: 'marca',        l: 'Marca',           hints: ['marca', 'fabricante', 'brand'] },
  { k: 'modelo',       l: 'Modelo',          hints: ['modelo', 'model'] },
  { k: 'serie',        l: 'Nro. Serie',      hints: ['nro serie', 'numero de serie', 'nro de serie', 'n serie', 'serie', 'serial'] },
  { k: 'medidas',      l: 'Medidas',         hints: ['medidas', 'medida', 'dimensiones', 'lxaxh'] },
  { k: 'capacidad',    l: 'Capacidad',       hints: ['capacidad', 'capac'] },
  { k: 'color',        l: 'Color',           hints: ['color'] },
  { k: 'estado',       l: 'Estado',          hints: ['estado de conservacion', 'estado conservacion', 'est conservacion', 'est de conservacion', 'estado operacion', 'estado_operacion', 'est con', 'estcon', 'est cons', 'est_con', 'estado', 'condicion', 'situacion'] },
  { k: 'detalle',      l: 'Detalle técnico', hints: ['detalle tecnico', 'detalle técnico', 'ficha tecnica', 'especificaciones', 'detalle'] },
  { k: 'observacion',  l: 'Observaciones',   hints: ['observaciones', 'observacion', 'nota', 'glosa'] },
];
// Dimensiones para los cuadros resumen (se muestran las detectadas, en orden).
const DIMS_CUADROS = ['familia', 'subfamilia', 'area', 'sede', 'centro', 'responsable', 'estado', 'ubicacion', 'linea'];
// Modelos de gráfico que el usuario puede elegir en cada cuadro.
const TIPOS_GRAF = [
  { v: 'dona',     t: '◕ Dona' },
  { v: 'semi',     t: '◗ Semicírculo' },
  { v: 'pastel',   t: '● Pastel' },
  { v: 'barrasV',  t: '▮ Barras verticales' },
  { v: 'barrasH',  t: '▬ Barras horizontales' },
  { v: 'linea',    t: '∿ Línea / Área' },
  { v: 'polar',    t: '✳ Área polar' },
  { v: 'embudo',   t: '▽ Embudo' },
  { v: 'piramide', t: '△ Pirámide' },
];
// Ícono por dimensión (identidad visual de cada cuadro).
const ICON_DIM = { familia: '🗂️', subfamilia: '🏷️', area: '🏢', sede: '📍', centro: '💰', responsable: '👤', estado: '🩺', ubicacion: '📌', linea: '🏭' };
// Dimensiones que SIEMPRE se muestran, aunque la columna no esté mapeada (se ven en 0).
const CORE_DIMS = ['familia', 'subfamilia', 'area', 'sede'];
// Variedad de tipos para que cada cuadro salga distinto por defecto (sin repetir).
const TIPOS_VARIADOS = ['barrasH', 'dona', 'pastel', 'barrasV', 'polar', 'semi', 'embudo', 'linea', 'piramide'];
/** Dimensiones a mostrar: las core siempre + las demás si su columna está mapeada. */
function dimsDashboard() { return DIMS_CUADROS.filter((k) => CORE_DIMS.includes(k) || D.cols[k]); }
/** Una dimensión está "vacía" si no tiene ninguna categoría real (solo "Sin dato" o nada). */
function dimVacio(k) { return !conteo(k).some((x) => x.valor && x.valor !== 'Sin dato'); }
// Orden fijo de columnas para el reporte/Excel depurado.
const ORDEN = ['cod_ubic', 'ubicacion', 'cod_centro', 'centro', 'cod_resp', 'responsable',
  'cod_linea', 'linea', 'bar_antigua', 'bar_padre', 'codigo', 'cod_catalogo', 'descripcion',
  'marca', 'modelo', 'serie', 'medidas', 'capacidad', 'color', 'estado', 'detalle', 'observacion'];
/** Campos detectados en el orden pedido; los no listados (sede/área/familia…) van al final. */
function camposOrdenados() {
  const det = CAMPOS.filter((c) => D.cols[c.k]);
  const enOrden = ORDEN.map((k) => det.find((c) => c.k === k)).filter(Boolean);
  const extra = det.filter((c) => !ORDEN.includes(c.k));
  return enOrden.concat(extra);
}

const D = {
  raw: [], headers: [], cols: {}, clean: [], nombre: '',
  maestro: null, alertas: [], M: {}, tab: 'resumen',
};
const CH = {};
const tipoDim = {};   // tipo de gráfico elegido por dimensión
let inicializado = false;

// ════════════════════════════════════════════════════════════
export const Depurador = { init };

function init() {
  if (inicializado) return;
  inicializado = true;
  $('depZona').onclick = () => $('depFile').click();
  $('depFile').onchange = (e) => leerBase(e.target.files[0]);
  dragDrop('depZona', (f) => leerBase(f));
  $('depProcesar').onclick = procesar;
  $('depReset').onclick = reset;
  $('depXlsClean').onclick = exportarExcelDepurado;
  $('depXlsCalidad').onclick = exportarExcelCalidad;
  $('depXlsDash').onclick = exportarExcelDashboard;
  $('depWordDash').onclick = exportarWordDashboard;
  $('depPdf').onclick = exportarPDF;
}

function dragDrop(id, cb) {
  const z = $(id);
  z.addEventListener('dragover', (e) => { e.preventDefault(); z.classList.add('drag'); });
  z.addEventListener('dragleave', () => z.classList.remove('drag'));
  z.addEventListener('drop', (e) => { e.preventDefault(); z.classList.remove('drag'); const f = e.dataTransfer.files[0]; if (f) cb(f); });
}

function reset() {
  D.raw = []; D.headers = []; D.cols = {}; D.clean = []; D.maestro = null; D.alertas = [];
  destruirCharts();
  $('depDash').style.display = 'none';
  $('depDetect').style.display = 'none';
  $('depUpload').style.display = 'block';
  const z = $('depZona');
  z.classList.remove('ok');
  z.querySelector('.ic').textContent = '🧹';
  z.querySelector('.zt').textContent = 'Haz clic o arrastra el Excel aquí';
  z.querySelector('.zs').textContent = '.xlsx · .xlsm · .xls · .csv';
}

// Compara ignorando acentos, espacios y símbolos ("Cod. Ubicación" → "CODUBICACION").
const clave = (s) => normalizar(s).replace(/[^A-Z0-9]/g, '');

/**
 * Detección por puntaje: exacto (1000) > empieza-con (650) > contenido (300),
 * con desempate por orden de la pista. Luego asignación global greedy: cada
 * columna se usa una sola vez, así "codigo" no se roba "CodigoUbicacion".
 */
function detectarColumnas(headers) {
  const H = headers.map((h) => ({ raw: h, key: clave(h) }));
  const cands = [];
  CAMPOS.forEach((campo) => {
    campo.hints.forEach((hint, hi) => {
      const hk = clave(hint);
      if (!hk) return;
      H.forEach((h) => {
        if (!h.key) return;
        let score = 0;
        if (h.key === hk) score = 1000 - hi;
        else if (h.key.startsWith(hk)) score = 650 - hi - (h.key.length - hk.length);
        else if (hk.startsWith(h.key)) score = 480 - hi - (hk.length - h.key.length);
        else if (h.key.includes(hk)) score = 300 - hi - (h.key.length - hk.length);
        if (score > 0) cands.push({ k: campo.k, header: h.raw, score });
      });
    });
  });
  cands.sort((a, b) => b.score - a.score);
  const cols = {}; const usadaH = new Set(); const usadaK = new Set();
  cands.forEach((c) => {
    if (usadaK.has(c.k) || usadaH.has(c.header)) return;
    cols[c.k] = c.header; usadaK.add(c.k); usadaH.add(c.header);
  });
  CAMPOS.forEach((c) => { if (!(c.k in cols)) cols[c.k] = ''; });
  return cols;
}

// ── 1) Leer base y detectar columnas ────────────────────────
async function leerBase(archivo) {
  if (!archivo) return;
  try {
    const buffer = await archivo.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    if (!filas.length) { alert('El archivo no tiene datos.'); return; }

    const idx = detectarFilaEncabezado(filas);
    const vistos = {};
    D.headers = (filas[idx] || []).map((c, j) => {
      let n = celdaATexto(c) || `Col${j + 1}`;
      if (vistos[n]) { vistos[n]++; n = `${n}(${vistos[n]})`; } else vistos[n] = 1;
      return n;
    });
    D.raw = [];
    for (let i = idx + 1; i < filas.length; i++) {
      const f = filas[i] || [];
      if (!f.some((c) => celdaATexto(c) !== '')) continue;
      const o = {}; D.headers.forEach((h, j) => { o[h] = f[j] ?? null; });
      D.raw.push(o);
    }
    D.nombre = archivo.name.replace(/\.[^.]+$/, '');

    // Detección automática por puntaje (exacto > prefijo > contiene), con
    // asignación global sin colisiones: cada columna va a un solo campo.
    D.cols = detectarColumnas(D.headers);

    const z = $('depZona');
    z.classList.add('ok');
    z.querySelector('.ic').textContent = '✅';
    z.querySelector('.zt').textContent = archivo.name;
    z.querySelector('.zs').textContent = `${num(D.raw.length)} registros · ${D.headers.length} columnas`;
    pintarDeteccion();
  } catch (e) {
    alert('Error leyendo el archivo: ' + e.message);
  }
}

function pintarDeteccion() {
  const opts = (sel) => ['<option value="">— (ninguna)</option>']
    .concat(D.headers.map((h) => `<option value="${esc(h)}"${h === sel ? ' selected' : ''}>${esc(h)}</option>`))
    .join('');
  $('depDetGrid').innerHTML = CAMPOS.map((c) => {
    const val = D.cols[c.k];
    return `<div class="det-row ${val ? 'hit' : ''} ${c.key ? 'key' : ''}">
      <div class="dl"><span class="dot"></span>${c.l}${c.key ? ' *' : ''}</div>
      <select data-k="${c.k}">${opts(val)}</select>
    </div>`;
  }).join('');
  $('depDetGrid').querySelectorAll('select').forEach((s) => {
    s.onchange = () => {
      D.cols[s.dataset.k] = s.value;
      s.closest('.det-row').classList.toggle('hit', !!s.value);
      revisarProcesable();
    };
  });
  $('depDetect').style.display = 'block';
  revisarProcesable();
}

function revisarProcesable() {
  const ok = !!D.cols.codigo;
  $('depProcesar').disabled = !ok;
  $('depDetAviso').innerHTML = ok ? ''
    : `<div class="aviso warn"><span>⚠️</span><div>Selecciona al menos la columna de <b>Código</b> para poder depurar.</div></div>`;
}

// ── 2) Depurar ──────────────────────────────────────────────
const gv = (row, k) => (D.cols[k] ? celdaATexto(row[D.cols[k]]) : '');

// Alias frecuentes → forma estándar (correcciones seguras de datos conocidos).
const COLOR_ALIAS = { 'PLOMO': 'GRIS', 'GRIS PLOMO': 'GRIS', 'GRISS': 'GRIS', 'BLANCA': 'BLANCO', 'NEGRA': 'NEGRO', 'CREMA CLARO': 'CREMA' };
const MARCA_ALIAS = { 'HEWLETT PACKARD': 'HP', 'HEWLETT-PACKARD': 'HP', 'H P': 'HP', 'SANSUNG': 'SAMSUNG', 'SANSUMG': 'SAMSUNG', 'SAMSUMG': 'SAMSUNG' };

const limpiar = (s) => String(s).replace(/\s+/g, ' ').trim().toUpperCase();
function normMedidas(v) {
  let s = limpiar(v);
  s = s.replace(/(\d)\s*,\s*(\d)/g, '$1.$2');   // coma decimal → punto
  s = s.replace(/\s*[x×*]\s*/gi, ' X ');         // separadores → X
  return s.replace(/\s+/g, ' ').trim();
}
/** Depura una celda: MAYÚSCULAS, sin espacios extra, medidas/color/marca estandarizados. */
function depurarCampo(k, v0) {
  if (typeof v0 !== 'string') return v0;
  let s = v0.replace(/\s+/g, ' ').trim();
  if (k === 'medidas') return normMedidas(s);
  s = s.toUpperCase().replace(/[;,]+$/, '').trim();
  if (k === 'color') s = COLOR_ALIAS[s] || s;
  else if (k === 'marca') s = MARCA_ALIAS[s] || s;
  return s;
}

function procesar() {
  D.clean = D.raw.map((row) => {
    const r = { _ch: 0 };
    CAMPOS.forEach((c) => {
      const o = gv(row, c.k);
      const v = depurarCampo(c.k, o);
      if (typeof o === 'string' && v !== o) r._ch++;
      r[c.k] = v;
    });
    return r;
  });
  validarMM();
  calcularMetricas();
  $('depUpload').style.display = 'none';
  $('depDash').style.display = 'block';
  $('depTitulo').textContent = `Base depurada · ${D.nombre}`;
  pintarTabs();
  irTab('resumen');
  window.scrollTo({ top: 0 });
}

/** Clasifica un registro frente al maestro (o solo vacíos si no hay maestro). */
function clasificarMM(r) {
  const mN = normalizar(r.marca), moN = normalizar(r.modelo);
  if (!mN && !moN) return { motivo: 'Sin marca ni modelo', tipo: 'vacio' };
  if (!mN) return { motivo: 'Sin marca', tipo: 'vacio' };
  if (!moN) return { motivo: 'Sin modelo', tipo: 'vacio' };
  if (D.maestro) {
    const m = D.maestro.get(mN);
    if (!m) return { motivo: 'Marca no registrada en el maestro', tipo: 'no_reg' };
    if (!m.modelos.has(moN)) return { motivo: 'Modelo no registrado para esa marca', tipo: 'no_reg' };
  }
  return null;
}

function validarMM() {
  D.alertas = [];
  D.clean.forEach((r, i) => {
    if (D.maestro) {   // enriquecer a la grafía canónica del maestro
      const m = D.maestro.get(normalizar(r.marca));
      if (m) { r.marca = m.nombre; const mo = m.modelos.get(normalizar(r.modelo)); if (mo) r.modelo = mo; }
    }
    const c = clasificarMM(r);
    r._mm = c ? c.tipo : 'ok';
    if (c) D.alertas.push({ n: 0, fila: i + 1, codigo: r.codigo, descripcion: r.descripcion, marca: r.marca, modelo: r.modelo, ubicacion: r.ubicacion, motivo: c.motivo });
  });
  D.alertas.forEach((a, i) => { a.n = i + 1; });
}

function calcularMetricas() {
  const total = D.clean.length;
  const presentes = CAMPOS.filter((c) => D.cols[c.k]);
  const compl = presentes.map((c) => {
    const llenos = D.clean.filter((r) => String(r[c.k] ?? '').trim() !== '').length;
    return { k: c.k, l: c.l, llenos, pct: total ? (llenos / total) * 100 : 0 };
  });
  const complGlobal = compl.length ? compl.reduce((s, x) => s + x.pct, 0) / compl.length : 0;
  const dup = (k) => {
    if (!D.cols[k]) return { n: 0, lista: [] };
    const vistos = new Map(), rep = new Map();
    D.clean.forEach((r) => { const v = normalizar(r[k]); if (!v) return; if (vistos.has(v)) rep.set(v, (rep.get(v) || 1) + 1); else vistos.set(v, 1); });
    return { n: rep.size, lista: [...rep.entries()].map(([v, c]) => ({ valor: v, veces: c })).sort((a, b) => b.veces - a.veces) };
  };
  const sinCodigo = D.clean.filter((r) => !normalizar(r.codigo)).length;
  const alertasVacio = D.alertas.filter((a) => a.motivo.startsWith('Sin')).length;
  D.M = {
    total, presentes, compl, complGlobal,
    dupCodigo: dup('codigo'), dupSerie: dup('serie'),
    sinCodigo, cambios: D.clean.reduce((s, r) => s + (r._ch || 0), 0),
    alertasVacio, alertasNoReg: D.alertas.length - alertasVacio, mmOk: D.clean.filter((r) => r._mm === 'ok').length,
  };
}

// ── 3) Dashboard ────────────────────────────────────────────
const TABS = [
  { k: 'resumen', l: 'Resúmenes' },
  { k: 'calidad', l: 'Calidad de datos' },
  { k: 'mm',      l: 'Marca / Modelo' },
];

function pintarTabs() {
  $('depTabs').innerHTML = TABS.map((t) => `<button class="dtab${t.k === D.tab ? ' activo' : ''}" data-t="${t.k}">${t.l}</button>`).join('');
  $('depTabs').querySelectorAll('.dtab').forEach((b) => { b.onclick = () => irTab(b.dataset.t); });
  const M = D.M;
  $('depKpis').innerHTML = [
    kpi('Registros', num(M.total)),
    kpi('Columnas mapeadas', num(M.presentes.length)),
    kpi('Completitud', pct1(M.complGlobal), semClass(M.complGlobal), 'promedio'),
    kpi('Códigos duplicados', num(M.dupCodigo.n), M.dupCodigo.n ? 'warn' : 'ok'),
    kpi('Celdas depuradas', num(M.cambios), 'b'),
    kpi('Sin marca/modelo reg.', num(D.alertas.length), D.alertas.length ? 'warn' : 'ok'),
  ].join('');
}

function irTab(t) {
  D.tab = t;
  $('depTabs').querySelectorAll('.dtab').forEach((b) => b.classList.toggle('activo', b.dataset.t === t));
  destruirCharts();
  if (t === 'resumen') panelCuadros();
  else if (t === 'calidad') panelCalidad();
  else if (t === 'mm') panelMM();
}

// ── Cuadros resumen por dimensión (dona + ranking numerado) ─
function cuadroData(k) {
  const arr = conteo(k);
  const total = arr.reduce((s, x) => s + x.veces, 0) || 1;
  const TOP = 8;
  const top = arr.slice(0, TOP);
  const restoSum = arr.slice(TOP).reduce((s, x) => s + x.veces, 0);
  const PAL = paleta();
  const slices = top.map((x, i) => ({ nombre: x.valor, veces: x.veces, color: PAL[i % PAL.length] }));
  if (restoSum > 0) slices.push({ nombre: `Otros (${arr.length - TOP})`, veces: restoSum, color: GRIS });
  const maxV = arr[0] ? arr[0].veces : 1;
  const list = arr.slice(0, 12).map((x, i) => ({ n: i + 1, nombre: x.valor, veces: x.veces, pct: (x.veces / total) * 100, rel: (x.veces / maxV) * 100, color: i < TOP ? PAL[i % PAL.length] : '#94A3B8' }));
  return { arr, total, slices, list, extra: Math.max(0, arr.length - 12) };
}

function panelCuadros() {
  const dims = dimsDashboard();
  if (!dims.length) {
    $('depPanel').innerHTML = `<div class="aviso warn"><span>ℹ️</span><div>No se detectaron dimensiones para resumir. Revisa el mapeo de columnas.</div></div>`;
    return;
  }
  // Tipo por defecto variado (sin repetir) para cada dimensión.
  dims.forEach((k, i) => { if (!tipoDim[k]) tipoDim[k] = TIPOS_VARIADOS[i % TIPOS_VARIADOS.length]; });
  const op = (v, sel, t) => `<option value="${v}"${sel === v ? ' selected' : ''}>${t}</option>`;
  const opciones = (tipo) => TIPOS_GRAF.map((g) => op(g.v, tipo, g.t)).join('');
  const opPal = Object.entries(PALETAS).map(([k, p]) => `<option value="${k}"${k === paletaActual ? ' selected' : ''}>${p.nombre}</option>`).join('');
  $('depPanel').innerHTML = `
    <div class="dash-tools">
      <div class="dash-tools-l"><b>${num(dims.length)}</b> dimensiones · <b>${num(D.M.total)}</b> activos analizados</div>
      <label class="dash-pal">🎨 Colores
        <select id="selPaleta">${opPal}</select>
      </label>
    </div>
    <div class="cuadros-wrap">${dims.map((k) => {
    const label = CAMPOS.find((c) => c.k === k).l;
    const tipo = tipoDim[k];
    if (dimVacio(k)) {
      return `<div class="cuadro cuadro-vacio">
        <div class="cuadro-head">
          <span class="cuadro-ico">${ICON_DIM[k] || '📊'}</span>
          <span class="cuadro-titulo">Activos por ${esc(label.toLowerCase())}</span>
          <span class="cuadro-tot">0 activos</span>
        </div>
        <div class="cuadro-body">
          <div class="cuadro-graf-box cuadro-vacio-box">
            <div class="vacio-msg"><div class="vacio-n">0</div><div class="vacio-t">Sin datos para esta dimensión</div>
              <div class="vacio-hint">La columna <b>${esc(label)}</b> no está mapeada o está vacía. Mapéala en la pantalla de detección para ver este gráfico.</div></div>
          </div>
        </div>
      </div>`;
    }
    const { total, arr, list, extra } = cuadroData(k);
    const top = arr[0];
    const topPct = top && total ? (top.veces / total) * 100 : 0;
    return `<div class="cuadro">
      <div class="cuadro-head">
        <span class="cuadro-ico">${ICON_DIM[k] || '📊'}</span>
        <span class="cuadro-titulo">Activos por ${esc(label.toLowerCase())}</span>
        <span class="cuadro-tot">${num(arr.length)} ${arr.length === 1 ? 'categoría' : 'categorías'} · ${num(total)} activos</span>
        <select class="cuadro-tipo" data-k="${k}">${opciones(tipo)}</select>
      </div>
      <div class="cuadro-body">
        <div class="cuadro-graf-box">
          <div class="cuadro-graf"><canvas id="cq_${k}" width="280" height="280"></canvas>
            <div class="cuadro-center"><b>${num(total)}</b><span>activos</span></div></div>
        </div>
        ${top ? `<div class="cuadro-insight">🏆 <b>${esc(top.valor)}</b> es la categoría con más activos — concentra el <b>${pct1(topPct)}</b> del total</div>` : ''}
        <div class="cuadro-rank-box">
          <div class="rank-tit">Ranking — de mayor a menor</div>
          <div class="rk-head">
            <span class="rk-n">N°</span><span></span>
            <span>${esc(label)}</span><span class="rk-val">Activos · %</span>
          </div>
          ${list.map((it) => `<div class="rk-row">
            <span class="rk-n">${it.n}</span><span class="rk-sw" style="background:${it.color}"></span>
            <span class="rk-name" title="${esc(it.nombre)}">${esc(it.nombre)}</span>
            <span class="rk-val">${num(it.veces)} <em>${pct1(it.pct)}</em></span>
            <span class="rk-bar"><i style="width:${it.rel.toFixed(0)}%;background:${it.color}"></i></span>
          </div>`).join('')}
          ${extra ? `<div class="rk-more">y ${num(extra)} categorías más…</div>` : ''}
        </div>
      </div>
    </div>`;
  }).join('')}</div>`;
  dims.forEach((k) => { if (!dimVacio(k)) renderCuadroChart(k, tipoDim[k]); });
  $('depPanel').querySelectorAll('.cuadro-tipo').forEach((s) => {
    s.onchange = () => { tipoDim[s.dataset.k] = s.value; renderCuadroChart(s.dataset.k, s.value); };
  });
  $('selPaleta').onchange = (e) => { paletaActual = e.target.value; panelCuadros(); };
}

function panelCalidad() {
  const M = D.M;
  const bar = (c, i) => {
    const cl = semClass(c.pct);
    return `<div class="q-row">
      <span class="q-n">${i}</span>
      <div class="qn">${esc(c.l)} <span style="color:var(--gris);font-weight:400">· ${num(c.llenos)}/${num(M.total)}</span></div>
      <div class="q-bar"><div class="q-bar-fill ${cl}" style="width:${c.pct.toFixed(0)}%"></div></div>
      <div class="q-pct ${cl}">${pct1(c.pct)}</div>
    </div>`;
  };
  const dupBlock = (titulo, d) => `<div class="cx" style="background:var(--blanco);border:1px solid var(--gris-200);border-radius:10px;padding:14px">
      <div class="gtitle">${titulo}: ${num(d.n)} valores repetidos</div>
      ${d.n ? `<div class="tabla-wrap" style="margin:0"><table><thead><tr><th style="width:34px">N°</th><th>Valor</th><th class="r">Veces</th></tr></thead><tbody>
        ${d.lista.slice(0, 10).map((x, i) => `<tr><td>${i + 1}</td><td>${esc(x.valor)}</td><td class="r">${num(x.veces)}</td></tr>`).join('')}
      </tbody></table></div>` : '<div class="mini-note">Sin duplicados.</div>'}
    </div>`;
  $('depPanel').innerHTML = `
    <div class="res-tit" style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:.04em;margin-bottom:10px">Completitud por campo</div>
    <div class="q-list" style="margin-bottom:16px">${M.compl.map((c, i) => bar(c, i + 1)).join('')}</div>
    <div class="graficos-grid" style="grid-template-columns:1fr 1fr">
      ${dupBlock('Códigos duplicados', M.dupCodigo)}
      ${dupBlock('Series duplicadas', M.dupSerie)}
    </div>
    ${M.sinCodigo ? `<div class="aviso warn" style="margin-top:12px"><span>⚠️</span><div><b>${num(M.sinCodigo)} registros sin código.</b> No podrán identificarse de forma única.</div></div>` : ''}`;
}

/** Marcas distintas detectadas en la base, con su estado frente al maestro. */
function marcasDetectadas() {
  const m = new Map();
  D.clean.forEach((r) => {
    const v = String(r.marca ?? '').trim(); if (!v) return;
    const key = normalizar(v);
    if (!m.has(key)) m.set(key, { marca: v, veces: 0 });
    m.get(key).veces++;
  });
  return [...m.values()]
    .map((x) => ({ ...x, registrada: D.maestro ? D.maestro.has(normalizar(x.marca)) : false }))
    .sort((a, b) => b.veces - a.veces);
}

function panelMM() {
  const M = D.M;
  const marcas = marcasDetectadas();
  const noReg = marcas.filter((x) => !x.registrada);
  const estado = D.maestro
    ? `<b>Maestro cargado</b>${num(D.maestro.size)} marcas registradas · valida contra el catálogo`
    : `<b>Sin maestro cargado</b>Aún no subes el catálogo, así que el 100% de las marcas figuran como no registradas. Carga un maestro para validar.`;
  const resumenMarcas = D.maestro
    ? `${num(noReg.length)} de ${num(marcas.length)} no registradas`
    : `${num(marcas.length)} marcas · 100% no registradas (sin maestro)`;
  $('depPanel').innerHTML = `
    <div class="maestro-box">
      <div class="mtxt">${estado}</div>
      <input type="file" id="depMaestro" accept=".xlsx,.xls,.csv,.json" hidden>
      <button class="b-desc teal" id="btnMaestro">⬆ Cargar maestro (Excel/JSON)</button>
      ${D.maestro ? '<button class="b-desc" id="btnMaestroClr">Quitar</button>' : ''}
    </div>
    <div class="kpis" style="grid-template-columns:repeat(4,1fr);margin-bottom:14px">
      ${kpi('Marcas distintas', num(marcas.length), 'b')}
      ${kpi('Marca/modelo OK', num(M.mmOk), 'ok')}
      ${kpi('Con campo vacío', num(M.alertasVacio), M.alertasVacio ? 'warn' : 'ok')}
      ${kpi('No registradas', num(noReg.length), noReg.length ? 'warn' : 'ok')}
    </div>
    <div class="res-tit" style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px">Marcas detectadas — ${resumenMarcas}</div>
    <div class="tabla-wrap">
      <div style="max-height:360px;overflow:auto"><table><thead><tr>
        <th style="width:36px">N°</th><th>Marca</th><th class="r">Activos</th><th>Estado</th>
      </tr></thead><tbody>
        ${marcas.slice(0, 30).map((x, i) => `<tr><td>${i + 1}</td><td>${esc(x.marca)}</td><td class="r">${num(x.veces)}</td>
          <td><span class="pill ${x.registrada ? 'ok' : 'bad'}">${x.registrada ? 'Registrada' : 'No registrada'}</span></td></tr>`).join('')
          || `<tr><td colspan="4" style="text-align:center;color:var(--gris);padding:16px">No hay marcas en la base.</td></tr>`}
      </tbody></table></div>
    </div>
    <div class="mini-note">${marcas.length > 30 ? `Se muestran 30 de ${num(marcas.length)}. ` : ''}Descarga el detalle completo abajo.</div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:12px">
      <button class="b-desc teal" id="btnMarcasXls" ${marcas.length ? '' : 'disabled'}>⬇ Marcas detectadas (Excel)</button>
      <button class="b-desc" id="btnAlertasXls" ${D.alertas.length ? '' : 'disabled'}>⬇ Activos sin marca/modelo (Excel)</button>
    </div>`;
  $('btnMaestro').onclick = () => $('depMaestro').click();
  $('depMaestro').onchange = (e) => cargarMaestro(e.target.files[0]);
  $('btnMarcasXls').onclick = exportarMarcasExcel;
  $('btnAlertasXls').onclick = exportarAlertasExcel;
  if (D.maestro) $('btnMaestroClr').onclick = () => { D.maestro = null; validarMM(); calcularMetricas(); pintarTabs(); irTab('mm'); };
}

// ── Maestro de marcas/modelos (fuente local; Supabase al final) ─
async function cargarMaestro(archivo) {
  if (!archivo) return;
  try {
    let pares = [];
    if (/\.json$/i.test(archivo.name)) {
      const arr = JSON.parse(await archivo.text());
      pares = (Array.isArray(arr) ? arr : []).map((o) => ({ marca: o.marca ?? o.Marca ?? '', modelo: o.modelo ?? o.Modelo ?? '' }));
    } else {
      const wb = XLSX.read(await archivo.arrayBuffer(), { type: 'array' });
      const filas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, blankrows: true });
      const idx = detectarFilaEncabezado(filas);
      const enc = (filas[idx] || []).map((c) => celdaATexto(c));
      const cMarca = adivinarColumna(enc, ['marca', 'fabricante', 'brand']);
      const cModelo = adivinarColumna(enc, ['modelo', 'model']);
      if (!cMarca) { alert('El maestro necesita al menos una columna "Marca".'); return; }
      const iM = enc.indexOf(cMarca), iMo = cModelo ? enc.indexOf(cModelo) : -1;
      for (let i = idx + 1; i < filas.length; i++) {
        const f = filas[i] || [];
        pares.push({ marca: celdaATexto(f[iM]), modelo: iMo >= 0 ? celdaATexto(f[iMo]) : '' });
      }
    }
    const mapa = new Map();
    for (const p of pares) {
      const mN = normalizar(p.marca); if (!mN) continue;
      if (!mapa.has(mN)) mapa.set(mN, { nombre: String(p.marca).trim(), modelos: new Map() });
      const moN = normalizar(p.modelo);
      if (moN) mapa.get(mN).modelos.set(moN, String(p.modelo).trim());
    }
    if (!mapa.size) { alert('No se pudieron leer marcas del maestro.'); return; }
    D.maestro = mapa;
    validarMM(); calcularMetricas(); pintarTabs(); irTab('mm');
  } catch (e) {
    alert('Error leyendo el maestro: ' + e.message);
  }
}

// ── Utilidades de datos ─────────────────────────────────────
function conteo(k) {
  const m = new Map();
  D.clean.forEach((r) => { const v = String(r[k] ?? '').trim() || 'Sin dato'; m.set(v, (m.get(v) || 0) + 1); });
  return [...m.entries()].map(([valor, veces]) => ({ valor, veces })).sort((a, b) => b.veces - a.veces);
}
const semClass = (v) => (v < 60 ? 'bad' : v < 90 ? 'warn' : 'ok');
const kpi = (l, v, cls = '', sub = '') => `<div class="kpi ${cls}"><div class="kl">${l}</div><div class="kv">${v}</div>${sub ? `<div class="ks">${sub}</div>` : ''}</div>`;

// ── Gráficos: dona + etiquetas de cantidad siempre visibles ─
const fondoBlanco = { id: 'fb', beforeDraw(c) { const x = c.ctx; x.save(); x.globalCompositeOperation = 'destination-over'; x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.restore(); } };

/** Dibuja el valor sobre cada elemento (arco o barra). Siempre visible. */
const etiquetas = {
  id: 'etiquetas',
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((ds, di) => {
      const meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;
      meta.data.forEach((el, i) => {
        const raw = ds.data[i];
        if (raw == null || raw === 0) return;
        const txt = (+raw).toLocaleString('es-PE');
        ctx.save();
        ctx.font = '700 11px "DM Sans", system-ui, sans-serif';
        if (el.outerRadius !== undefined) {          // arco (dona/polar)
          if (el.circumference !== undefined && el.circumference < 0.28) { ctx.restore(); return; }
          const p = el.tooltipPosition();
          ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
          ctx.lineWidth = 3; ctx.strokeStyle = '#fff'; ctx.strokeText(txt, p.x, p.y);
          ctx.fillStyle = '#1E293B'; ctx.fillText(txt, p.x, p.y);
        } else {                                     // barra
          ctx.fillStyle = '#334155';
          if (chart.options.indexAxis === 'y') { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(txt, el.x + 4, el.y); }
          else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(txt, el.x, el.y - 3); }
        }
        ctx.restore();
      });
    });
  },
};

function destruirCharts() { for (const k in CH) { CH[k]?.destroy?.(); delete CH[k]; } }

/** Construye la config de Chart.js para un tipo dado (reutilizada en pantalla y export). */
function chartConfig(tipo, slices, animate) {
  const labels = slices.map((s) => s.nombre);
  const data = slices.map((s) => s.veces);
  const colors = slices.map((s) => s.color);
  const tot = data.reduce((s, v) => s + (+v || 0), 0) || 1;
  const common = { responsive: false, maintainAspectRatio: false,
    animation: animate ? { duration: 750, easing: 'easeOutQuart' } : false,
    plugins: { legend: { display: false },
      tooltip: { padding: 9, titleFont: { size: 12 }, bodyFont: { size: 12 }, boxPadding: 4,
        callbacks: { label: (c) => { const v = +(c.raw ?? c.parsed?.y ?? c.parsed) || 0; return `  ${num(v)} activos · ${pct1((v / tot) * 100)}`; } } } } };
  const ejes = { x: { grid: { color: '#EEF2F6' }, ticks: { font: { size: 9 }, color: '#64748B' } },
                 y: { grid: { color: '#EEF2F6' }, ticks: { font: { size: 9 }, color: '#64748B' }, beginAtZero: true } };
  if (tipo === 'polar') {
    return { type: 'polarArea', plugins: [fondoBlanco, etiquetas],
      data: { labels, datasets: [{ data, backgroundColor: colors.map((c) => c + 'D0'), borderColor: '#fff', borderWidth: 1, hoverOffset: 6 }] },
      options: { ...common, scales: { r: { ticks: { display: false }, grid: { color: '#EEF2F6' }, angleLines: { color: '#EEF2F6' } } } } };
  }
  if (tipo === 'semi') {
    return { type: 'doughnut', plugins: [fondoBlanco, etiquetas],
      data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 2, borderColor: '#fff', hoverOffset: 10 }] },
      options: { ...common, rotation: -90, circumference: 180, cutout: '55%' } };
  }
  if (tipo === 'pastel') {
    return { type: 'pie', plugins: [fondoBlanco, etiquetas],
      data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 2, borderColor: '#fff', hoverOffset: 10 }] },
      options: { ...common } };
  }
  if (tipo === 'barrasV' || tipo === 'barrasH') {
    return { type: 'bar', plugins: [fondoBlanco, etiquetas],
      data: { labels, datasets: [{ data, backgroundColor: colors, borderRadius: 6, borderSkipped: false, maxBarThickness: 46 }] },
      options: { ...common, indexAxis: tipo === 'barrasH' ? 'y' : 'x', scales: ejes } };
  }
  if (tipo === 'linea') {
    return { type: 'line', plugins: [fondoBlanco, etiquetas],
      data: { labels, datasets: [{ data, borderColor: colors[0], backgroundColor: colors[0] + '33', fill: true, tension: 0.35, pointBackgroundColor: colors[0], pointRadius: 3, pointHoverRadius: 6 }] },
      options: { ...common, scales: ejes } };
  }
  // dona (por defecto)
  return { type: 'doughnut', plugins: [fondoBlanco, etiquetas],
    data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 2, borderColor: '#fff', hoverOffset: 10 }] },
    options: { ...common, cutout: '58%' } };
}

/** Dibujo propio de embudo / pirámide (Chart.js no los trae). Segmentos apilados. */
function dibujarFunnel(cv, slices, invertido) {
  const ctx = cv.getContext('2d');
  ctx.clearRect(0, 0, cv.width, cv.height);
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height);
  const W = cv.width, H = cv.height, padX = 14, padY = 12;
  const seg = slices.filter((s) => s.veces > 0);
  const n = seg.length; if (!n) return;
  const maxW = W - padX * 2, rowH = (H - padY * 2) / n, cx = W / 2, maxV = seg[0].veces || 1;
  const wOf = (i) => {
    // Embudo: ancho ∝ valor (grande arriba). Pirámide: ancho crece hacia abajo.
    const frac = invertido ? (n - i) / n : Math.max(0.14, seg[i].veces / maxV);
    return Math.max(18, maxW * frac);
  };
  for (let i = 0; i < n; i++) {
    const y = padY + rowH * i;
    const wTop = wOf(i);
    const wBot = i < n - 1 ? wOf(i + 1) : (invertido ? wTop : wTop * 0.55);
    ctx.beginPath();
    ctx.moveTo(cx - wTop / 2, y);
    ctx.lineTo(cx + wTop / 2, y);
    ctx.lineTo(cx + wBot / 2, y + rowH - 3);
    ctx.lineTo(cx - wBot / 2, y + rowH - 3);
    ctx.closePath();
    ctx.fillStyle = seg[i].color; ctx.fill();
    ctx.strokeStyle = '#fff'; ctx.lineWidth = 2; ctx.stroke();
    ctx.font = '700 11px "DM Sans", system-ui, sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.lineWidth = 3; ctx.strokeStyle = '#fff'; ctx.strokeText(num(seg[i].veces), cx, y + rowH / 2);
    ctx.fillStyle = '#1E293B'; ctx.fillText(num(seg[i].veces), cx, y + rowH / 2);
  }
}

/** Renderiza el gráfico de un cuadro en el tipo elegido. */
function renderCuadroChart(k, tipo) {
  const id = `cq_${k}`;
  if (CH[id]) { CH[id].destroy(); delete CH[id]; }
  const cv = $(id); if (!cv) return;
  const { slices } = cuadroData(k);
  const center = cv.parentElement.querySelector('.cuadro-center');
  if (center) center.style.display = (tipo === 'dona') ? 'flex' : 'none';
  if (tipo === 'embudo' || tipo === 'piramide') {
    dibujarFunnel(cv, slices, tipo === 'piramide');
    return;
  }
  CH[id] = new Chart(cv.getContext('2d'), chartConfig(tipo, slices, true));
}

// ════════════════════════════════════════════════════════════
// EXPORTACIONES
// ════════════════════════════════════════════════════════════
function bajar(blob, nombre) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = nombre; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function styleHead(row) {
  row.eachCell((c) => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; });
  row.height = 18;
}

async function exportarExcelDepurado() {
  if (!D.clean.length) { alert('Primero procesa una base.'); return; }
  const cols = camposOrdenados();
  const lastCol = cols.length + 1;
  const HROW = 6;                                   // fila de cabeceras de la tabla
  const wb = new ExcelJS.Workbook(); wb.creator = 'NEXOVA Suite';
  const ws = wb.addWorksheet('INVENTARIO FISICO', { views: [{ state: 'frozen', ySplit: HROW }] });

  ws.getColumn(1).width = 7;
  cols.forEach((c, i) => { ws.getColumn(i + 2).width = Math.max(12, Math.min(42, c.l.length + 8)); });

  // ── Cabecera de marca (filas 1-5) ──
  try {
    const ext = LOGO_AQUARIUS.slice(11, LOGO_AQUARIUS.indexOf(';'));   // "jpeg" / "png"
    const imgId = wb.addImage({ base64: LOGO_AQUARIUS.split(',')[1], extension: ext === 'jpg' ? 'jpeg' : ext });
    ws.addImage(imgId, { tl: { col: 0.15, row: 0.2 }, ext: { width: 150, height: 82 } });
  } catch { /* sin logo */ }
  const banner = (row, text, font) => {
    ws.mergeCells(row, 3, row, Math.max(4, lastCol));
    const c = ws.getCell(row, 3); c.value = text; c.font = font; c.alignment = { vertical: 'middle' };
  };
  banner(2, 'INVENTARIO FÍSICO DEL ACTIVO FIJO', { bold: true, size: 16, color: { argb: 'FF0F2B47' } });
  banner(3, (D.nombre || '').toUpperCase(), { size: 11, color: { argb: 'FF5A6B7E' } });
  banner(4, `${num(D.M.total)} ACTIVOS INVENTARIADOS`, { bold: true, size: 12, color: { argb: 'FF1E5A8E' } });
  banner(5, 'ATF-PR-01-FO-03 / VER.02   ·   DEPURADO CON NEXOVA SUITE', { size: 9, color: { argb: 'FF8899AA' } });

  // ── Cabecera de la tabla (MAYÚSCULAS) ──
  const hr = ws.getRow(HROW);
  hr.values = ['N°'].concat(cols.map((c) => c.l.toUpperCase()));
  styleHead(hr);
  ws.autoFilter = { from: { row: HROW, column: 1 }, to: { row: HROW, column: lastCol } };

  // ── Datos (ya vienen en MAYÚSCULAS de la depuración) ──
  D.clean.forEach((r, i) => ws.addRow([i + 1].concat(cols.map((c) => r[c.k] ?? ''))));

  bajar(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Inventario_Depurado_${D.nombre}.xlsx`);
}

async function exportarMarcasExcel() {
  const marcas = marcasDetectadas();
  if (!marcas.length) { alert('No hay marcas para exportar.'); return; }
  const wb = new ExcelJS.Workbook(); wb.creator = 'NEXOVA Suite';
  const ws = wb.addWorksheet('MARCAS DETECTADAS');
  ws.columns = [{ header: 'N°', key: 'n', width: 6 }, { header: 'MARCA', key: 'm', width: 32 }, { header: 'ACTIVOS', key: 'a', width: 12 }, { header: 'ESTADO', key: 'e', width: 18 }];
  styleHead(ws.getRow(1));
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  marcas.forEach((x, i) => ws.addRow({ n: i + 1, m: x.marca, a: x.veces, e: x.registrada ? 'REGISTRADA' : 'NO REGISTRADA' }));
  bajar(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Marcas_Detectadas_${D.nombre}.xlsx`);
}

async function exportarExcelCalidad() {
  if (!D.clean.length) { alert('Primero procesa una base.'); return; }
  const M = D.M;
  const wb = new ExcelJS.Workbook(); wb.creator = 'NEXOVA Suite';

  const ws = wb.addWorksheet('RESUMEN');
  ws.getColumn(1).width = 34; ws.getColumn(2).width = 16;
  ws.getCell('A1').value = 'CALIDAD DE DATOS'; ws.getCell('A1').font = { bold: true, size: 15, color: { argb: 'FF1E293B' } };
  ws.getCell('A2').value = `${D.nombre}  ·  ${new Date().toLocaleString('es-PE')}`; ws.getCell('A2').font = { size: 10, color: { argb: 'FF64748B' } };
  [['Registros', M.total], ['Columnas mapeadas', M.presentes.length], ['Completitud promedio', `${M.complGlobal.toFixed(1)}%`],
   ['Códigos duplicados', M.dupCodigo.n], ['Series duplicadas', M.dupSerie.n], ['Registros sin código', M.sinCodigo],
   ['Celdas depuradas', M.cambios], ['Sin marca/modelo registrado', D.alertas.length]].forEach(([k, v], i) => {
    const f = ws.getRow(4 + i); f.getCell(1).value = k; f.getCell(2).value = v;
    f.getCell(1).font = { bold: true }; f.getCell(2).font = { bold: true, color: { argb: 'FF0F766E' } }; f.getCell(2).alignment = { horizontal: 'right' };
  });

  const wc = wb.addWorksheet('COMPLETITUD');
  wc.columns = [{ header: 'N°', key: 'n', width: 6 }, { header: 'Campo', key: 'l', width: 26 }, { header: 'Con dato', key: 'd', width: 12 }, { header: 'Total', key: 't', width: 12 }, { header: '% completitud', key: 'p', width: 14 }];
  styleHead(wc.getRow(1));
  M.compl.forEach((c, i) => wc.addRow({ n: i + 1, l: c.l, d: c.llenos, t: M.total, p: +c.pct.toFixed(1) }));

  const wd = wb.addWorksheet('DUPLICADOS');
  wd.columns = [{ header: 'N°', key: 'n', width: 6 }, { header: 'Tipo', key: 't', width: 12 }, { header: 'Valor', key: 'v', width: 32 }, { header: 'Veces', key: 'x', width: 10 }];
  styleHead(wd.getRow(1));
  let dn = 0;
  M.dupCodigo.lista.forEach((x) => wd.addRow({ n: ++dn, t: 'Código', v: x.valor, x: x.veces }));
  M.dupSerie.lista.forEach((x) => wd.addRow({ n: ++dn, t: 'Serie', v: x.valor, x: x.veces }));

  bajar(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Calidad_${D.nombre}.xlsx`);
}

async function exportarAlertasExcel() {
  if (!D.alertas.length) { alert('No hay activos sin marca/modelo para exportar.'); return; }
  const wb = new ExcelJS.Workbook(); wb.creator = 'NEXOVA Suite';
  const wa = wb.addWorksheet('SIN MARCA-MODELO');
  wa.columns = [{ header: 'N°', key: 'n', width: 6 }, { header: 'Fila base', key: 'f', width: 9 }, { header: 'Código', key: 'c', width: 20 },
    { header: 'Descripción', key: 'd', width: 40 }, { header: 'Marca', key: 'm', width: 18 }, { header: 'Modelo', key: 'mo', width: 18 },
    { header: 'Ubicación', key: 'u', width: 24 }, { header: 'Motivo', key: 'mt', width: 34 }];
  styleHead(wa.getRow(1));
  wa.views = [{ state: 'frozen', ySplit: 1 }];
  D.alertas.forEach((a) => wa.addRow({ n: a.n, f: a.fila, c: a.codigo, d: a.descripcion, m: a.marca, mo: a.modelo, u: a.ubicacion, mt: a.motivo }));
  bajar(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Sin_Marca_Modelo_${D.nombre}.xlsx`);
}

/** Renderiza un cuadro a PNG respetando el tipo elegido (para incrustar en Word). */
function cuadroAPng(k, cont, opts = {}) {
  const { legend = true, w = 460, h = 340 } = opts;
  const { slices } = cuadroData(k);
  const tipo = tipoDim[k] || 'dona';
  const cv = document.createElement('canvas'); cv.width = w; cv.height = h; cont.appendChild(cv);
  if (tipo === 'embudo' || tipo === 'piramide') {
    dibujarFunnel(cv, slices, tipo === 'piramide');
    return cv.toDataURL('image/png', 1);
  }
  const cfg = chartConfig(tipo, slices, false);   // sin animación → Chart.js dibuja síncrono
  if (legend) cfg.options = { ...cfg.options, plugins: { ...cfg.options.plugins, legend: { display: true, position: 'right', labels: { font: { size: 10 }, boxWidth: 12 } } } };
  const ch = new Chart(cv.getContext('2d'), cfg);
  const img = cv.toDataURL('image/png', 1); ch.destroy();
  return img;
}

/** Excel del dashboard: una hoja con el ranking de cada dimensión (Familia, Subfamilia, Área, Sede…). */
async function exportarExcelDashboard() {
  if (!D.clean.length) { alert('Primero procesa una base.'); return; }
  const dims = dimsDashboard().filter((k) => !dimVacio(k));
  if (!dims.length) { alert('No hay dimensiones con datos para el dashboard.'); return; }
  const cont = document.createElement('div');
  cont.style.cssText = 'position:fixed;left:-99999px;top:0;pointer-events:none';
  document.body.appendChild(cont);
  const wb = new ExcelJS.Workbook(); wb.creator = 'NEXOVA Suite';
  const ws = wb.addWorksheet('DASHBOARD', { views: [{ state: 'frozen', ySplit: 2 }] });
  ws.getColumn(1).width = 7; ws.getColumn(2).width = 44; ws.getColumn(3).width = 14; ws.getColumn(4).width = 12;
  ws.mergeCells(1, 1, 1, 4);
  const t = ws.getCell(1, 1);
  t.value = `DASHBOARD DE ACTIVOS · ${(D.nombre || '').toUpperCase()}`;
  t.font = { bold: true, size: 14, color: { argb: 'FF1E293B' } };
  ws.mergeCells(2, 1, 2, 4);
  const st = ws.getCell(2, 1);
  st.value = `${num(D.M.total)} activos · ${dims.length} dimensiones · ${new Date().toLocaleString('es-PE')}`;
  st.font = { size: 10, color: { argb: 'FF64748B' } };
  let row = 4;
  try {
    dims.forEach((k) => {
      const label = CAMPOS.find((c) => c.k === k).l;
      const { arr, total } = cuadroData(k);
      const startRow = row;
      ws.mergeCells(row, 1, row, 4);
      const h = ws.getCell(row, 1);
      h.value = `ACTIVOS POR ${label.toUpperCase()}   —   ${num(arr.length)} categorías · ${num(total)} activos`;
      h.font = { bold: true, size: 11, color: { argb: 'FF0F766E' } };
      h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDFA' } };
      row++;
      const hr = ws.getRow(row);
      hr.values = ['N°', label.toUpperCase(), 'ACTIVOS', '%']; styleHead(hr); row++;
      arr.forEach((x, i) => {
        const r = ws.getRow(row);
        r.getCell(1).value = i + 1; r.getCell(2).value = x.valor; r.getCell(3).value = x.veces;
        r.getCell(4).value = total ? +((x.veces / total) * 100).toFixed(1) : 0;
        r.getCell(3).alignment = { horizontal: 'right' }; r.getCell(4).alignment = { horizontal: 'right' };
        row++;
      });
      const tr = ws.getRow(row);
      tr.getCell(2).value = 'TOTAL'; tr.getCell(3).value = total; tr.getCell(4).value = 100;
      tr.font = { bold: true }; tr.getCell(3).alignment = { horizontal: 'right' }; tr.getCell(4).alignment = { horizontal: 'right' };
      row++;
      // Gráfico (según el tipo elegido) a la derecha de la tabla
      try {
        const img = cuadroAPng(k, cont, { legend: true, w: 480, h: 320 });
        const imgId = wb.addImage({ base64: img.split(',')[1], extension: 'png' });
        ws.addImage(imgId, { tl: { col: 5, row: startRow - 1 }, ext: { width: 380, height: 250 } });
      } catch { /* sin gráfico */ }
      const blockRows = Math.max(row - startRow, 14);   // deja sitio para la imagen
      row = startRow + blockRows + 1;
    });
  } finally { cont.remove(); }
  bajar(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Dashboard_${D.nombre}.xlsx`);
}

/** Word (.doc) del dashboard: gráfico + tabla por dimensión. Abre en MS Word. */
async function exportarWordDashboard() {
  if (!D.clean.length) { alert('Primero procesa una base.'); return; }
  const dims = dimsDashboard().filter((k) => !dimVacio(k));
  if (!dims.length) { alert('No hay dimensiones con datos para el dashboard.'); return; }
  const cont = document.createElement('div');
  cont.style.cssText = 'position:fixed;left:-99999px;top:0;pointer-events:none';
  document.body.appendChild(cont);
  let secciones = '';
  try {
    for (const k of dims) {
      const label = CAMPOS.find((c) => c.k === k).l;
      const { arr, total } = cuadroData(k);
      const img = await cuadroAPng(k, cont);
      const filas = arr.map((x, i) => `<tr><td style="text-align:center">${i + 1}</td><td>${esc(x.valor)}</td><td style="text-align:right">${num(x.veces)}</td><td style="text-align:right">${pct1(total ? (x.veces / total) * 100 : 0)}</td></tr>`).join('');
      secciones += `<h2 style="color:#0F766E;font-family:Calibri,sans-serif">Activos por ${esc(label.toLowerCase())}</h2>
        <p style="color:#64748B;font-family:Calibri,sans-serif;margin:0 0 6px">${num(arr.length)} categorías · ${num(total)} activos</p>
        <img src="${img}" width="440"/>
        <table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;width:100%;font-family:Calibri,sans-serif;font-size:11px;margin:8px 0 18px">
          <thead><tr style="background:#1E293B;color:#fff"><th style="width:36px">N°</th><th>${esc(label)}</th><th style="width:70px">Activos</th><th style="width:60px">%</th></tr></thead>
          <tbody>${filas}<tr style="font-weight:bold;background:#F1F5F9"><td></td><td>TOTAL</td><td style="text-align:right">${num(total)}</td><td style="text-align:right">100%</td></tr></tbody>
        </table>`;
    }
  } finally { cont.remove(); }
  const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>Dashboard de activos</title></head>
    <body style="font-family:Calibri,sans-serif">
      <h1 style="color:#1E293B;margin-bottom:2px">Dashboard de activos</h1>
      <p style="color:#64748B;margin-top:0">${esc(D.nombre)} · ${new Date().toLocaleDateString('es-PE')} · ${num(D.M.total)} activos · depurado con NEXOVA Suite</p>
      <hr/>
      ${secciones}
    </body></html>`;
  bajar(new Blob(['﻿' + html], { type: 'application/msword' }), `Dashboard_${D.nombre}.doc`);
}

async function exportarPDF() {
  if (!D.clean.length) { alert('Primero procesa una base.'); return; }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  const M = 16, W = 210; const m = D.M;

  doc.setFillColor(30, 41, 59); doc.rect(0, 0, W, 30, 'F');
  doc.setFillColor(15, 118, 110); doc.rect(0, 30, W, 1.6, 'F');
  doc.setFont('helvetica', 'bold').setFontSize(15).setTextColor(255, 255, 255);
  doc.text('DEPURACIÓN Y CALIDAD DE DATOS', M, 15);
  doc.setFont('helvetica', 'normal').setFontSize(9).setTextColor(203, 213, 225);
  doc.text(`${D.nombre}  ·  ${new Date().toLocaleDateString('es-PE')}`, M, 22);

  let y = 40;
  const kpis = [['Registros', num(m.total)], ['Completitud', `${m.complGlobal.toFixed(1)}%`], ['Cód. duplicados', num(m.dupCodigo.n)], ['Sin marca/modelo', num(D.alertas.length)]];
  const kw = (W - M * 2 - 9) / 4;
  kpis.forEach(([l, v], i) => {
    const x = M + (kw + 3) * i;
    doc.setFillColor(241, 245, 249); doc.roundedRect(x, y, kw, 18, 1.5, 1.5, 'F');
    doc.setFont('helvetica', 'normal').setFontSize(7).setTextColor(100, 116, 139); doc.text(l.toUpperCase(), x + 3, y + 6);
    doc.setFont('helvetica', 'bold').setFontSize(12).setTextColor(15, 118, 110); doc.text(String(v), x + 3, y + 14);
  });
  y += 26;

  doc.setFillColor(15, 118, 110); doc.rect(M, y - 4, 3, 6, 'F');
  doc.setFont('helvetica', 'bold').setFontSize(12).setTextColor(30, 41, 59); doc.text('Completitud por campo', M + 6, y + 1); y += 6;
  doc.autoTable({
    startY: y, margin: { left: M, right: M },
    head: [['N°', 'Campo', 'Con dato', 'Total', '% completitud']],
    body: m.compl.map((c, i) => [i + 1, c.l, num(c.llenos), num(m.total), `${c.pct.toFixed(1)}%`]),
    theme: 'grid', styles: { fontSize: 8, cellPadding: 1.8 }, headStyles: { fillColor: [30, 41, 59], textColor: 255, fontSize: 8 },
    alternateRowStyles: { fillColor: [248, 250, 251] }, columnStyles: { 0: { cellWidth: 10, halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' } },
  });
  y = doc.lastAutoTable.finalY + 10;

  const cont = document.createElement('div'); cont.style.cssText = 'position:fixed;left:-99999px;top:0;pointer-events:none'; document.body.appendChild(cont);
  try {
    const dims = dimsDashboard().filter((k) => !dimVacio(k));
    for (const k of dims) {
      if (y > 200) { doc.addPage(); y = 20; }
      const label = CAMPOS.find((c) => c.k === k).l;
      const { list, total, extra } = cuadroData(k);
      doc.setFillColor(15, 118, 110); doc.rect(M, y - 4, 3, 6, 'F');
      doc.setFont('helvetica', 'bold').setFontSize(12).setTextColor(30, 41, 59); doc.text(`Activos por ${label.toLowerCase()}`, M + 6, y + 1); y += 8;

      // Gráfico en el tipo elegido por el usuario (canvas oculto, render síncrono)
      const img = cuadroAPng(k, cont, { legend: false, w: 460, h: 460 });
      doc.addImage(img, 'PNG', M, y, 58, 58);

      // Ranking numerado al lado
      doc.autoTable({
        startY: y, margin: { left: M + 64, right: M },
        head: [['N°', label, 'Activos', '% ']],
        body: list.map((it) => [it.n, it.nombre, num(it.veces), pct1(it.pct)]).concat(extra ? [['', `… y ${num(extra)} más`, '', '']] : []),
        theme: 'grid', styles: { fontSize: 7.5, cellPadding: 1.4 }, headStyles: { fillColor: [30, 41, 59], textColor: 255, fontSize: 7.5 },
        alternateRowStyles: { fillColor: [248, 250, 251] }, columnStyles: { 0: { cellWidth: 8, halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' } },
      });
      y = Math.max(y + 62, doc.lastAutoTable.finalY + 10);
    }
    bajar(doc.output('blob'), `Informe_Calidad_${D.nombre}.pdf`);
  } finally {
    cont.remove();
  }
}
