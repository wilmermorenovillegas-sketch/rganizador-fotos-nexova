/**
 * ============================================================================
 * NEXOVA Suite · app2.js
 * ----------------------------------------------------------------------------
 * Shell de la plataforma (hub de herramientas) + Organizador Fotográfico.
 *
 * Solo este archivo y index.html se reescriben. motor.js, archivos.js y
 * reportes.js quedan intactos: aquí se orquestan sus funciones.
 *
 * Correcciones respecto a la versión anterior:
 *  · Scroll infinito → todos los Chart usan responsive:false con tamaño fijo.
 *  · PDF de cobertura sin imágenes → se renderizan en canvas OCULTOS en memoria
 *    y se pasan a generarPDF (nunca se reutilizan los canvas del dashboard).
 *  · IDs duplicados de progreso eliminados.
 *
 * Novedad: dashboard de DISTRIBUCIÓN "Activos por" (Familia, Subfamilia, Área,
 * Sede, Estado, Línea de producción). Si la dimensión no existe en el Excel,
 * el gráfico queda en cero. Todo exportable a Excel y PDF.
 * ============================================================================
 */
import {
  celdaATexto, cargarActivos, construirPlan, calcularMetricas, agruparPorFamilia, normalizar,
} from './motor.js';
import { leerTabla } from './shared/excel-io.js';
import { detectarColumnas, puntuarHoja } from './shared/deteccion.js';
import { soportado, pedirCarpeta, escanearFotos, ejecutarPlan, guardarEnDestino } from './archivos.js';
import { generarExcel, generarPDF, graficoEstado, graficoCobertura } from './reportes.js';
import { Depurador } from './depurador.js';
import { generarCatalogoPPT, analizarCatalogo } from './catalogo.js';
import { $, num, pct, esc, icono, toast, confirmar, descargar, zonaArchivo, zonaLista, menu } from './shared/ui.js';

const TEAL = '#0F766E';
const PALETA = ['#0F766E', '#14B8A6', '#0D9488', '#2DD4BF', '#0891B2', '#0EA5E9',
                '#6366F1', '#7C3AED', '#DB2777', '#F59E0B', '#DC2626', '#65A30D',
                '#CA8A04', '#EA580C'];

// ── Estado global de la app ─────────────────────────────────
const S = {
  inv: { registros: [], nombre: '', cols: {} },
  origen: null, destino: null,
  plan: null, simulado: false, totalCod: 0,
  blobExcel: null, blobPDF: null,
  charts: {}, tiempoInicio: null, datosDims: {}, dimActiva: 'familia',
};

// Dimensiones del dashboard. Siempre visibles; vacías si no existe la columna.
const DIMS = [
  { v: 'familia',    l: 'Familia',            pistas: ['descripcion de familia', 'descripcion familia', 'descfamilia', 'familia', 'grupo', 'rubro'] },
  { v: 'subfamilia', l: 'Subfamilia',         pistas: ['descripcion de sub familia', 'descripcion sub familia', 'descsubfamilia', 'subfamilia', 'sub familia'] },
  { v: 'area',       l: 'Área',               pistas: ['descripcion de ubicacion', 'descripcion ubicacion', 'descubicacion', 'area', 'ubicacion', 'ambiente'] },
  { v: 'sede',       l: 'Sede',               pistas: ['descripcion de centro costo', 'desccentrocosto', 'centro de costo', 'centro costo', 'sede', 'sucursal', 'local'] },
  { v: 'estado',     l: 'Estado',             pistas: ['estado de conservacion', 'estado conservacion', 'est con', 'estcon', 'estado operacion', 'estado', 'condicion', 'situacion'] },
  { v: 'linea',      l: 'Línea de producción', pistas: ['descripcion linea produccion', 'linea de produccion', 'linea produccion', 'linea', 'proceso', 'planta'] },
];

// Campos que detecta el organizador (motor de puntaje compartido: shared/deteccion.js).
const CAMPOS_ORG = [
  { k: 'codigo', hints: ['barra nueva', 'barnueva', 'barnue', 'codigo de barras', 'codigo de barra', 'codbar', 'barcode', 'codigo inventario', 'codigo del activo', 'codigo activo', 'codigo patrimonial', 'codigo interno', 'placa', 'etiqueta', 'correlativo', 'codigo', 'cod'] },
  { k: 'descripcion', hints: ['descripcion de catalogo', 'descripcion catalogo', 'desccatalogo', 'descripcion del activo', 'descripcion del bien', 'descripcion', 'denominacion', 'detalle', 'bien'] },
  ...DIMS.map((d) => ({ k: d.v, hints: d.pistas })),
  // Solo para las especificaciones del catálogo PPT
  { k: "marca", hints: ["marca", "fabricante", "brand"] },
  { k: "modelo", hints: ["modelo", "model"] },
  { k: "serie", hints: ["nro serie", "numero de serie", "nro de serie", "n serie", "serie", "serial"] },
];
const detectarColsOrg = (headers) => detectarColumnas(headers, CAMPOS_ORG);

// ════════════════════════════════════════════════════════════
// HUB — pantalla de inicio con las herramientas de la suite
// ════════════════════════════════════════════════════════════
const TOOLS = [
  { id: 'organizador', on: true,  icon: 'camara',   name: 'Organizador fotográfico',
    desc: 'Empareja las fotos del levantamiento con el inventario por código de activo, las ordena por familia y genera la conciliación, el dashboard y el catálogo en PowerPoint.' },
  { id: 'depurador',   on: true,  icon: 'datos',    name: 'Depurador y análisis de bases',
    desc: 'Normaliza la base operativa, mide la calidad de los datos, valida marcas y modelos y entrega el inventario depurado con su dashboard.' },
  { id: 'proyectos',   on: false, icon: 'proyecto', name: 'Diagnóstico de la base contable',
    desc: 'Genera el formato de diagnóstico de la base contable de activos fijos.' },
  { id: 'entregables', on: false, icon: 'entrega',  name: 'Generador de entregables',
    desc: 'Produce la documentación y los entregables del proyecto de forma automatizada.' },
  { id: 'correos',     on: false, icon: 'correo',   name: 'Asistente de correos',
    desc: 'Redacta correos profesionales para la comunicación con el cliente.' },
  { id: 'formatos',    on: false, icon: 'formato',  name: 'Generador de formatos',
    desc: 'Crea formatos y requerimientos estandarizados.' },
];

function pintarHub() {
  const activas = TOOLS.filter((t) => t.on), proximas = TOOLS.filter((t) => !t.on);
  $('toolGrid').innerHTML = activas.map((t) => `
    <button class="tool-card" data-tool="${t.id}">
      <div class="tool-icon">${icono(t.icon)}</div>
      <div class="tool-title">${esc(t.name)}</div>
      <div class="tool-desc">${esc(t.desc)}</div>
      <div class="tool-foot"><span class="chip on">Disponible</span><span class="tool-open">Abrir ${icono('flecha')}</span></div>
    </button>`).join('');
  $('toolGridSoon').innerHTML = proximas.map((t) => `
    <div class="tool-card soon">
      <div class="tool-icon">${icono(t.icon)}</div>
      <div class="tool-title">${esc(t.name)}</div>
      <div class="tool-desc">${esc(t.desc)}</div>
      <div class="tool-foot"><span class="chip soon">Próximamente</span></div>
    </div>`).join('');
  $('stActive').textContent = activas.length;
  $('stSoon').textContent = proximas.length;
  $('toolGrid').onclick = (e) => { const b = e.target.closest('[data-tool]'); if (b) abrirHerramienta(b.dataset.tool); };
}

function abrirHerramienta(id) {
  const tool = TOOLS.find((t) => t.id === id);
  if (!tool || !tool.on) return;
  document.querySelectorAll('.view').forEach((v) => v.classList.remove('visible'));
  $('crumbs').hidden = false;
  $('hTool').textContent = tool.name;
  document.title = `${tool.name} · NEXOVA Suite`;

  if (id === 'organizador') {
    $('view-organizador').classList.add('visible');
    // Compatibilidad: el organizador necesita la File System Access API.
    if (!soportado()) { $('orgApp').style.display = 'none'; $('nocompat').style.display = 'grid'; }
    else { $('orgApp').style.display = 'block'; $('nocompat').style.display = 'none'; }
  } else if (id === 'depurador') {
    $('view-depurador').classList.add('visible');
    Depurador.init();   // idempotente
  }
  window.scrollTo({ top: 0 });
}
function volverHub() {
  document.querySelectorAll('.view').forEach((v) => v.classList.remove('visible'));
  $('view-hub').classList.add('visible');
  $('crumbs').hidden = true;
  document.title = 'NEXOVA Suite · Plataforma de inventarios de activos fijos';
  window.scrollTo({ top: 0 });
}

// ════════════════════════════════════════════════════════════
// ARRANQUE
// ════════════════════════════════════════════════════════════
// requestAnimationFrame se pausa con la pestaña oculta y el PDF de reportes.js lo
// espera: si el usuario cambia de pestaña, el proceso quedaría detenido.
const rafNativo = window.requestAnimationFrame.bind(window);
window.requestAnimationFrame = (cb) => (document.hidden ? setTimeout(() => cb(performance.now()), 16) : rafNativo(cb));

pintarHub();
$('hBack').onclick = volverHub;
$('hHome').onclick = volverHub;
addEventListener('scroll', () => $('nav').classList.toggle('scrolled', scrollY > 8), { passive: true });
iniciarOrganizador();

function iniciarOrganizador() {
  zonaArchivo('zInv', (f) => cargarExcel(f), () => $('fInv').click());
  $('fInv').onchange = (e) => { cargarExcel(e.target.files[0]); e.target.value = ''; };
  zonaArchivo('zOrigen', null, () => elegirCarpeta('origen'));
  zonaArchivo('zDestino', null, () => elegirCarpeta('destino'));
  // Navegación del asistente
  $('btn1').onclick = () => irA(2);
  $('btn2Atras').onclick = () => irA(1);
  $('btn2').onclick = () => { mostrarConfirmacion(); irA(3); };
  $('btn3Atras').onclick = () => irA(2);
  $('btnVolverOrg').onclick = () => irA(3);
  $('btnSimular').onclick = () => correr(true);
  $('btnEjecutar').onclick = () => correr(false);
  $('btnEjecutarBarra').onclick = () => correr(false);
  $('selGrafico').onchange = () => pintarPanel(S.dimActiva);
  $('dimTabs').onclick = (e) => { const b = e.target.closest('.dtab'); if (b) cambiarDim(b.dataset.dim, b); };
  // Descargas
  menu('btnDescargas', 'menuDescargas');
  $('btnResumenDim').onclick = descargarResumenDim;
  $('btnAnalisisXls').onclick = descargarAnalisisExcel;
  $('btnAnalisisPdf').onclick = descargarAnalisisPDF;
  $('btnExcel').onclick = () => descargar(S.blobExcel, `Conciliacion_${S.inv.nombre}.xlsx`);
  $('btnPDF').onclick = () => descargar(S.blobPDF, `Informe_Cobertura_${S.inv.nombre}.pdf`);
  $('btnCatalogoPpt').onclick = catalogoPPT;
}

function irA(paso) {
  [1, 2, 3, 4].forEach((n) => {
    $(`p${n}`).classList.toggle('visible', n === paso);
    const ws = $(`ws${n}`); ws.classList.remove('activo', 'listo');
    ws.removeAttribute('aria-current');
    if (n === paso) { ws.classList.add('activo'); ws.setAttribute('aria-current', 'step'); } else if (n < paso) ws.classList.add('listo');
    if (n < 4) $(`wl${n}`).classList.toggle('ok', n < paso);
  });
  // La barra de "Ejecutar" solo acompaña a los resultados de una simulación.
  $('barraFija').style.display = paso === 4 && S.simulado && !S.ejecutado ? 'block' : 'none';
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── Paso 1: cargar Excel y detectar columnas ────────────────
async function cargarExcel(archivo) {
  if (!archivo) return;
  try {
    const { headers: enc, registros } = leerTabla(await archivo.arrayBuffer(), (h) => puntuarHoja(h, CAMPOS_ORG));
    const detectadas = detectarColsOrg(enc);
    const colCodigo = detectadas.codigo || '';
    if (!colCodigo) {
      toast('No se encontró la columna de código del activo (por ejemplo "Barra nueva", "Código de barras" o "CodBar").', 'error', { titulo: 'Excel sin columna de código' });
      return;
    }

    const cols = { codigo: colCodigo, descripcion: detectadas.descripcion || null };
    DIMS.forEach((d) => { cols[d.v] = detectadas[d.v] || null; });
    ["marca", "modelo", "serie"].forEach((k) => { cols[k] = detectadas[k] || null; });

    S.inv = { registros, nombre: archivo.name.replace(/\.[^.]+$/, ''), cols };
    invalidarSimulacion();

    zonaLista('zInv', archivo.name, `${num(registros.length)} registros · haga clic para cambiar de archivo`);
    $('inv-info').hidden = false;

    const tags = [`<span class="tag">Código: <b>${esc(colCodigo)}</b></span>`]
      .concat(DIMS.map((d) => (cols[d.v]
        ? `<span class="tag">${d.l}: <b>${esc(cols[d.v])}</b></span>`
        : `<span class="tag vacio">${d.l}: no está en el Excel</span>`)));
    $('inv-resumen').innerHTML = tags.join(' ');
    $('btn1').disabled = false;
  } catch (e) {
    toast(e.message, 'error', { titulo: 'No se pudo leer el archivo' });
  }
}

// ── Paso 2: carpetas ────────────────────────────────────────
async function elegirCarpeta(cual) {
  try {
    const handle = await pedirCarpeta();
    S[cual] = handle;
    invalidarSimulacion();
    let noVacia = false;
    if (cual === 'destino') { for await (const _ of handle.values()) { noVacia = true; break; } }
    if (cual === 'destino') S.destinoNoVacio = noVacia;
    if (cual === 'destino' && noVacia) {
      zonaLista('zDestino', handle.name, 'La carpeta no está vacía: los archivos con el mismo nombre se reemplazarán', true);
    } else {
      zonaLista(cual === 'origen' ? 'zOrigen' : 'zDestino', handle.name, 'Seleccionada · haga clic para cambiarla');
    }
    $('btn2').disabled = !(S.origen && S.destino);
  } catch (e) {
    if (e.name !== 'AbortError') toast(e.message, 'error', { titulo: 'No se pudo abrir la carpeta' });
  }
}

function mostrarConfirmacion() {
  $('confirmGrid').innerHTML = `
    <div class="citem"><div class="cl">Inventario</div><div class="cv">${esc(S.inv.nombre)}</div></div>
    <div class="citem"><div class="cl">Registros</div><div class="cv">${num(S.inv.registros.length)}</div></div>
    <div class="citem"><div class="cl">Origen</div><div class="cv">${esc(S.origen.name)}</div></div>
    <div class="citem"><div class="cl">Destino</div><div class="cv">${esc(S.destino.name)}</div></div>`;
}

// ── Paso 3: simular / ejecutar ──────────────────────────────
async function correr(simulacion) {
  if (!simulacion && !S.simulado) { toast('Primero corra la simulación para revisar el resultado.', 'warn'); return; }
  const modo = document.querySelector('#modos input:checked').value;
  if (!simulacion) {
    const ok = await confirmar({
      titulo: `¿${modo === 'mover' ? 'Mover' : 'Copiar'} las fotos ahora?`,
      texto: `Se van a ${modo === 'mover' ? 'MOVER' : 'COPIAR'} las fotos hacia la carpeta "${S.destino.name}".`
        + (S.destinoNoVacio ? '\n\nLa carpeta destino no está vacía: los archivos con el mismo nombre se reemplazarán.' : '')
        + (modo === 'mover' ? '\n\nAl mover, las fotos dejan de estar en la carpeta de origen.' : ''),
      aceptar: modo === 'mover' ? 'Mover fotos' : 'Copiar fotos',
      peligro: modo === 'mover' || S.destinoNoVacio,
    });
    if (!ok) return;
  }

  bloquear(true);
  $('progWrap').style.display = 'block';   // visible ANTES de setProg (fix barra)
  $('progWrap').classList.add('activo');
  $('logDet').hidden = false;
  irA(3);   // el progreso se ve en el paso 3 (también al ejecutar desde la barra inferior)
  $('barraFija').style.display = 'none';
  S.tiempoInicio = Date.now();

  const timer = setInterval(() => {
    const s = Math.floor((Date.now() - S.tiempoInicio) / 1000);
    $('progTiempo').textContent = `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`;
  }, 500);

  try {
    log('Iniciando ' + (simulacion ? 'simulación' : 'ejecución') + '…');
    if (!S.origen) throw new Error('No se seleccionó carpeta de fotos.');
    if (!S.destino) throw new Error('No se seleccionó carpeta destino.');
    log('Carpeta origen: ' + S.origen.name);
    log('Carpeta destino: ' + S.destino.name);

    setProg(0.05, 'Escaneando carpeta de fotos…');
    const fotos = await escanearFotos(S.origen, true, S.destino,
      (n) => setProg(0.05 + Math.min(n / 5000, 1) * 0.1, `Escaneando… ${num(n)} fotos encontradas`));
    log(`Fotos encontradas: ${num(fotos.length)}`);
    if (!fotos.length) throw new Error('No se encontraron fotos en la carpeta de origen.');

    setProg(0.18, 'Procesando inventario…');
    const cols = S.inv.cols;
    const cfg = {
      colCodigo: cols.codigo, colEnlaceInv: null, colDescripcion: cols.descripcion,
      usarCatalogo: false, colFamilia: cols.familia, colSubfamilia: cols.subfamilia,
      niveles: cols.subfamilia ? 2 : 1, moverHuerfanas: true,
      modo, cliente: S.inv.nombre,
    };
    const { activos, alertas } = cargarActivos(S.inv.registros, [], cfg);
    log(`Activos cargados: ${num(activos.size)}`);
    S.plan = construirPlan(fotos, activos, alertas, cfg);
    setProg(0.25, 'Plan construido…');

    if (!simulacion) {
      let prev = 0;
      await ejecutarPlan(S.plan, S.destino, modo, (h, t) => {
        setProg(0.25 + (h / t) * 0.6, `${modo === 'mover' ? 'Moviendo' : 'Copiando'} ${num(h)} de ${num(t)} fotos…`);
        if (h - prev >= 100) { log(`${num(h)} / ${num(t)} fotos procesadas`); prev = h; }
      });
      log('Archivos procesados. Generando reportes…');
    }

    setProg(0.9, 'Generando reportes de conciliación…');
    S.blobExcel = await generarExcel(S.plan, cfg, simulacion);
    S.blobPDF = await generarPDFCobertura(cfg, simulacion);   // con canvas ocultos (fix imágenes)

    if (!simulacion) {
      await guardarEnDestino(S.destino, 'Conciliacion.xlsx', S.blobExcel);
      await guardarEnDestino(S.destino, 'Informe_Cobertura.pdf', S.blobPDF);
      log('Reportes guardados en la carpeta destino.');
    }
    setProg(1, simulacion ? 'Simulación completada.' : 'Proceso completado.');
    S.simulado = S.simulado || simulacion;
    S.ejecutado = !simulacion;

    const m = calcularMetricas(S.plan);
    const seg = Math.floor((Date.now() - S.tiempoInicio) / 1000);
    const tStr = seg >= 60 ? `${Math.floor(seg / 60)} min ${seg % 60} s` : `${seg} s`;

    if (simulacion) {
      $('bannerRes').innerHTML = `<div class="banner info">
        <div class="b-ic">${icono('ojo')}</div>
        <div class="b-txt"><b>Vista previa de la simulación</b>Así quedará la organización: ${num(m.emparejadas)} fotos emparejadas y ${pct(m.cobertura)} de cobertura. Todavía no se ha copiado ni movido ningún archivo.</div>
        <button class="b-pri" id="btnEjecutarBanner">${icono('rayo')}Ejecutar ahora</button></div>`;
      $('btnEjecutarBanner').onclick = () => correr(false);
      $('bfResumen').textContent = `${num(m.emparejadas)} fotos listas · ${pct(m.cobertura)} de cobertura`;
      $('btnEjecutar').disabled = false;
    } else {
      $('bannerRes').innerHTML = `<div class="banner exito">
        <div class="b-ic">${icono('ok')}</div>
        <div class="b-txt"><b>${num(m.emparejadas)} fotos organizadas en ${tStr}</b>Las carpetas y los reportes (Conciliacion.xlsx e Informe_Cobertura.pdf) quedaron guardados en "${esc(S.destino.name)}".</div></div>`;
      sonido();
      toast(`${num(m.emparejadas)} fotos organizadas en ${tStr}.`, 'ok', { titulo: 'Proceso completado' });
    }

    pintarDashboard(cfg, m);
    irA(4);
  } catch (e) {
    log('ERROR: ' + e.message);
    toast(e.message, 'error', { titulo: simulacion ? 'La simulación no pudo completarse' : 'La ejecución no pudo completarse' });
    setProg(0, 'Error.');
  } finally {
    clearInterval(timer);
    $('progWrap').classList.remove('activo');
    bloquear(false);
  }
}

/** PDF de cobertura (reportes.js) con gráficos renderizados en canvas OCULTOS. */
async function generarPDFCobertura(cfg, simulacion) {
  const cont = document.createElement('div');
  cont.style.cssText = 'position:fixed;left:-99999px;top:0;pointer-events:none';
  const cEstado = document.createElement('canvas'); cEstado.width = 300; cEstado.height = 300;
  const cCob = document.createElement('canvas'); cCob.width = 720; cCob.height = 400;
  cont.append(cEstado, cCob);
  document.body.appendChild(cont);
  const m = calcularMetricas(S.plan);
  const familias = agruparPorFamilia(S.plan);
  const chEstado = graficoEstado(cEstado, m);
  const chCob = graficoCobertura(cCob, familias);
  try {
    return await generarPDF(S.plan, cfg, { estado: cEstado, cobertura: cCob }, simulacion);
  } finally {
    chEstado.destroy(); chCob.destroy(); cont.remove();
  }
}

// ════════════════════════════════════════════════════════════
// DASHBOARD
// ════════════════════════════════════════════════════════════
function pintarDashboard(cfg, m) {
  const cls = (v, b, med) => (v < b ? 'bad' : v < med ? 'warn' : 'ok');
  $('tituloRes').textContent = `Resultados · ${S.inv.nombre}`;
  pintarAvisos(S.plan.alertas || []);

  $('kpis').innerHTML = [
    kpi('Total activos',    num(m.totalActivos)),
    kpi('Fotos procesadas', num(m.totalFotos)),
    kpi('Cobertura',        pct(m.cobertura), cls(m.cobertura, 60, 90), `${num(m.activosConFoto)} con foto`),
    kpi('Emparejadas',      num(m.emparejadas), 'ok'),
    kpi('Sin código',       num(m.huerfanas), m.huerfanas ? 'warn' : 'ok'),
    kpi('Sin foto',         num(m.activosSinFoto), m.activosSinFoto ? 'warn' : 'ok'),
  ].join('');

  S.totalCod = new Set(S.inv.registros.map((r) => normalizar(celdaATexto(r[S.inv.cols.codigo]))).filter(Boolean)).size;

  // Distribución de todas las dimensiones (vacío si no hay columna)
  S.datosDims = {};
  DIMS.forEach((d) => { S.datosDims[d.v] = S.inv.cols[d.v] ? calcDim(S.inv.cols[d.v]) : []; });

  $('dimTabs').innerHTML = DIMS.map((d, i) => {
    const n = S.datosDims[d.v].length;
    const badge = `<span class="badge${n ? '' : ' empty'}">${n}</span>`;
    return `<button class="dtab${i === 0 ? ' activo' : ''}" role="tab" aria-selected="${i === 0}" data-dim="${d.v}">${d.l}${badge}</button>`;
  }).join('');

  // Dona global de cobertura: fija, se dibuja una sola vez.
  pintarDonaCobertura(m);
  pintarPanel('familia');
}

/** Alertas del motor (duplicados, claves ambiguas…) agrupadas por tipo, con ejemplos. */
const TEXTO_ALERTA = {
  'Codigo duplicado en el inventario': 'Códigos duplicados en el inventario',
  'Codigo repetido (se usa el primero)': 'Códigos repetidos: se usa el primer registro',
  'Clave flexible ambigua (solo match exacto)': 'Códigos parecidos entre sí: solo se emparejan por coincidencia exacta',
  'Catalogo duplicado con familias distintas': 'Catálogo duplicado con familias distintas',
  'Enlace de catalogo no encontrado': 'Enlace de catálogo no encontrado',
};
function pintarAvisos(alertas) {
  if (!alertas.length) { $('avisos').innerHTML = ''; return; }
  const grupos = new Map();
  for (const [tipo, det] of alertas) {
    if (!grupos.has(tipo)) grupos.set(tipo, []);
    grupos.get(tipo).push(det);
  }
  $('avisos').innerHTML = [...grupos.entries()].map(([tipo, dets]) => `
    <div class="aviso warn">${icono('alerta')}<div>
      <b>${esc(TEXTO_ALERTA[tipo] || tipo)}: ${num(dets.length)}</b>
      <details><summary>Ver ${dets.length > 8 ? 'los primeros 8' : 'detalle'}</summary>
        <ul>${dets.slice(0, 8).map((d) => `<li>${esc(Array.isArray(d) ? d.join(', ') : d)}</li>`).join('')}</ul>
      </details>
    </div></div>`).join('');
}

/** Agrupa los activos por el valor de una columna del inventario. */
function calcDim(col) {
  const mapa = new Map();
  const total = S.totalCod || 1;
  const vistos = new Set();
  for (const reg of S.inv.registros) {
    const cod = normalizar(celdaATexto(reg[S.inv.cols.codigo]));
    if (!cod || vistos.has(cod)) continue;
    vistos.add(cod);
    const val = celdaATexto(reg[col]) || 'SIN DATO';
    if (!mapa.has(val)) mapa.set(val, { nombre: val, activos: 0, conFoto: 0 });
    const g = mapa.get(val);
    g.activos++;
    if (S.plan.codigosConFoto.has(cod)) g.conFoto++;
  }
  return [...mapa.values()]
    .map((g) => ({
      ...g,
      sinFoto: g.activos - g.conFoto,
      cobertura: g.activos ? (g.conFoto / g.activos) * 100 : 0,
      pctTotal: (g.activos / total) * 100,
    }))
    .sort((a, b) => b.activos - a.activos);
}

function pintarPanel(dim) {
  S.dimActiva = dim;
  const datos = S.datosDims[dim] || [];
  const label = DIMS.find((d) => d.v === dim)?.l || dim;
  const tipo = $('selGrafico').value || 'barrasH';
  $('thDim').textContent = label;
  $('gPrincipalTit').textContent = `Activos por ${label.toLowerCase()}`;
  $('menuDimTxt').textContent = `Solo la dimensión activa: ${label}`;

  // Si la dimensión no está en el inventario, el gráfico se muestra EN CERO
  // (mismo formato que las demás, sin datos), no un mensaje.
  pintarGraficoDim('gPrincipal', datos, tipo);
  pintarTabla(datos, label);
  pintarCards(datos, label);
}

function cambiarDim(dim, btn) {
  document.querySelectorAll('#dimTabs .dtab').forEach((b) => { b.classList.remove('activo'); b.setAttribute('aria-selected', 'false'); });
  btn.classList.add('activo');
  btn.setAttribute('aria-selected', 'true');
  pintarPanel(dim);
}

// ── Gráficos (todos responsive:false → sin scroll infinito) ─
const fondoBlanco = {
  id: 'fondoBlanco',
  beforeDraw(c) {
    const ctx = c.ctx; ctx.save();
    ctx.globalCompositeOperation = 'destination-over';
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); ctx.restore();
  },
};
function destruir(id) { if (S.charts[id]) { S.charts[id].destroy(); delete S.charts[id]; } }

const BASE = { responsive: false, maintainAspectRatio: false, animation: false };

/** Gráfico "Activos por dimensión": conteo de activos por categoría. */
function pintarGraficoDim(id, datos, tipo) {
  destruir(id);
  const canvas = $(id); if (!canvas) return;
  const top = datos.slice(0, 14);
  const vacio = top.length === 0;                    // dimensión no presente → en cero
  const labels = vacio ? ['Sin dato'] : top.map((d) => d.nombre);

  if (tipo === 'dona') {
    S.charts[id] = new Chart(canvas.getContext('2d'), {
      type: 'doughnut', plugins: [fondoBlanco],
      data: {
        labels,
        datasets: [{
          // Ring gris "en cero" cuando no hay datos.
          data: vacio ? [1] : top.map((d) => d.activos),
          backgroundColor: vacio ? ['#E2E8F0'] : PALETA.slice(0, top.length),
          borderWidth: 2, borderColor: '#fff',
        }],
      },
      options: {
        ...BASE, cutout: '55%',
        plugins: {
          legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12 } },
          tooltip: { enabled: !vacio },
        },
      },
    });
    return;
  }

  const horizontal = tipo === 'barrasH';
  S.charts[id] = new Chart(canvas.getContext('2d'), {
    type: 'bar', plugins: [fondoBlanco],
    data: {
      labels,
      datasets: [{
        label: 'Activos', data: vacio ? [0] : top.map((d) => d.activos),
        backgroundColor: vacio ? ['#E2E8F0'] : top.map((_, i) => PALETA[i % PALETA.length]),
        borderRadius: 4, barThickness: horizontal ? 14 : undefined, maxBarThickness: 34,
      }],
    },
    options: {
      ...BASE, indexAxis: horizontal ? 'y' : 'x',
      scales: {
        x: { beginAtZero: true, suggestedMax: vacio ? 1 : undefined, ticks: { font: { size: 10 }, color: '#475569', maxRotation: horizontal ? 0 : 40, autoSkip: false }, grid: { color: horizontal ? '#F1F5F9' : 'transparent' } },
        y: { beginAtZero: true, suggestedMax: vacio && !horizontal ? 1 : undefined, ticks: { font: { size: 10 }, color: '#475569', autoSkip: false }, grid: { color: horizontal ? 'transparent' : '#F1F5F9' } },
      },
      plugins: { legend: { display: false }, tooltip: { enabled: !vacio, callbacks: { label: (c) => ` ${num(c.raw)} activos` } } },
    },
  });
}

/** Dona global de cobertura fotográfica (con foto / sin foto). */
function pintarDonaCobertura(m) {
  destruir('gDona');
  const canvas = $('gDona'); if (!canvas) return;
  S.charts.gDona = new Chart(canvas.getContext('2d'), {
    type: 'doughnut', plugins: [fondoBlanco],
    data: { labels: ['Con foto', 'Sin foto'], datasets: [{ data: [m.activosConFoto, m.activosSinFoto], backgroundColor: [TEAL, '#CBD5E1'], borderWidth: 2, borderColor: '#fff' }] },
    options: { ...BASE, cutout: '64%', plugins: { legend: { position: 'bottom', labels: { font: { size: 11 }, boxWidth: 12 } } } },
  });
}

// ── Tabla ───────────────────────────────────────────────────
function pintarTabla(datos, label) {
  const cls = (v, b, med) => (v < b ? 'bad' : v < med ? 'warn' : 'ok');
  if (!datos.length) {
    $('tbody').innerHTML = `<tr>
      <td style="color:var(--gris)">Sin dato — ${esc(label || '')} no está en este inventario</td>
      <td class="r" style="font-weight:700">0</td>
      <td class="r">0.0%</td>
      <td class="r">0</td>
      <td class="r"><span class="pill bad">0.0%</span></td>
    </tr>`;
    return;
  }
  $('tbody').innerHTML = datos.map((d) => `<tr>
    <td>${esc(d.nombre)}</td>
    <td class="r" style="font-weight:700">${num(d.activos)}</td>
    <td class="r">${pct(d.pctTotal)}</td>
    <td class="r" style="color:var(--teal)">${num(d.conFoto)}</td>
    <td class="r"><span class="pill ${cls(d.cobertura, 60, 90)}">${pct(d.cobertura)}</span></td>
  </tr>`).join('');
}

// ── Cuadros resumen ─────────────────────────────────────────
function pintarCards(datos, label) {
  if (!datos.length) {
    $('resumenCards').innerHTML = `
      <div class="res-tit">Cuadros resumen — activos por ${esc(label.toLowerCase())}</div>
      <div class="cards-grid">
        <div class="rcard">
          <div class="rcard-nombre">Sin dato</div>
          <div class="rcard-nums"><span class="rcard-total">0 activos</span><span class="rcard-pct">0.0%</span></div>
          <div class="rcard-bar-bg"><div class="rcard-bar-fill" style="width:0%"></div></div>
          <div class="rcard-sub">Esta dimensión no está en este inventario</div>
        </div>
      </div>`;
    return;
  }
  const top = datos.slice(0, 12);
  const maxA = Math.max(...top.map((d) => d.activos), 1);
  $('resumenCards').innerHTML = `
    <div class="res-tit">Cuadros resumen — activos por ${esc(label.toLowerCase())}</div>
    <div class="cards-grid">
      ${top.map((d) => `
        <div class="rcard">
          <div class="rcard-nombre">${esc(d.nombre)}</div>
          <div class="rcard-nums">
            <span class="rcard-total">${num(d.activos)} activos</span>
            <span class="rcard-pct">${pct(d.pctTotal)}</span>
          </div>
          <div class="rcard-bar-bg"><div class="rcard-bar-fill" style="width:${Math.round((d.activos / maxA) * 100)}%"></div></div>
          <div class="rcard-sub">${num(d.conFoto)} con foto · ${num(d.sinFoto)} sin foto</div>
        </div>`).join('')}
    </div>`;
}

// ════════════════════════════════════════════════════════════
// EXPORTACIONES
// ════════════════════════════════════════════════════════════

/** Excel de la dimensión activa. */
async function descargarResumenDim() {
  const dim = S.dimActiva;
  const label = DIMS.find((d) => d.v === dim)?.l || dim;
  const datos = S.datosDims[dim] || [];
  if (!datos.length) { toast(`La dimensión "${label}" no está en este inventario: no hay datos para exportar.`, 'warn'); return; }
  const wb = new ExcelJS.Workbook();
  hojaDim(wb, label, datos);
  await bajarWb(wb, `Resumen_${label}_${S.inv.nombre}.xlsx`);
}

/** Excel con TODAS las dimensiones + hoja resumen. */
async function descargarAnalisisExcel() {
  if (!S.plan) { toast('Primero corra la simulación.', 'warn'); return; }
  const m = calcularMetricas(S.plan);
  const wb = new ExcelJS.Workbook();
  wb.creator = 'NEXOVA Suite';

  const ws = wb.addWorksheet('RESUMEN');
  ws.getColumn(1).width = 32; ws.getColumn(2).width = 18;
  ws.mergeCells('A1:B1');
  ws.getCell('A1').value = 'ANÁLISIS DEL INVENTARIO';
  ws.getCell('A1').font = { bold: true, size: 15, color: { argb: 'FF1E293B' } };
  ws.getCell('A2').value = `Inventario: ${S.inv.nombre}  ·  ${new Date().toLocaleString('es-PE')}`;
  ws.getCell('A2').font = { size: 10, color: { argb: 'FF64748B' } };
  [
    ['Total de activos', m.totalActivos],
    ['Fotos procesadas', m.totalFotos],
    ['Activos con foto', m.activosConFoto],
    ['Activos sin foto', m.activosSinFoto],
    ['Cobertura fotográfica', `${m.cobertura.toFixed(1)}%`],
  ].forEach(([k, v], i) => {
    const f = ws.getRow(4 + i);
    f.getCell(1).value = k; f.getCell(2).value = v;
    f.getCell(1).font = { bold: true, color: { argb: 'FF1E293B' } };
    f.getCell(2).font = { bold: true, color: { argb: 'FF0F766E' } };
    f.getCell(2).alignment = { horizontal: 'right' };
  });

  // Todas las dimensiones (las ausentes quedan en cero, como en el dashboard).
  DIMS.forEach((d) => hojaDim(wb, d.l, S.datosDims[d.v] || []));
  await bajarWb(wb, `Analisis_${S.inv.nombre}.xlsx`);
}

function hojaDim(wb, label, datos) {
  const ws = wb.addWorksheet(label.slice(0, 31).replace(/[\\/*?:[\]]/g, ' '));
  ws.columns = [
    { header: label, key: 'n', width: 42 },
    { header: 'Activos', key: 'a', width: 12 },
    { header: '% del total', key: 'p', width: 12 },
    { header: 'Con foto', key: 'c', width: 12 },
    { header: 'Sin foto', key: 's', width: 12 },
    { header: 'Cobertura %', key: 'cb', width: 13 },
  ];
  const h = ws.getRow(1);
  h.eachCell((c) => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  });
  h.height = 18;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  if (!datos.length) {
    ws.addRow({ n: 'Sin dato (dimensión no presente en el inventario)', a: 0, p: 0, c: 0, s: 0, cb: 0 });
    return;
  }
  datos.forEach((d) => ws.addRow({
    n: d.nombre, a: d.activos, p: +d.pctTotal.toFixed(1),
    c: d.conFoto, s: d.sinFoto, cb: +d.cobertura.toFixed(1),
  }));
  ws.addRow({ n: 'TOTAL', a: datos.reduce((s, d) => s + d.activos, 0),
              c: datos.reduce((s, d) => s + d.conFoto, 0), s: datos.reduce((s, d) => s + d.sinFoto, 0) })
    .eachCell((c) => { c.font = { bold: true }; });
}

async function bajarWb(wb, nombre) {
  const buf = await wb.xlsx.writeBuffer();
  descargar(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), nombre);
}

/** PDF del análisis por dimensiones: gráficos + cuadros resumen. */
async function descargarAnalisisPDF() {
  if (!S.plan) { toast('Primero corra la simulación.', 'warn'); return; }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  const M = 16, W = 210;
  const m = calcularMetricas(S.plan);

  // Encabezado
  doc.setFillColor(30, 41, 59); doc.rect(0, 0, W, 30, 'F');
  doc.setFillColor(15, 118, 110); doc.rect(0, 30, W, 1.6, 'F');
  doc.setFont('helvetica', 'bold').setFontSize(15).setTextColor(255, 255, 255);
  doc.text('ANÁLISIS DEL INVENTARIO', M, 15);
  doc.setFont('helvetica', 'normal').setFontSize(9).setTextColor(203, 213, 225);
  doc.text(`${S.inv.nombre}  ·  ${new Date().toLocaleDateString('es-PE')}`, M, 22);

  // KPIs
  let y = 40;
  const kpis = [
    ['Total activos', num(m.totalActivos)], ['Fotos', num(m.totalFotos)],
    ['Con foto', num(m.activosConFoto)], ['Cobertura', `${m.cobertura.toFixed(1)}%`],
  ];
  const kw = (W - M * 2 - 9) / 4;
  kpis.forEach(([l, v], i) => {
    const x = M + (kw + 3) * i;
    doc.setFillColor(241, 245, 249); doc.roundedRect(x, y, kw, 18, 1.5, 1.5, 'F');
    doc.setFont('helvetica', 'normal').setFontSize(7).setTextColor(100, 116, 139);
    doc.text(l.toUpperCase(), x + 3, y + 6);
    doc.setFont('helvetica', 'bold').setFontSize(13).setTextColor(15, 118, 110);
    doc.text(String(v), x + 3, y + 14);
  });
  y += 26;

  // Contenedor oculto para renderizar los gráficos
  const cont = document.createElement('div');
  cont.style.cssText = 'position:fixed;left:-99999px;top:0;pointer-events:none';
  document.body.appendChild(cont);

  // Renderiza un Chart en canvas oculto y devuelve su PNG.
  const aPNG = async (w, h, cfg) => {
    const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
    cont.appendChild(cv);
    const ch = new Chart(cv.getContext('2d'), cfg);
    await new Promise((r) => requestAnimationFrame(r));
    const png = cv.toDataURL('image/png', 1);
    ch.destroy();
    return png;
  };

  try {
    // Gráfico global de cobertura (la dona que se ve en el dashboard).
    doc.setFillColor(15, 118, 110); doc.rect(M, y - 4, 3, 6, 'F');
    doc.setFont('helvetica', 'bold').setFontSize(12).setTextColor(30, 41, 59);
    doc.text('Cobertura fotográfica (global)', M + 6, y + 1);
    y += 8;
    const donaPng = await aPNG(460, 380, {
      type: 'doughnut', plugins: [fondoBlanco],
      data: { labels: ['Con foto', 'Sin foto'], datasets: [{ data: [m.activosConFoto, m.activosSinFoto], backgroundColor: ['#0F766E', '#CBD5E1'], borderWidth: 2, borderColor: '#fff' }] },
      options: { ...BASE, cutout: '60%', plugins: { legend: { position: 'bottom', labels: { font: { size: 12 } } } } },
    });
    doc.addImage(donaPng, 'PNG', (W - 62) / 2, y, 62, 51);
    y += 58;

    // Una sección por dimensión: TODAS se incluyen (las ausentes, en cero).
    for (const d of DIMS) {
      const datos = S.datosDims[d.v] || [];
      const vacio = !datos.length;
      const top = datos.slice(0, 14);

      if (y > 232) { doc.addPage(); y = 20; }
      doc.setFillColor(15, 118, 110); doc.rect(M, y - 4, 3, 6, 'F');
      doc.setFont('helvetica', 'bold').setFontSize(12).setTextColor(30, 41, 59);
      doc.text(`Activos por ${d.l.toLowerCase()}`, M + 6, y + 1);
      if (vacio) {
        doc.setFont('helvetica', 'normal').setFontSize(8).setTextColor(148, 163, 184);
        doc.text('No disponible en este inventario (0)', M + 6, y + 6);
      }
      y += vacio ? 11 : 8;

      // Gráfico (en cero si la dimensión no existe)
      const png = await aPNG(900, 460, {
        type: 'bar', plugins: [fondoBlanco],
        data: { labels: vacio ? ['Sin dato'] : top.map((x) => x.nombre), datasets: [{ data: vacio ? [0] : top.map((x) => x.activos), backgroundColor: vacio ? ['#E2E8F0'] : top.map((_, i) => PALETA[i % PALETA.length]), borderRadius: 4 }] },
        options: { ...BASE, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true, suggestedMax: vacio ? 1 : undefined, ticks: { color: '#475569' } }, y: { ticks: { color: '#475569', font: { size: 11 } } } } },
      });
      const imgH = 56;
      doc.addImage(png, 'PNG', M, y, W - M * 2, imgH);
      y += imgH + 4;

      // Tabla
      doc.autoTable({
        startY: y, margin: { left: M, right: M },
        head: [[d.l, 'Activos', '% total', 'Con foto', 'Cobertura']],
        body: vacio ? [['Sin dato', '0', '0.0%', '0', '0.0%']]
                    : datos.slice(0, 12).map((x) => [x.nombre, num(x.activos), pct(x.pctTotal), num(x.conFoto), pct(x.cobertura)]),
        theme: 'grid', styles: { fontSize: 8, cellPadding: 1.8 },
        headStyles: { fillColor: [30, 41, 59], textColor: 255, fontSize: 8 },
        alternateRowStyles: { fillColor: [248, 250, 251] },
        columnStyles: { 1: { halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' } },
      });
      y = doc.lastAutoTable.finalY + 8;

      // Cuadros resumen (tarjetas) — con salto de página si no caben.
      const filasCards = Math.ceil(Math.min(datos.length || 1, 9) / 3);
      const altoCards = 4 + filasCards * (19 + 4);
      if (y + altoCards > 286) { doc.addPage(); y = 20; }
      doc.setFont('helvetica', 'bold').setFontSize(8.5).setTextColor(100, 116, 139);
      doc.text('CUADROS RESUMEN', M, y);
      y = cardsPDF(doc, y + 3, datos, M, W) + 10;
    }
    descargar(doc.output('blob'), `Analisis_${S.inv.nombre}.pdf`);
  } finally {
    cont.remove();
  }
}

/** Dibuja los cuadros resumen (tarjetas) en el PDF; devuelve la nueva y. */
function cardsPDF(doc, y, datos, M, W) {
  const top = datos.length ? datos.slice(0, 9) : [{ nombre: 'Sin dato', activos: 0, pctTotal: 0 }];
  const cols = 3, gap = 4, chh = 19;
  const cw = (W - M * 2 - gap * (cols - 1)) / cols;
  const maxA = Math.max(...top.map((d) => d.activos), 1);
  top.forEach((d, i) => {
    const x = M + (i % cols) * (cw + gap);
    const yy = y + Math.floor(i / cols) * (chh + 4);
    doc.setFillColor(241, 245, 249); doc.roundedRect(x, yy, cw, chh, 1.5, 1.5, 'F');
    doc.setFont('helvetica', 'bold').setFontSize(7.5).setTextColor(30, 41, 59);
    doc.text(doc.splitTextToSize(String(d.nombre), cw - 6)[0], x + 3, yy + 6);
    doc.setFont('helvetica', 'normal').setFontSize(7).setTextColor(100, 116, 139);
    doc.text(`${num(d.activos)} activos`, x + 3, yy + 11);
    doc.setFont('helvetica', 'bold').setFontSize(9).setTextColor(15, 118, 110);
    doc.text(pct(d.pctTotal), x + cw - 3, yy + 11, { align: 'right' });
    doc.setFillColor(226, 232, 240); doc.roundedRect(x + 3, yy + 13.5, cw - 6, 2.4, 1, 1, 'F');
    doc.setFillColor(15, 118, 110);
    doc.roundedRect(x + 3, yy + 13.5, Math.max(0.4, (cw - 6) * (d.activos / maxA)), 2.4, 1, 1, 'F');
  });
  return y + Math.ceil(top.length / cols) * (chh + 4);
}

// ── Catálogo fotográfico (PPT) ──────────────────────────────
// Flujo: analizar fotos → revisar la foto principal (solo dudosos) → generar PPT.
const CAT = { analisis: null, filtro: 'revisar', limite: 60, manual: new Set(), ocupado: false };
const ETIQ_CERTEZA = { codigo: 'Código de barras', imagen: 'Probable', revisar: 'Revisar', manual: 'Elegida por usted' };
const ETIQ_TIPO = { etiqueta: 'Etiqueta', datos: 'Placa de datos', 'probable-etiqueta': 'Etiqueta?' };

function catModal(abrir) {
  $('catModal').hidden = !abrir;
  document.body.style.overflow = abrir ? 'hidden' : '';
  if (abrir) setTimeout(() => $('catCerrar').focus(), 0);
  else $('btnCatalogoPpt').focus();
}
function catProgreso(txt, h, t) {
  $('catProg').hidden = false;
  $('catProgFill').style.width = `${t ? Math.round((h / t) * 100) : 0}%`;
  $('catProgTxt').textContent = t ? `${txt} · ${num(h)} de ${num(t)} fotos` : txt;
}

async function catalogoPPT() {
  if (!S.plan) { toast('Primero corra la simulación.', 'warn'); return; }
  if (!calcularMetricas(S.plan).emparejadas) { toast('No hay fotos emparejadas para el catálogo.', 'warn'); return; }
  if (CAT.ocupado) return;
  // Enlaces del modal (una sola vez)
  if (!catModal.listo) {
    catModal.listo = true;
    const cerrar = () => { if (!CAT.ocupado) catModal(false); };
    $('catCerrar').onclick = cerrar;
    $('catCancelar').onclick = cerrar;
    $('catModal').addEventListener('keydown', (e) => { if (e.key === 'Escape') cerrar(); });
    $('catFiltros').onclick = (e) => {
      const b = e.target.closest('.cat-fil'); if (!b) return;
      CAT.filtro = b.dataset.f; CAT.limite = 60; catPintar();
    };
    $('catGrid').onclick = (e) => {
      const b = e.target.closest('.cat-foto');
      if (b) {
        const sel = CAT.analisis.seleccion.get(b.dataset.cod);
        sel.principal = +b.dataset.i; CAT.manual.add(b.dataset.cod);
        const card = b.closest('.cat-card'); card.outerHTML = catTarjeta(b.dataset.cod);
        return;
      }
      if (e.target.closest('.cat-mas')) { CAT.limite += 60; catPintar(); }
    };
    $('catGenerar').onclick = catGenerar;
  }
  CAT.analisis = null; CAT.manual = new Set(); CAT.filtro = 'revisar'; CAT.limite = 60;
  $('catSub').textContent = 'Analizando las fotografías: se lee el código de barras de cada etiqueta para identificar la toma del bien.';
  $('catResumen').hidden = true; $('catFiltros').hidden = true; $('catGrid').innerHTML = '';
  $('catGenerar').disabled = true;
  catModal(true);
  CAT.ocupado = true;
  try {
    CAT.analisis = await analizarCatalogo(S.plan, (h, t) => catProgreso('Analizando', h, t));
  } catch (e) {
    CAT.ocupado = false; catModal(false);
    toast(e.message, 'error', { titulo: 'No se pudieron analizar las fotos' }); return;
  }
  CAT.ocupado = false;
  $('catProg').hidden = true;
  const sel = [...CAT.analisis.seleccion.values()];
  const n = (c) => sel.filter((x) => x.certeza === c).length;
  const nRev = n('imagen') + n('revisar');
  $('catSub').textContent = nRev
    ? `Revise los ${num(nRev)} activos marcados. Los confirmados por código de barras ya están listos.`
    : 'Todas las fotos principales se identificaron por código de barras. Puede generar el catálogo.';
  $('catResumen').innerHTML = `
    <span class="cat-chip ok"><b>${num(n('codigo'))}</b> confirmados por código de barras</span>
    <span class="cat-chip med"><b>${num(n('imagen'))}</b> probables</span>
    <span class="cat-chip rev"><b>${num(n('revisar'))}</b> por revisar</span>`;
  $('catResumen').hidden = false; $('catFiltros').hidden = false;
  if (!nRev) CAT.filtro = 'todos';
  $('catGenerar').disabled = false;
  catPintar();
}

/** Lista de activos según el filtro (dudosos primero). */
function catLista() {
  const orden = { revisar: 0, imagen: 1, codigo: 2 };
  const regs = CAT.analisis.familias.flatMap((f) => f.activos);
  const lista = CAT.filtro === 'todos' ? regs
    : regs.filter((r) => CAT.analisis.seleccion.get(r.activo.codigo).certeza !== 'codigo');
  return lista.slice().sort((a, b) =>
    orden[CAT.analisis.seleccion.get(a.activo.codigo).certeza] - orden[CAT.analisis.seleccion.get(b.activo.codigo).certeza]);
}

function catTarjeta(cod) {
  const reg = CAT.analisis.familias.flatMap((f) => f.activos).find((r) => r.activo.codigo === cod);
  const sel = CAT.analisis.seleccion.get(cod);
  const manual = CAT.manual.has(cod);
  const cert = manual ? 'manual' : sel.certeza;
  const fotos = sel.thumbs.map((src, i) => {
    const es = i === sel.principal, tipo = ETIQ_TIPO[sel.tipos[i]];
    return `<button class="cat-foto${es ? ' sel' : ''}" data-cod="${esc(cod)}" data-i="${i}" aria-pressed="${es}" title="Usar como foto principal">
      ${src ? `<img src="${src}" alt="Foto ${i + 1} de ${esc(cod)}" loading="lazy">` : '<span class="tag">Sin vista</span>'}
      ${es ? '<span class="tag tag-p">Principal</span>' : tipo ? `<span class="tag">${tipo}</span>` : ''}
    </button>`;
  }).join('');
  return `<div class="cat-card${manual ? ' hecho' : ''}">
    <div class="cat-card-top"><span class="cat-cod">${esc(cod)}</span><span class="cat-badge ${cert}">${ETIQ_CERTEZA[cert]}</span></div>
    <div class="cat-desc" title="${esc(reg.activo.descripcion || '')}">${esc(reg.activo.descripcion || 'Sin descripción')}</div>
    <div class="cat-fotos">${fotos}</div>
  </div>`;
}

function catPintar() {
  document.querySelectorAll('#catFiltros .cat-fil').forEach((b) => b.classList.toggle('activo', b.dataset.f === CAT.filtro));
  const lista = catLista();
  const vis = lista.slice(0, CAT.limite);
  $('catGrid').innerHTML = (vis.length ? vis.map((r) => catTarjeta(r.activo.codigo)).join('')
    : '<div class="cat-vacio">No hay activos por revisar. Puede generar el catálogo.</div>')
    + (lista.length > vis.length ? `<button class="b-sec cat-mas">Mostrar ${num(Math.min(60, lista.length - vis.length))} más (${num(lista.length - vis.length)} restantes)</button>` : '');
}

async function catGenerar() {
  if (!CAT.analisis || CAT.ocupado) return;
  CAT.ocupado = true;
  $('catGenerar').disabled = true;
  try {
    const cl = S.inv.cols;
    const val = (r, k) => (cl[k] ? celdaATexto(r[cl[k]]) : '');
    const extras = new Map();
    for (const r of S.inv.registros) {
      const cod = celdaATexto(r[cl.codigo]);
      if (cod && !extras.has(cod)) extras.set(cod, { marca: val(r, 'marca'), modelo: val(r, 'modelo'), serie: val(r, 'serie'), estado: val(r, 'estado'), ubicacion: val(r, 'area') || val(r, 'sede') });
    }
    const partes = await generarCatalogoPPT(S.plan, S.inv.nombre, (h, t) => catProgreso('Armando el PowerPoint', h, t),
      { extras, seleccion: CAT.analisis.seleccion });
    $('catProgTxt').textContent = 'Catálogo generado y descargado.';
    toast(partes > 1 ? `Por su tamaño, el catálogo se dividió en ${partes} archivos.` : 'El catálogo se descargó correctamente.', 'ok', { titulo: 'Catálogo PPT listo' });
    setTimeout(() => { CAT.ocupado = false; catModal(false); }, 900);
  } catch (e) {
    CAT.ocupado = false;
    $('catGenerar').disabled = false;
    toast(e.message, 'error', { titulo: 'No se pudo generar el catálogo PPT' });
  }
}

/** Cambiar el Excel o las carpetas invalida la simulación anterior (evita ejecutar un plan viejo). */
function invalidarSimulacion() {
  S.plan = null; S.simulado = false; S.blobExcel = null; S.blobPDF = null;
  CAT.analisis = null;
  S.ejecutado = false;
  const bf = $('barraFija'); if (bf) bf.style.display = 'none';
  const be = $('btnEjecutar'); if (be) be.disabled = true;
  const br = $('bannerRes'); if (br) br.innerHTML = '';
}

// ── Utilidades ──────────────────────────────────────────────
const kpi = (l, v, cls = '', sub = '') =>
  `<div class="kpi ${cls}"><div class="kl">${l}</div><div class="kv">${v}</div>${sub ? `<div class="ks">${sub}</div>` : ''}</div>`;

function sonido() {
  try {
    const ctx = new AudioContext(), o = ctx.createOscillator(), g = ctx.createGain();
    o.connect(g); g.connect(ctx.destination);
    o.frequency.setValueAtTime(880, ctx.currentTime);
    o.frequency.setValueAtTime(1100, ctx.currentTime + 0.15);
    g.gain.setValueAtTime(0.3, ctx.currentTime);
    g.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.5);
    o.start(); o.stop(ctx.currentTime + 0.5);
  } catch { /* sin audio */ }
}
function bloquear(b) {
  $('btnSimular').disabled = b;
  $('btnEjecutar').disabled = b || !S.simulado;
  $('btnEjecutarBarra').disabled = b || !S.simulado;
  const bb = $('btnEjecutarBanner'); if (bb) bb.disabled = b;
  $('btnSimular').innerHTML = b ? `${icono('reiniciar')}Procesando…` : `${icono('ojo')}${S.simulado ? 'Volver a simular' : 'Simular primero (sin tocar archivos)'}`;
}
function setProg(f, txt) {
  const fill = $('progFill'); if (fill) fill.style.width = `${Math.min(100, f * 100).toFixed(0)}%`;
  if (txt != null && $('progTxt')) $('progTxt').textContent = txt;
}
function log(t) {
  const el = $('logbox'); if (!el) return;
  el.textContent += new Date().toLocaleTimeString('es-PE') + '  ' + t + '\n';
  el.scrollTop = el.scrollHeight;
}
