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
  celdaATexto, detectarFilaEncabezado, adivinarColumna,
  cargarActivos, construirPlan, calcularMetricas, agruparPorFamilia, normalizar,
} from './motor.js';
import { soportado, pedirCarpeta, escanearFotos, ejecutarPlan, guardarEnDestino } from './archivos.js';
import { generarExcel, generarPDF, graficoEstado, graficoCobertura } from './reportes.js';

// ── Helpers cortos ──────────────────────────────────────────
const $ = (id) => document.getElementById(id);
const num = (n) => (+n || 0).toLocaleString('es-PE');
const pct = (n) => (+n || 0).toFixed(1) + '%';
const esc = (t) => String(t).replace(/[<>&]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c]));

const TEAL = '#0F766E', AMBAR = '#D97706', ROJO = '#B91C1C';
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
  { v: 'familia',    l: 'Familia',            pistas: ['DESCRIPCION DE FAMILIA', 'descripcion de familia', 'familia', 'grupo', 'rubro'] },
  { v: 'subfamilia', l: 'Subfamilia',         pistas: ['DESCRIPCION DE SUB FAMILIA', 'descripcion de sub familia', 'subfamilia', 'sub familia'] },
  { v: 'area',       l: 'Área',               pistas: ['DescUbicacion', 'Desc Ubicacion', 'descripcion ubicacion', 'area', 'ubicacion'] },
  { v: 'sede',       l: 'Sede',               pistas: ['DescCentroCosto', 'Desc Centro Costo', 'centro costo descripcion', 'sede', 'sucursal', 'centro de costo'] },
  { v: 'estado',     l: 'Estado',             pistas: ['Estado', 'estado', 'condicion', 'situacion'] },
  { v: 'linea',      l: 'Línea de producción', pistas: ['linea', 'linea produccion', 'linea de produccion', 'proceso', 'planta'] },
];

// ════════════════════════════════════════════════════════════
// HUB — pantalla de inicio con las herramientas de la suite
// ════════════════════════════════════════════════════════════
const IC = {
  camara:  '<rect x="3" y="6.5" width="18" height="13" rx="2.5"/><circle cx="12" cy="13" r="3.6"/><path d="M8.5 6.5 10 4h4l1.5 2.5"/>',
  datos:   '<ellipse cx="12" cy="5.5" rx="7.5" ry="2.8"/><path d="M4.5 5.5v6c0 1.55 3.36 2.8 7.5 2.8s7.5-1.25 7.5-2.8v-6"/><path d="M4.5 11.5v6c0 1.55 3.36 2.8 7.5 2.8s7.5-1.25 7.5-2.8v-6"/>',
  proyecto:'<rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 3.5h6a1 1 0 0 1 1 1V6a1 1 0 0 1-1 1H9a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1z"/><path d="M8.5 13l2 2 4-4.2"/>',
  entrega: '<path d="M9 3.5h5l4 4V17a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2V5.5a2 2 0 0 1 2-2z"/><path d="M14 3.5V8h4"/><path d="M6 7.5V19a2.5 2.5 0 0 0 2.5 2.5H16"/>',
  correo:  '<rect x="3" y="5.5" width="18" height="13" rx="2.5"/><path d="M3.5 7l8.5 5.5L20.5 7"/>',
  formato: '<rect x="4.5" y="3.5" width="15" height="17" rx="2"/><path d="M8 8.5h8M8 12h8M8 15.5h5"/>',
};
const TOOLS = [
  { id: 'organizador', on: true,  icon: IC.camara,   name: 'Organizador Fotográfico',
    desc: 'Empareja y clasifica las fotos del inventario por código de activo, con dashboard y reportes.' },
  { id: 'depurador',   on: false, icon: IC.datos,    name: 'Depurador de Bases',
    desc: 'Limpia y normaliza bases operativas y arma los cuadros para el informe.' },
  { id: 'proyectos',   on: false, icon: IC.proyecto, name: 'Gestor de Proyectos de Inventario',
    desc: 'Planifica y ejecuta los proyectos de levantamiento de activos en campo.' },
  { id: 'entregables', on: false, icon: IC.entrega,  name: 'Generador de Entregables',
    desc: 'Produce la documentación y los entregables del proyecto de forma automatizada.' },
  { id: 'correos',     on: false, icon: IC.correo,   name: 'Asistente de Correos',
    desc: 'Redacta y potencia correos profesionales para la comunicación con el cliente.' },
  { id: 'formatos',    on: false, icon: IC.formato,  name: 'Generador de Formatos',
    desc: 'Crea formatos y requerimientos estandarizados de manera automatizada.' },
];

function pintarHub() {
  const flecha = '<svg viewBox="0 0 24 24"><path d="M5 12h14M13 6l6 6-6 6"/></svg>';
  $('toolGrid').innerHTML = TOOLS.map((t) => `
    <div class="tool-card ${t.on ? 'activa' : 'soon'}" ${t.on ? `onclick="abrirHerramienta('${t.id}')"` : ''}>
      <div class="tool-icon"><svg viewBox="0 0 24 24">${t.icon}</svg></div>
      <div class="tool-title">${t.name}</div>
      <div class="tool-desc">${t.desc}</div>
      <div class="tool-foot">
        ${t.on ? '<span class="chip on">Disponible</span>' : '<span class="chip soon">Próximamente</span>'}
        ${t.on ? `<span class="tool-open">Abrir ${flecha}</span>` : ''}
      </div>
    </div>`).join('');
  $('stTools').textContent = TOOLS.length;
  $('stActive').textContent = TOOLS.filter((t) => t.on).length;
}

function abrirHerramienta(id) {
  const tool = TOOLS.find((t) => t.id === id);
  if (!tool || !tool.on) return;
  $('view-hub').classList.remove('visible');
  $('view-organizador').classList.add('visible');
  $('hBack').style.display = 'flex';
  $('hTool').style.display = 'inline';
  $('hTool').textContent = tool.name;
  // Compatibilidad: el organizador necesita la File System Access API.
  if (!soportado()) { $('orgApp').style.display = 'none'; $('nocompat').style.display = 'grid'; }
  else { $('orgApp').style.display = 'block'; $('nocompat').style.display = 'none'; }
  window.scrollTo({ top: 0 });
}
function volverHub() {
  $('view-organizador').classList.remove('visible');
  $('view-hub').classList.add('visible');
  $('hBack').style.display = 'none';
  $('hTool').style.display = 'none';
  window.scrollTo({ top: 0 });
}
window.abrirHerramienta = abrirHerramienta;
window.volverHub = volverHub;

// ════════════════════════════════════════════════════════════
// ARRANQUE
// ════════════════════════════════════════════════════════════
pintarHub();
$('hBack').onclick = volverHub;
iniciarOrganizador();

function iniciarOrganizador() {
  $('zInv').onclick = () => $('fInv').click();
  $('fInv').onchange = (e) => cargarExcel(e.target.files[0]);
  dragDrop('zInv', (f) => cargarExcel(f));
  $('zOrigen').onclick = () => elegirCarpeta('origen');
  $('zDestino').onclick = () => elegirCarpeta('destino');
  $('modos').onclick = (e) => {
    const m = e.target.closest('.modo'); if (!m) return;
    document.querySelectorAll('#modos .modo').forEach((x) => x.classList.remove('sel'));
    m.classList.add('sel');
  };
  $('btnSimular').onclick = () => correr(true);
  $('btnEjecutar').onclick = () => correr(false);
  $('btnEjecutarBarra').onclick = () => correr(false);
  $('selGrafico').onchange = () => pintarPanel(S.dimActiva);
  // Descargas
  $('btnResumenDim').onclick = descargarResumenDim;
  $('btnAnalisisXls').onclick = descargarAnalisisExcel;
  $('btnAnalisisPdf').onclick = descargarAnalisisPDF;
  $('btnExcel').onclick = () => descargar(S.blobExcel, `Conciliacion_${S.inv.nombre}.xlsx`);
  $('btnPDF').onclick = () => descargar(S.blobPDF, `Informe_Cobertura_${S.inv.nombre}.pdf`);
}

function dragDrop(id, cb) {
  const z = $(id);
  z.addEventListener('dragover', (e) => { e.preventDefault(); z.classList.add('drag'); });
  z.addEventListener('dragleave', () => z.classList.remove('drag'));
  z.addEventListener('drop', (e) => {
    e.preventDefault(); z.classList.remove('drag');
    const f = e.dataTransfer.files[0]; if (f) cb(f);
  });
}

function irA(paso) {
  [1, 2, 3, 4].forEach((n) => {
    $(`p${n}`).classList.toggle('visible', n === paso);
    const ws = $(`ws${n}`); ws.classList.remove('activo', 'listo');
    if (n === paso) ws.classList.add('activo'); else if (n < paso) ws.classList.add('listo');
    if (n < 4) $(`wl${n}`).classList.toggle('ok', n < paso);
  });
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
window.irA = irA;

// ── Paso 1: cargar Excel y detectar columnas ────────────────
async function cargarExcel(archivo) {
  if (!archivo) return;
  try {
    const buffer = await archivo.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    const idx = detectarFilaEncabezado(filas);

    const vistos = {};
    const enc = (filas[idx] || []).map((c, j) => {
      let n = celdaATexto(c) || `Col${j + 1}`;
      if (vistos[n]) { vistos[n]++; n = `${n}(${vistos[n]})`; } else vistos[n] = 1;
      return n;
    });

    const registros = [];
    for (let i = idx + 1; i < filas.length; i++) {
      const f = filas[i] || [];
      if (!f.some((c) => celdaATexto(c) !== '')) continue;
      const o = {}; enc.forEach((h, j) => { o[h] = f[j] ?? null; });
      registros.push(o);
    }

    const colCodigo = adivinarColumna(enc, ['BarNue', 'Codigo Activo', 'codigo activo', 'codigo', 'cod', 'placa', 'etiqueta']);
    if (!colCodigo) { alert('No encontré la columna de código del activo en el Excel.'); return; }

    const cols = {
      codigo: colCodigo,
      descripcion: adivinarColumna(enc, ['DescCatalogo', 'descripcion', 'denominacion', 'detalle', 'bien']) || null,
    };
    DIMS.forEach((d) => { cols[d.v] = adivinarColumna(enc, d.pistas) || null; });

    S.inv = { registros, nombre: archivo.name.replace(/\.[^.]+$/, ''), cols };

    const z = $('zInv');
    z.classList.add('ok');
    z.querySelector('.ic').textContent = '✅';
    z.querySelector('.zt').textContent = archivo.name;
    z.querySelector('.zs').textContent = `${num(registros.length)} activos`;
    $('inv-info').style.display = 'block';

    const tags = [`<span class="tag">Código: <b>${esc(colCodigo)}</b></span>`]
      .concat(DIMS.map((d) => `<span class="tag">${d.l}: <b>${cols[d.v] ? esc(cols[d.v]) : '—'}</b></span>`));
    $('inv-resumen').innerHTML = tags.join(' ');
    $('btn1').disabled = false;
  } catch (e) {
    alert('Error leyendo el archivo: ' + e.message);
  }
}

// ── Paso 2: carpetas ────────────────────────────────────────
async function elegirCarpeta(cual) {
  try {
    const handle = await pedirCarpeta();
    S[cual] = handle;
    const z = $(cual === 'origen' ? 'zOrigen' : 'zDestino');
    z.classList.add('ok');
    z.querySelector('.ic').textContent = '✅';
    z.querySelector('.zt').textContent = handle.name;
    z.querySelector('.zs').textContent = 'Seleccionada';
    $('btn2').disabled = !(S.origen && S.destino);
  } catch (e) {
    if (e.name !== 'AbortError') alert('No se pudo abrir la carpeta: ' + e.message);
  }
}

function mostrarConfirmacion() {
  $('confirmGrid').innerHTML = `
    <div class="citem"><div class="cl">Inventario</div><div class="cv">${esc(S.inv.nombre)}</div></div>
    <div class="citem"><div class="cl">Activos</div><div class="cv">${num(S.inv.registros.length)}</div></div>
    <div class="citem"><div class="cl">Origen</div><div class="cv">${esc(S.origen.name)}</div></div>
    <div class="citem"><div class="cl">Destino</div><div class="cv">${esc(S.destino.name)}</div></div>`;
}
window.mostrarConfirmacion = mostrarConfirmacion;

// ── Paso 3: simular / ejecutar ──────────────────────────────
async function correr(simulacion) {
  if (!simulacion && !S.simulado) { alert('Primero corre la simulación.'); return; }
  const modo = document.querySelector('#modos .modo.sel').dataset.v;
  if (!simulacion &&
      !confirm(`Se van a ${modo.toUpperCase()} las fotos hacia:\n${S.destino.name}\n\n¿Continuar?`)) return;

  bloquear(true);
  $('progWrap').style.display = 'block';   // visible ANTES de setProg (fix barra)
  $('logbox').style.display = 'block';
  $('completado').style.display = 'none';
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
    setProg(1, simulacion ? 'Simulación completada.' : '¡Proceso completado!');
    S.simulado = S.simulado || simulacion;

    const m = calcularMetricas(S.plan);
    const seg = Math.floor((Date.now() - S.tiempoInicio) / 1000);
    const tStr = seg >= 60 ? `${Math.floor(seg / 60)} min ${seg % 60} seg` : `${seg} seg`;

    if (!simulacion) {
      $('completado').style.display = 'flex';
      $('completado-msg').innerHTML = `✅ <b>${num(m.emparejadas)} fotos organizadas</b> en ${tStr}`;
      sonido();
    }

    pintarDashboard(cfg, m);
    irA(4);

    if (simulacion) {
      $('bfResumen').textContent = `${num(m.emparejadas)} fotos listas · ${pct(m.cobertura)} de cobertura`;
      $('barraFija').style.display = 'block';
      $('btnEjecutar').disabled = false;
    } else {
      $('barraFija').style.display = 'none';
    }
  } catch (e) {
    log('ERROR: ' + e.message);
    alert('Error:\n' + e.message);
    setProg(0, 'Error.');
  } finally {
    clearInterval(timer);
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
  $('tituloRes').textContent = `Dashboard · ${S.inv.nombre}`;

  $('kpis').innerHTML = [
    kpi('Total activos',    num(m.totalActivos)),
    kpi('Fotos procesadas', num(m.totalFotos)),
    kpi('Cobertura',        pct(m.cobertura), cls(m.cobertura, 60, 90), `${num(m.activosConFoto)} con foto`),
    kpi('Emparejadas',      num(m.emparejadas), 'ok'),
    kpi('Sin código',       num(m.huerfanas), m.huerfanas ? 'warn' : 'ok'),
    kpi('Sin foto',         num(m.activosSinFoto), m.activosSinFoto ? 'warn' : 'ok'),
  ].join('');

  S.totalCod = S.inv.registros.filter((r) => normalizar(celdaATexto(r[S.inv.cols.codigo]))).length;

  // Distribución de todas las dimensiones (vacío si no hay columna)
  S.datosDims = {};
  DIMS.forEach((d) => { S.datosDims[d.v] = S.inv.cols[d.v] ? calcDim(S.inv.cols[d.v]) : []; });

  $('dimTabs').innerHTML = DIMS.map((d, i) => {
    const n = S.datosDims[d.v].length;
    const badge = `<span class="badge${n ? '' : ' empty'}">${n}</span>`;
    return `<button class="dtab${i === 0 ? ' activo' : ''}" data-dim="${d.v}" onclick="cambiarDim('${d.v}',this)">${d.l}${badge}</button>`;
  }).join('');

  // Dona global de cobertura: fija, se dibuja una sola vez.
  pintarDonaCobertura(m);
  pintarPanel('familia');
}

/** Agrupa los activos por el valor de una columna del inventario. */
function calcDim(col) {
  const mapa = new Map();
  const total = S.totalCod || 1;
  for (const reg of S.inv.registros) {
    const cod = normalizar(celdaATexto(reg[S.inv.cols.codigo]));
    if (!cod) continue;
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
  $('thDim').textContent = label;
  $('gPrincipalTit').textContent = `Activos por ${label.toLowerCase()}`;
  const panel = $('panelPrincipal');

  if (!datos.length) {
    destruir('gPrincipal');
    panel.classList.add('vacio');
    $('panelPrincipal').querySelector('.gcanvas-wrap').innerHTML =
      `<canvas id="gPrincipal" width="620" height="300" style="display:none"></canvas>
       <div>Esta dimensión no está disponible en este inventario.</div>`;
    $('tbody').innerHTML = `<tr><td colspan="5" style="text-align:center;color:var(--gris);padding:22px">Sin datos para <b>${esc(label)}</b> en este inventario.</td></tr>`;
    $('resumenCards').innerHTML = '';
    return;
  }

  panel.classList.remove('vacio');
  // Reponer el canvas si venía del estado "vacío".
  if (!$('gPrincipal') || $('gPrincipal').style.display === 'none') {
    $('panelPrincipal').querySelector('.gcanvas-wrap').innerHTML =
      '<canvas id="gPrincipal" width="620" height="300"></canvas>';
  }

  pintarGraficoDim('gPrincipal', datos, $('selGrafico').value || 'barrasH');
  pintarTabla(datos);
  pintarCards(datos, label);
}

function cambiarDim(dim, btn) {
  document.querySelectorAll('.dtab').forEach((b) => b.classList.remove('activo'));
  btn.classList.add('activo');
  pintarPanel(dim);
}
window.cambiarDim = cambiarDim;

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
  const labels = top.map((d) => d.nombre);
  const valores = top.map((d) => d.activos);

  if (tipo === 'dona') {
    S.charts[id] = new Chart(canvas.getContext('2d'), {
      type: 'doughnut', plugins: [fondoBlanco],
      data: { labels, datasets: [{ data: valores, backgroundColor: PALETA.slice(0, top.length), borderWidth: 2, borderColor: '#fff' }] },
      options: { ...BASE, cutout: '55%', plugins: { legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12 } } } },
    });
    return;
  }

  const horizontal = tipo === 'barrasH';
  S.charts[id] = new Chart(canvas.getContext('2d'), {
    type: 'bar', plugins: [fondoBlanco],
    data: {
      labels,
      datasets: [{
        label: 'Activos', data: valores,
        backgroundColor: top.map((_, i) => PALETA[i % PALETA.length]),
        borderRadius: 4, barThickness: horizontal ? 14 : undefined, maxBarThickness: 34,
      }],
    },
    options: {
      ...BASE, indexAxis: horizontal ? 'y' : 'x',
      scales: {
        x: { ticks: { font: { size: 10 }, color: '#475569', maxRotation: horizontal ? 0 : 40, autoSkip: false }, grid: { color: horizontal ? '#F1F5F9' : 'transparent' } },
        y: { ticks: { font: { size: 10 }, color: '#475569', autoSkip: false }, grid: { color: horizontal ? 'transparent' : '#F1F5F9' } },
      },
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => ` ${num(c.raw)} activos` } } },
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
function pintarTabla(datos) {
  const cls = (v, b, med) => (v < b ? 'bad' : v < med ? 'warn' : 'ok');
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
  if (!datos.length) { alert(`La dimensión "${label}" no tiene datos para exportar.`); return; }
  const wb = new ExcelJS.Workbook();
  hojaDim(wb, label, datos);
  await bajarWb(wb, `Resumen_${label}_${S.inv.nombre}.xlsx`);
}

/** Excel con TODAS las dimensiones + hoja resumen. */
async function descargarAnalisisExcel() {
  if (!S.plan) { alert('Primero corre la simulación.'); return; }
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

  DIMS.forEach((d) => {
    const datos = S.datosDims[d.v] || [];
    if (datos.length) hojaDim(wb, d.l, datos);
  });
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
  if (!S.plan) { alert('Primero corre la simulación.'); return; }
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

  try {
    for (const d of DIMS) {
      const datos = S.datosDims[d.v] || [];
      if (!datos.length) continue;
      if (y > 250) { doc.addPage(); y = 20; }

      doc.setFillColor(15, 118, 110); doc.rect(M, y - 4, 3, 6, 'F');
      doc.setFont('helvetica', 'bold').setFontSize(12).setTextColor(30, 41, 59);
      doc.text(`Activos por ${d.l.toLowerCase()}`, M + 6, y + 1);
      y += 8;

      // Gráfico en canvas oculto
      const cv = document.createElement('canvas'); cv.width = 900; cv.height = 460;
      cont.appendChild(cv);
      const top = datos.slice(0, 14);
      const ch = new Chart(cv.getContext('2d'), {
        type: 'bar', plugins: [fondoBlanco],
        data: { labels: top.map((x) => x.nombre), datasets: [{ data: top.map((x) => x.activos), backgroundColor: top.map((_, i) => PALETA[i % PALETA.length]), borderRadius: 4 }] },
        options: { ...BASE, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { ticks: { color: '#475569' } }, y: { ticks: { color: '#475569', font: { size: 11 } } } } },
      });
      await new Promise((r) => requestAnimationFrame(r));
      const imgH = 58;
      doc.addImage(cv.toDataURL('image/png', 1), 'PNG', M, y, W - M * 2, imgH);
      ch.destroy();
      y += imgH + 4;

      // Tabla top 12
      doc.autoTable({
        startY: y, margin: { left: M, right: M },
        head: [[d.l, 'Activos', '% total', 'Con foto', 'Cobertura']],
        body: datos.slice(0, 12).map((x) => [x.nombre, num(x.activos), pct(x.pctTotal), num(x.conFoto), pct(x.cobertura)]),
        theme: 'grid', styles: { fontSize: 8, cellPadding: 1.8 },
        headStyles: { fillColor: [30, 41, 59], textColor: 255, fontSize: 8 },
        alternateRowStyles: { fillColor: [248, 250, 251] },
        columnStyles: { 1: { halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' } },
      });
      y = doc.lastAutoTable.finalY + 12;
    }
    descargar(doc.output('blob'), `Analisis_${S.inv.nombre}.pdf`);
  } finally {
    cont.remove();
  }
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
  $('btnEjecutarBarra').disabled = b;
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
function descargar(blob, nombre) {
  if (!blob) { alert('Aún no hay nada que descargar. Corre la simulación primero.'); return; }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = nombre; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
