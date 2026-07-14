/**
 * ============================================================================
 * app.js — INTERFAZ Y ORQUESTACION
 * ============================================================================
 * Solo estado de pantalla y eventos. La logica vive en motor.js, el disco en
 * archivos.js y los entregables en reportes.js.
 * ============================================================================
 */

import {
  celdaATexto, detectarFilaEncabezado, adivinarColumna, PISTAS,
  cargarActivos, construirPlan, calcularMetricas, agruparPorFamilia,
} from './motor.js';
import { soportado, pedirCarpeta, escanearFotos, ejecutarPlan, guardarEnDestino } from './archivos.js';
import { graficoEstado, graficoCobertura, graficoActivos, generarExcel, generarPDF } from './reportes.js';

const $ = (id) => document.getElementById(id);
const num = (n) => n.toLocaleString('es-PE');
const CLAVE_CFG = 'nexova_organizador_v1';

// Enlace entre cada combo de la pantalla y su clave en la config guardada
const CLAVE_COMBO = {
  cCodigo: 'colCodigo', cEnlaceInv: 'colEnlaceInv', cDesc: 'colDescripcion',
  cEnlaceCat: 'colEnlaceCat', cFamilia: 'colFamilia', cSubfamilia: 'colSubfamilia',
};

// --- Estado global de la pantalla
const S = {
  inv: { filas: [], hojas: [], enc: [], registros: [] },
  cat: { filas: [], hojas: [], enc: [], registros: [] },
  origen: null, destino: null,
  fotos: [], plan: null, simulado: false,
  blobExcel: null, blobPDF: null, canvases: {},
};

// ===========================================================================
// ARRANQUE — deteccion de capacidad ANTES de mostrar nada
// ===========================================================================
if (!soportado()) {
  $('app').style.display = 'none';
  $('acciones').style.display = 'none';
  $('nocompat').style.display = 'grid';
} else {
  iniciar();
}

function iniciar() {
  // Selectores de archivo
  $('zInv').onclick = () => $('fInv').click();
  $('zCat').onclick = () => $('fCat').click();
  $('fInv').onchange = (e) => cargarArchivo(e.target.files[0], 'inv');
  $('fCat').onchange = (e) => cargarArchivo(e.target.files[0], 'cat');

  $('hInv').onchange = () => releerHoja('inv', true);
  $('hCat').onchange = () => releerHoja('cat', true);
  $('filaInv').onchange = () => releerHoja('inv', false);
  $('filaCat').onchange = () => releerHoja('cat', false);

  // Selectores de carpeta (obligatorio: desde un click real del usuario)
  $('zOrigen').onclick = () => elegirCarpeta('origen');
  $('zDestino').onclick = () => elegirCarpeta('destino');

  $('sinCat').onchange = alternarCatalogo;
  grupoOpciones('modo');
  grupoOpciones('niveles');

  $('btnSimular').onclick = () => correr(true);
  $('btnEjecutar').onclick = () => correr(false);
  $('btnExcel').onclick = () => descargar(S.blobExcel, `Conciliacion_${nombreCliente()}.xlsx`);
  $('btnPDF').onclick = () => descargar(S.blobPDF, `Informe_Cobertura_${nombreCliente()}.pdf`);

  restaurarConfig();
}

// ===========================================================================
// LECTURA DE EXCEL
// ===========================================================================

async function cargarArchivo(archivo, tipo) {
  if (!archivo) return;
  try {
    estado(`Leyendo ${archivo.name}…`);
    const buffer = await archivo.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });

    const st = S[tipo];
    st.nombre = archivo.name;
    st.wb = wb;
    st.hojas = wb.SheetNames;

    const sel = $(tipo === 'inv' ? 'hInv' : 'hCat');
    sel.innerHTML = st.hojas.map((h) => `<option>${h}</option>`).join('');
    sel.disabled = false;

    const z = $(tipo === 'inv' ? 'zInv' : 'zCat');
    z.classList.add('lista');
    z.querySelector('.ic').textContent = '✅';
    z.querySelector('.t').textContent = archivo.name;
    z.querySelector('.s').textContent = `${st.hojas.length} hoja(s)`;

    releerHoja(tipo, true);
    estado('Archivo cargado. Revisa el mapeo de columnas.');
  } catch (e) {
    alert(`No se pudo leer el archivo:\n${e.message}`);
  }
}

/** Convierte la hoja activa en filas + encabezados. @param auto ¿autodetectar la fila? */
function releerHoja(tipo, auto) {
  const st = S[tipo];
  if (!st.wb) return;

  const hoja = $(tipo === 'inv' ? 'hInv' : 'hCat').value;
  const ws = st.wb.Sheets[hoja];
  // blankrows:true para que el número de fila coincida con el que se ve en Excel
  const filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  st.filas = filas;

  const inputFila = $(tipo === 'inv' ? 'filaInv' : 'filaCat');
  const idx = auto ? detectarFilaEncabezado(filas) : Math.max(0, (+inputFila.value || 1) - 1);
  inputFila.value = idx + 1;
  inputFila.disabled = false;

  // Encabezados (desduplicados: los Excel de cliente repiten nombres de columna)
  const vistos = {};
  st.enc = (filas[idx] || []).map((c, j) => {
    let n = celdaATexto(c) || `Columna ${j + 1}`;
    if (vistos[n]) { vistos[n]++; n = `${n} (${vistos[n]})`; } else { vistos[n] = 1; }
    return n;
  });

  st.registros = [];
  for (let i = idx + 1; i < filas.length; i++) {
    const f = filas[i] || [];
    if (!f.some((c) => celdaATexto(c) !== '')) continue;   // fila vacía
    const o = {};
    st.enc.forEach((h, j) => { o[h] = f[j] ?? null; });
    st.registros.push(o);
  }

  poblarCombos(tipo, auto);
  log(`${tipo === 'inv' ? 'Inventario' : 'Catálogo'}: ${num(st.registros.length)} filas · encabezado en la fila ${idx + 1}`);
}

function poblarCombos(tipo, auto) {
  const st = S[tipo];
  const ops = ['<option value=""></option>']
    .concat(st.enc.map((h) => `<option>${h.replace(/</g, '&lt;')}</option>`)).join('');

  const combos = tipo === 'inv'
    ? [['cCodigo', PISTAS.codigo], ['cEnlaceInv', PISTAS.enlaceInv], ['cDesc', PISTAS.descripcion]]
    : [['cEnlaceCat', PISTAS.enlaceCat], ['cFamilia', PISTAS.familia], ['cSubfamilia', PISTAS.subfamilia]];

  for (const [id, pistas] of combos) {
    const sel = $(id);
    sel.innerHTML = ops;
    sel.disabled = false;
    // Prioridad: 1) el mapeo del cliente anterior si esa columna existe aqui,
    //            2) la adivinanza automatica. Asi, si repites formato de Excel,
    //               la herramienta ya viene configurada sola.
    const previa = S.cfgPrevia?.[CLAVE_COMBO[id]];
    if (previa && st.enc.includes(previa)) sel.value = previa;
    else if (auto || !sel.value) sel.value = adivinarColumna(st.enc, pistas);
  }

  // Si el inventario ya trae familia/subfamilia, esos combos se llenan desde el inventario
  if (tipo === 'inv' && $('sinCat').checked) alternarCatalogo();
}

function alternarCatalogo() {
  const sinCat = $('sinCat').checked;
  $('bloqueCat').style.display = sinCat ? 'none' : 'block';

  if (sinCat && S.inv.enc.length) {
    // Las columnas de familia salen del propio inventario
    const ops = ['<option value=""></option>']
      .concat(S.inv.enc.map((h) => `<option>${h.replace(/</g, '&lt;')}</option>`)).join('');
    for (const [id, pistas] of [['cFamilia', PISTAS.familia], ['cSubfamilia', PISTAS.subfamilia]]) {
      const sel = $(id);
      sel.innerHTML = ops;
      sel.disabled = false;
      sel.value = adivinarColumna(S.inv.enc, pistas);
    }
    // Se muestran dentro del paso 2 aunque el bloque del archivo esté oculto
    $('bloqueCat').style.display = 'block';
    $('zCat').style.display = 'none';
    $('hCat').closest('.campo').style.display = 'none';
    $('filaCat').closest('.campo').style.display = 'none';
    $('cEnlaceCat').closest('.campo').style.display = 'none';
  } else {
    // Se devuelve el control al CSS en lugar de forzar display inline
    $('zCat').style.display = '';
    ['hCat', 'filaCat', 'cEnlaceCat'].forEach((id) => {
      const c = $(id).closest('.campo');
      if (c) c.style.display = '';
    });
  }
}

// ===========================================================================
// CARPETAS
// ===========================================================================

async function elegirCarpeta(cual) {
  try {
    const handle = await pedirCarpeta();
    S[cual] = handle;
    const z = $(cual === 'origen' ? 'zOrigen' : 'zDestino');
    z.classList.add('lista');
    z.querySelector('.ic').textContent = '✅';
    z.querySelector('.s').textContent = handle.name;
    S.simulado = false;
    $('btnEjecutar').disabled = true;
  } catch (e) {
    if (e.name !== 'AbortError') alert(`No se pudo abrir la carpeta:\n${e.message}`);
  }
}

// ===========================================================================
// CONFIG Y VALIDACION
// ===========================================================================

function leerConfig() {
  return {
    colCodigo: $('cCodigo').value,
    colEnlaceInv: $('cEnlaceInv').value,
    colDescripcion: $('cDesc').value,
    usarCatalogo: !$('sinCat').checked,
    colEnlaceCat: $('cEnlaceCat').value,
    colFamilia: $('cFamilia').value,
    colSubfamilia: $('cSubfamilia').value,
    recursivo: $('recursivo').checked,
    modo: document.querySelector('#modo .op.sel').dataset.v,
    niveles: +document.querySelector('#niveles .op.sel').dataset.v,
    moverHuerfanas: $('huerfanas').checked,
    cliente: $('cliente').value.trim(),
  };
}

async function validar(cfg) {
  const e = [];
  if (!S.inv.registros.length) e.push('Carga el Excel del inventario.');
  if (!cfg.colCodigo) e.push('Indica la columna del CÓDIGO del activo.');
  if (cfg.usarCatalogo) {
    if (!S.cat.registros.length) e.push('Carga el Excel del catálogo.');
    if (!cfg.colEnlaceInv || !cfg.colEnlaceCat) e.push('Indica las columnas de ENLACE entre inventario y catálogo.');
  }
  if (!cfg.colFamilia) e.push('Indica la columna FAMILIA.');
  if (!S.origen) e.push('Elige la carpeta de fotos.');
  if (!S.destino) e.push('Elige la carpeta destino.');
  if (S.origen && S.destino && await S.origen.isSameEntry(S.destino)) {
    e.push('El destino no puede ser la misma carpeta de origen.');
  }
  return e;
}

// ===========================================================================
// PROCESO
// ===========================================================================

async function correr(simulacion) {
  const cfg = leerConfig();
  const errores = await validar(cfg);
  if (errores.length) { alert('Faltan datos:\n\n· ' + errores.join('\n· ')); return; }

  if (!simulacion) {
    if (!S.simulado) { alert('Corre primero la SIMULACIÓN. Así ves el resultado sin tocar un solo archivo.'); return; }
    const verbo = cfg.modo === 'mover' ? 'MOVER' : 'COPIAR';
    if (!confirm(`Se van a ${verbo} ${num(S.plan.items.length)} fotos hacia:\n\n${S.destino.name}\n\n¿Continuar?`)) return;
  }

  bloquear(true);
  guardarConfig(cfg);
  $('log').style.display = 'block';

  try {
    // 1 · Escanear el disco
    estado('Escaneando la carpeta de fotos…');
    S.fotos = await escanearFotos(S.origen, cfg.recursivo, S.destino,
      (n) => estado(`Escaneando… ${num(n)} fotos encontradas`));
    log(`Fotos encontradas: ${num(S.fotos.length)}`);
    if (!S.fotos.length) throw new Error('No se encontró ninguna foto en la carpeta de origen.');

    // 2 · Cruzar datos y construir el plan (esto NO toca el disco)
    estado('Cruzando inventario con catálogo…');
    const { activos, alertas } = cargarActivos(S.inv.registros, S.cat.registros, cfg);
    if (!activos.size) throw new Error('No se leyó ningún activo. Revisa la columna de CÓDIGO.');
    log(`Activos cargados: ${num(activos.size)}`);

    S.plan = construirPlan(S.fotos, activos, alertas, cfg);

    // 3 · Ejecutar (solo si no es simulación)
    if (!simulacion) {
      const t0 = performance.now();
      await ejecutarPlan(S.plan, S.destino, cfg.modo,
        (hechos, total, item) => {
          progreso(hechos / total);
          estado(`${cfg.modo === 'mover' ? 'Moviendo' : 'Copiando'} ${num(hechos)} de ${num(total)} — ${item.nombre}`);
        });
      const seg = ((performance.now() - t0) / 1000).toFixed(1);
      log(`Proceso terminado en ${seg} s.`);
    }

    // 4 · Resultados y entregables
    estado('Generando reportes…');
    pintarResultados(S.plan, cfg, simulacion);

    S.blobExcel = await generarExcel(S.plan, cfg, simulacion);
    S.blobPDF = await generarPDF(S.plan, cfg, S.canvases, simulacion);

    // Al ejecutar, los reportes quedan guardados junto a las fotos organizadas
    if (!simulacion) {
      await guardarEnDestino(S.destino, 'Conciliacion.xlsx', S.blobExcel);
      await guardarEnDestino(S.destino, 'Informe_Cobertura_Fotografica.pdf', S.blobPDF);
      log('Reportes guardados en la carpeta destino.');
    }

    S.simulado = simulacion ? true : false;
    $('btnEjecutar').disabled = !simulacion;
    progreso(1);
    estado(simulacion
      ? 'Simulación lista. Revisa el resultado y, si está correcto, pulsa Ejecutar.'
      : '✅ Proceso completado.');
    $('resultados').scrollIntoView({ behavior: 'smooth', block: 'start' });

  } catch (e) {
    console.error(e);
    log('ERROR: ' + e.message);
    alert('Error:\n\n' + e.message);
    estado('Se detuvo por un error.');
  } finally {
    bloquear(false);
  }
}

// ===========================================================================
// RESULTADOS
// ===========================================================================

function pintarResultados(plan, cfg, simulacion) {
  const m = calcularMetricas(plan);
  const familias = agruparPorFamilia(plan);

  $('resultados').style.display = 'block';
  $('tResultado').textContent = simulacion
    ? 'Resultado de la simulación (no se tocó ningún archivo)'
    : `Proceso ejecutado — modo ${cfg.modo}`;

  const clase = (v, malo, medio) => (v < malo ? 'bad' : v < medio ? 'warn' : 'ok');
  $('kpis').innerHTML = `
    ${kpi('Activos en inventario', num(m.totalActivos))}
    ${kpi('Fotos procesadas', num(m.totalFotos))}
    ${kpi('Fotos emparejadas', num(m.emparejadas), 'ok', `${((m.emparejadas / (m.totalFotos || 1)) * 100).toFixed(0)}% del total`)}
    ${kpi('Fotos sin código', num(m.huerfanas), m.huerfanas ? 'warn' : 'ok', '→ _SIN_CLASIFICAR')}
    ${kpi('Cobertura fotográfica', m.cobertura.toFixed(1) + '%', clase(m.cobertura, 60, 90), `${num(m.activosConFoto)} de ${num(m.totalActivos)} activos`)}
    ${kpi('Activos sin foto', num(m.activosSinFoto), m.activosSinFoto ? 'warn' : 'ok', 'faltan fotografiar')}
    ${m.errores ? kpi('Errores', num(m.errores), 'bad', 'ver hoja ALERTAS') : ''}
  `;

  // Semáforo: el modo de fallo más probable de todo el sistema
  const avisos = [];
  if (m.sinFamilia > 0) {
    const grave = m.pctSinFamilia > 20;
    avisos.push(`<div class="aviso ${grave ? 'bad' : 'warn'}">⚠️<div>
      <b>${num(m.sinFamilia)} activos (${m.pctSinFamilia.toFixed(0)}%) quedaron SIN FAMILIA.</b>
      ${grave ? 'Casi con certeza la columna de <b>ENLACE</b> entre inventario y catálogo está mal mapeada. Revísala y vuelve a simular antes de ejecutar.'
              : 'Esos activos no encontraron su familia en el catálogo. Revisa la hoja ALERTAS del Excel.'}
    </div></div>`);
  }
  if (m.huerfanas > m.totalFotos * 0.3) {
    avisos.push(`<div class="aviso warn">📷<div>
      <b>${((m.huerfanas / m.totalFotos) * 100).toFixed(0)}% de las fotos no corresponde a ningún código.</b>
      Revisa que los nombres de archivo coincidan con los códigos del inventario.
    </div></div>`);
  }
  $('avisos').innerHTML = avisos.join('');

  // Gráficos
  S.canvases = { estado: $('gEstado'), cobertura: $('gCobertura'), activos: $('gActivos') };
  graficoEstado(S.canvases.estado, m);
  graficoCobertura(S.canvases.cobertura, familias);
  graficoActivos(S.canvases.activos, familias);

  // Tabla familia / subfamilia
  const filas = [];
  for (const f of familias) {
    for (const s of f.subfamilias) {
      const c = s.cobertura;
      filas.push(`<tr>
        <td><b>${esc(f.familia)}</b></td><td>${esc(s.subfamilia)}</td>
        <td class="r">${num(s.activos)}</td><td class="r">${num(s.conFoto)}</td>
        <td class="r">${num(s.sinFoto)}</td>
        <td class="r"><span class="pill ${clase(c, 60, 90)}">${c.toFixed(0)}%</span></td>
        <td class="r">${num(s.fotos)}</td>
      </tr>`);
    }
  }
  $('tbody').innerHTML = filas.join('');
}

const kpi = (l, v, cls = '', sub = '') =>
  `<div class="kpi ${cls}"><div class="l">${l}</div><div class="v">${v}</div>${sub ? `<div class="sub">${sub}</div>` : ''}</div>`;

const esc = (t) => String(t).replace(/[<>&]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c]));

// ===========================================================================
// UTILIDADES DE PANTALLA
// ===========================================================================

function grupoOpciones(id) {
  const grupo = $(id);
  grupo.onclick = (e) => {
    const op = e.target.closest('.op');
    if (!op) return;
    grupo.querySelectorAll('.op').forEach((o) => o.classList.remove('sel'));
    op.classList.add('sel');
    S.simulado = false;
    $('btnEjecutar').disabled = true;
  };
}

const estado = (t) => { $('txtEstado').innerHTML = t; };
const progreso = (f) => { $('barra').style.width = `${Math.min(100, f * 100)}%`; };
function log(t) {
  const el = $('log');
  el.textContent += `${new Date().toLocaleTimeString('es-PE')}  ${t}\n`;
  el.scrollTop = el.scrollHeight;
}
function bloquear(b) {
  $('btnSimular').disabled = b;
  $('btnEjecutar').disabled = b || !S.simulado;
}
function descargar(blob, nombre) {
  if (!blob) return;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = nombre;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
const nombreCliente = () =>
  ($('cliente').value.trim() || 'Inventario').replace(/[^\w\sáéíóúñÁÉÍÓÚÑ-]/g, '').replace(/\s+/g, '_');

// ===========================================================================
// PERSISTENCIA DEL MAPEO — si el próximo cliente usa el mismo formato,
// la herramienta ya viene configurada.
// ===========================================================================

function guardarConfig(cfg) {
  try { localStorage.setItem(CLAVE_CFG, JSON.stringify(cfg)); } catch { /* modo incógnito */ }
}

function restaurarConfig() {
  let cfg;
  try { cfg = JSON.parse(localStorage.getItem(CLAVE_CFG) || 'null'); } catch { return; }
  if (!cfg) return;

  $('sinCat').checked = !cfg.usarCatalogo;
  $('recursivo').checked = cfg.recursivo !== false;
  $('huerfanas').checked = cfg.moverHuerfanas !== false;
  if (cfg.cliente) $('cliente').value = cfg.cliente;

  for (const [grupo, valor] of [['modo', cfg.modo], ['niveles', String(cfg.niveles)]]) {
    const op = document.querySelector(`#${grupo} .op[data-v="${valor}"]`);
    if (op) {
      document.querySelectorAll(`#${grupo} .op`).forEach((o) => o.classList.remove('sel'));
      op.classList.add('sel');
    }
  }
  alternarCatalogo();
  // El mapeo de columnas se re-aplica al cargar los archivos (ver poblarCombos)
  S.cfgPrevia = cfg;
  log('Configuración anterior restaurada.');
}
