/**
 * NEXOVA · Organizador de Fotos v2.1
 * Mejoras: indicadores por dimensión, gráficos seleccionables,
 * progreso en tiempo real, drag&drop, nombres completos.
 */

import {
  celdaATexto, detectarFilaEncabezado, adivinarColumna, PISTAS,
  cargarActivos, construirPlan, calcularMetricas, agruparPorFamilia,
  normalizar
} from './motor.js';
import { soportado, pedirCarpeta, escanearFotos, ejecutarPlan, guardarEnDestino } from './archivos.js';
import { generarExcel, generarPDF } from './reportes.js';

const $ = id => document.getElementById(id);
const num = n => (+n).toLocaleString('es-PE');
const pct = n => n.toFixed(1) + '%';
const esc = t => String(t).replace(/[<>&]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c]));

// ── Estado global ──────────────────────────────────────────
const S = {
  inv: { registros: [], nombre: '', cols: {} },
  origen: null, destino: null,
  plan: null, simulado: false,
  blobExcel: null, blobPDF: null,
  charts: {},
  tiempoInicio: null,
};

// ── Arranque ───────────────────────────────────────────────
if (!soportado()) {
  $('app').style.display = 'none';
  $('nocompat').style.display = 'grid';
} else {
  iniciar();
}

function iniciar() {
  // Carga Excel — clic y drag&drop
  $('zInv').onclick = () => $('fInv').click();
  $('fInv').onchange = e => cargarExcel(e.target.files[0]);
  configurarDragDrop('zInv', f => cargarExcel(f));

  // Carpetas
  $('zOrigen').onclick = () => elegirCarpeta('origen');
  $('zDestino').onclick = () => elegirCarpeta('destino');

  // Modos
  $('modos').onclick = e => {
    const m = e.target.closest('.modo');
    if (!m) return;
    document.querySelectorAll('.modo').forEach(x => x.classList.remove('sel'));
    m.classList.add('sel');
  };

  // Botones principales
  $('btnSimular').onclick       = () => correr(true);
  $('btnEjecutar').onclick      = () => correr(false);
  $('btnEjecutarBarra').onclick = () => correr(false);
  $('btnExcel').onclick = () => descargar(S.blobExcel, `Conciliacion_${S.inv.nombre}.xlsx`);
  $('btnPDF').onclick   = () => descargar(S.blobPDF,   `Informe_${S.inv.nombre}.pdf`);

  // Selectores de dimensión y tipo de gráfico
  $('selDimension').onchange = () => pintarGraficosActivos();
  $('selGrafico').onchange   = () => pintarGraficosActivos();
}

// ── Drag & Drop ────────────────────────────────────────────
function configurarDragDrop(idZona, callback) {
  const zona = $(idZona);
  zona.addEventListener('dragover', e => {
    e.preventDefault();
    zona.classList.add('drag');
  });
  zona.addEventListener('dragleave', () => zona.classList.remove('drag'));
  zona.addEventListener('drop', e => {
    e.preventDefault();
    zona.classList.remove('drag');
    const file = e.dataTransfer.files[0];
    if (file) callback(file);
  });
}

// ── Wizard ─────────────────────────────────────────────────
function irA(paso) {
  [1,2,3,4].forEach(n => {
    $(`p${n}`).classList.toggle('visible', n === paso);
    const ws = $(`ws${n}`);
    ws.classList.remove('activo','listo');
    if (n === paso) ws.classList.add('activo');
    else if (n < paso) ws.classList.add('listo');
    if (n < 4) $(`wl${n}`).classList.toggle('ok', n < paso);
  });
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
window.irA = irA;

// ── Cargar Excel ───────────────────────────────────────────
async function cargarExcel(archivo) {
  if (!archivo) return;
  try {
    progTxt('Leyendo ' + archivo.name + '…');
    const buffer = await archivo.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:null, blankrows:true });

    const idx = detectarFilaEncabezado(filas);
    const vistos = {};
    const enc = (filas[idx] || []).map((c, j) => {
      let n = celdaATexto(c) || `Col${j+1}`;
      if (vistos[n]) { vistos[n]++; n = `${n}(${vistos[n]})`; } else vistos[n] = 1;
      return n;
    });

    const registros = [];
    for (let i = idx+1; i < filas.length; i++) {
      const f = filas[i] || [];
      if (!f.some(c => celdaATexto(c) !== '')) continue;
      const o = {};
      enc.forEach((h, j) => { o[h] = f[j] ?? null; });
      registros.push(o);
    }

    // Detectar columnas — orden de prioridad adaptado al formato NEXOVA/Gate Gourmet
    const cols = {
      codigo:      adivinarColumna(enc, ['BarNue','Codigo Activo','codigo activo','cod activo','codigo','cod','placa','etiqueta','correlativo']),
      familia:     adivinarColumna(enc, ['DESCRIPCION DE FAMILIA','descripcion de familia','familia','grupo','rubro','clase','categoria']),
      subfamilia:  adivinarColumna(enc, ['DESCRIPCION DE SUB FAMILIA','descripcion de sub familia','subfamilia','sub familia','subgrupo','tipo']),
      codfamilia:  adivinarColumna(enc, ['CODIGO DE FAMILIA','codigo de familia','cod familia']),
      codsubfam:   adivinarColumna(enc, ['CODIGO DE SUB FAMILIA','codigo de sub familia','cod subfamilia']),
      descripcion: adivinarColumna(enc, PISTAS.descripcion),
      area:        adivinarColumna(enc, ['DescUbicacion','Desc Ubicacion','descripcion ubicacion','area','ubicacion descripcion','ambiente']),
      sede:        adivinarColumna(enc, ['DescCentroCosto','Desc Centro Costo','descripcion centro costo','sede','sucursal','local']),
      estado:      adivinarColumna(enc, ['estado','condicion','situacion','estado activo']),
      linea:       adivinarColumna(enc, ['linea','linea produccion','linea de produccion','proceso']),
    };

    if (!cols.codigo) {
      mostrarAlerta('No encontré la columna de código del activo. Verifica que el Excel sea correcto.');
      return;
    }

    S.inv = { registros, nombre: archivo.name.replace(/\.[^.]+$/,''), cols };

    // Actualizar zona
    const z = $('zInv');
    z.classList.add('ok');
    z.querySelector('.ic').textContent = '✅';
    z.querySelector('.zt').textContent = archivo.name;
    z.querySelector('.zs').textContent = `${num(registros.length)} activos · fila encabezado ${idx+1}`;
    $('inv-info').style.display = 'block';

    // Mostrar columnas detectadas
    const detectadas = Object.entries(cols)
      .filter(([,v]) => v)
      .map(([k,v]) => `<b>${k}:</b> ${v}`)
      .join(' &nbsp;·&nbsp; ');
    $('inv-resumen').innerHTML = detectadas;

    // Construir opciones de dimensión disponibles
    const dims = [
      { v:'familia',    l:'Familia',           ok: !!cols.familia },
      { v:'subfamilia', l:'Subfamilia',         ok: !!cols.subfamilia },
      { v:'area',       l:'Área',               ok: !!cols.area },
      { v:'sede',       l:'Sede / Centro Costo',ok: !!cols.sede },
      { v:'estado',     l:'Estado',             ok: !!cols.estado },
      { v:'linea',      l:'Línea de producción',ok: !!cols.linea },
    ].filter(d => d.ok);

    const sel = $('selDimension');
    sel.innerHTML = dims.map(d => `<option value="${d.v}">${d.l}</option>`).join('');

    if (!cols.familia) mostrarAlerta('No detecté columna de Familia. Las fotos se organizarán solo por código.');
    else ocultarAlerta();

    $('btn1').disabled = false;
    progTxt('');
  } catch(e) {
    mostrarAlerta('No se pudo leer el archivo: ' + e.message);
  }
}

function mostrarAlerta(txt) {
  $('inv-alerta').style.display = 'flex';
  $('inv-alerta-txt').textContent = txt;
  $('inv-info').style.display = 'block';
}
function ocultarAlerta() { $('inv-alerta').style.display = 'none'; }

// ── Carpetas ───────────────────────────────────────────────
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
  } catch(e) {
    if (e.name !== 'AbortError') alert('No se pudo abrir la carpeta: ' + e.message);
  }
}

// ── Confirmación ───────────────────────────────────────────
function mostrarConfirmacion() {
  $('confirmGrid').innerHTML = `
    <div class="citem"><div class="cl">Inventario</div><div class="cv">${esc(S.inv.nombre)}</div></div>
    <div class="citem"><div class="cl">Activos</div><div class="cv">${num(S.inv.registros.length)}</div></div>
    <div class="citem"><div class="cl">Carpeta de fotos</div><div class="cv">${esc(S.origen.name)}</div></div>
    <div class="citem"><div class="cl">Carpeta destino</div><div class="cv">${esc(S.destino.name)}</div></div>
  `;
}
window.mostrarConfirmacion = mostrarConfirmacion;

// ── Proceso principal ──────────────────────────────────────
async function correr(simulacion) {
  if (!simulacion && !S.simulado) {
    alert('Primero corre la simulación para ver el resultado antes de ejecutar.');
    return;
  }
  if (!simulacion) {
    const modo = document.querySelector('#modos .modo.sel').dataset.v;
    const verbo = modo === 'mover' ? 'MOVER' : 'COPIAR';
    if (!confirm(`Se van a ${verbo} las fotos hacia:\n\n${S.destino.name}\n\n¿Continuar?`)) return;
  }

  bloquear(true);
  $('logbox').style.display = 'block';
  $('progWrap').style.display = 'block';
  $('pantalla-completado').style.display = 'none';
  S.tiempoInicio = Date.now();

  // Timer de tiempo transcurrido
  const timer = simulacion ? null : setInterval(() => {
    const seg = Math.floor((Date.now() - S.tiempoInicio) / 1000);
    const mm = String(Math.floor(seg/60)).padStart(2,'0');
    const ss = String(seg%60).padStart(2,'0');
    $('tiempo-transcurrido').textContent = `${mm}:${ss}`;
  }, 1000);

  try {
    // 1. Escanear fotos
    progTxt('Escaneando carpeta de fotos…');
    progFill(0);
    const fotos = await escanearFotos(S.origen, true, S.destino,
      n => progTxt(`Escaneando… ${num(n)} fotos encontradas`));
    log(`Fotos encontradas: ${num(fotos.length)}`);
    if (!fotos.length) throw new Error('No se encontraron fotos en la carpeta de origen.');
    progFill(0.1);

    // 2. Config automática desde columnas detectadas
    const cols = S.inv.cols;
    const cfg = {
      colCodigo:      cols.codigo,
      colEnlaceInv:   null,
      colDescripcion: cols.descripcion,
      usarCatalogo:   false,
      colFamilia:     cols.familia,
      colSubfamilia:  cols.subfamilia,
      colCodfamilia:  cols.codfamilia,
      colCodsubfam:   cols.codsubfam,
      colArea:        cols.area,
      colSede:        cols.sede,
      colEstado:      cols.estado,
      colLinea:       cols.linea,
      niveles:        cols.subfamilia ? 2 : 1,
      moverHuerfanas: true,
      modo:           document.querySelector('#modos .modo.sel').dataset.v,
      cliente:        S.inv.nombre,
    };

    // 3. Cruzar datos
    progTxt('Procesando inventario…');
    const { activos, alertas } = cargarActivos(S.inv.registros, [], cfg);
    log(`Activos cargados: ${num(activos.size)}`);
    S.plan = construirPlan(fotos, activos, alertas, cfg);
    progFill(0.2);

    // 4. Ejecutar si no es simulación
    if (!simulacion) {
      const total = S.plan.items.filter(i => i.carpetas).length;
      await ejecutarPlan(S.plan, S.destino, cfg.modo,
        (h, t, item) => {
          const p = 0.2 + (h/t)*0.75;
          progFill(p);
          const seg = Math.floor((Date.now() - S.tiempoInicio) / 1000);
          const restante = seg > 0 ? Math.round((seg / h) * (t - h)) : '...';
          progTxt(`${cfg.modo === 'mover' ? 'Moviendo' : 'Copiando'} ${num(h)} de ${num(t)} · Faltan ~${restante}s`);
        });
      log('Proceso completado.');
    }
    progFill(0.95);

    // 5. Generar reportes
    progTxt('Generando reportes…');
    S.blobExcel = await generarExcel(S.plan, cfg, simulacion);
    S.blobPDF   = await generarPDF(S.plan, cfg, {
      estado: $('gEstado'), cobertura: $('gCobertura')
    }, simulacion);

    if (!simulacion) {
      await guardarEnDestino(S.destino, 'Conciliacion.xlsx', S.blobExcel);
      await guardarEnDestino(S.destino, 'Informe_Cobertura.pdf', S.blobPDF);
      log('Reportes guardados en la carpeta destino.');
    }

    progFill(1);
    S.simulado = simulacion;

    // 6. Mostrar resultado
    pintarResultado(simulacion);
    irA(4);

    // Pantalla de completado
    const m = calcularMetricas(S.plan);
    const seg = Math.floor((Date.now() - S.tiempoInicio) / 1000);
    const mm = Math.floor(seg/60), ss = seg%60;
    const tiempoStr = mm > 0 ? `${mm} min ${ss} seg` : `${ss} seg`;

    if (!simulacion) {
      $('completado-msg').innerHTML =
        `✅ ¡Listo! Se organizaron <b>${num(m.emparejadas)} fotos</b> en <b>${tiempoStr}</b>`;
      $('pantalla-completado').style.display = 'flex';
      // Sonido de notificación
      try {
        const ctx = new AudioContext();
        const o = ctx.createOscillator();
        const g = ctx.createGain();
        o.connect(g); g.connect(ctx.destination);
        o.frequency.setValueAtTime(880, ctx.currentTime);
        o.frequency.setValueAtTime(1100, ctx.currentTime + 0.1);
        g.gain.setValueAtTime(0.3, ctx.currentTime);
        g.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.4);
        o.start(); o.stop(ctx.currentTime + 0.4);
      } catch {}
    }

    if (simulacion) {
      $('bfResumen').textContent = `${num(m.emparejadas)} fotos listas · cobertura ${pct(m.cobertura)}`;
      $('barraFija').style.display = 'block';
      $('btnEjecutar').disabled = false;
    } else {
      $('barraFija').style.display = 'none';
    }

  } catch(e) {
    log('ERROR: ' + e.message);
    alert('Error: ' + e.message);
    progTxt('Se detuvo por un error.');
  } finally {
    if (timer) clearInterval(timer);
    bloquear(false);
  }
}

// ── Resultados ─────────────────────────────────────────────
function pintarResultado(simulacion) {
  const m = calcularMetricas(S.plan);
  const cls = (v, b, med) => v < b ? 'bad' : v < med ? 'warn' : 'ok';

  $('tituloRes').textContent = simulacion
    ? 'Simulación — sin tocar archivos'
    : '✅ Proceso ejecutado';

  // KPIs
  $('kpis').innerHTML = `
    ${kpi('Activos en inventario', num(m.totalActivos))}
    ${kpi('Fotos procesadas',      num(m.totalFotos))}
    ${kpi('Cobertura fotográfica', pct(m.cobertura), cls(m.cobertura,60,90), `${num(m.activosConFoto)} de ${num(m.totalActivos)}`)}
    ${kpi('Fotos emparejadas',     num(m.emparejadas), 'ok')}
    ${kpi('Sin código reconocido', num(m.huerfanas),   m.huerfanas ? 'warn' : 'ok', '→ carpeta _SIN_CLASIFICAR')}
    ${kpi('Activos sin foto',      num(m.activosSinFoto), m.activosSinFoto ? 'warn' : 'ok', 'pendientes de fotografiar')}
  `;

  // Avisos
  const avisos = [];
  if (m.sinFamilia > m.totalActivos * 0.2)
    avisos.push(`<div class="aviso warn">⚠️<div><b>${num(m.sinFamilia)} activos sin familia.</b> Revisa la columna de familia en el Excel.</div></div>`);
  $('avisos').innerHTML = avisos.join('');

  // Guardar datos para gráficos y tabla por dimensión
  S.datosPlan = {
    familias:    agruparPorDimension('familia'),
    subfamilias: agruparPorDimension('subfamilia'),
    areas:       agruparPorDimension('area'),
    sedes:       agruparPorDimension('sede'),
    estados:     agruparPorDimension('estado'),
    lineas:      agruparPorDimension('linea'),
  };

  pintarGraficosActivos();
  pintarTablaResumen();
}

function agruparPorDimension(dim) {
  const cols = S.inv.cols;
  const colDim = cols[dim];
  if (!colDim) return [];

  const mapa = new Map();
  for (const reg of S.inv.registros) {
    const codActivo = normalizar(celdaATexto(reg[cols.codigo]));
    if (!codActivo) continue;
    const val = celdaATexto(reg[colDim]) || 'SIN DATO';

    // Código de la dimensión si existe
    let codigo = '';
    if (dim === 'familia' && cols.codfamilia) codigo = celdaATexto(reg[cols.codfamilia]);
    if (dim === 'subfamilia' && cols.codsubfam) codigo = celdaATexto(reg[cols.codsubfam]);

    const etiqueta = codigo ? `${val} (${codigo})` : val;
    if (!mapa.has(etiqueta)) mapa.set(etiqueta, { nombre: etiqueta, activos: 0, conFoto: 0 });
    const g = mapa.get(etiqueta);
    g.activos++;
    if (S.plan && S.plan.codigosConFoto.has(codActivo)) g.conFoto++;
  }

  return [...mapa.values()]
    .map(g => ({ ...g, sinFoto: g.activos - g.conFoto, cobertura: g.activos ? (g.conFoto/g.activos)*100 : 0 }))
    .sort((a,b) => b.activos - a.activos);
}

// ── Gráficos ───────────────────────────────────────────────
const TEAL = '#0F766E', AMBAR = '#D97706', ROJO = '#B91C1C', GRIS = '#CBD5E1';

function colorBarra(v) { return v < 60 ? ROJO : v < 90 ? AMBAR : TEAL; }

function destruirChart(id) {
  if (S.charts[id]) { S.charts[id].destroy(); delete S.charts[id]; }
}

function fondoBlanco() {
  return { id:'fondoBlanco', beforeDraw(c) {
    const ctx = c.ctx; ctx.save();
    ctx.globalCompositeOperation = 'destination-over';
    ctx.fillStyle = '#fff'; ctx.fillRect(0,0,c.width,c.height);
    ctx.restore();
  }};
}

function pintarGraficosActivos() {
  if (!S.datosPlan) return;
  const dim = $('selDimension').value;
  const tipo = $('selGrafico').value;
  const datos = S.datosPlan[dim + 's'] || S.datosPlan[dim] || [];
  const top = datos.slice(0, 12);

  // Gráfico de cobertura
  destruirChart('cobertura');
  const ctxCob = $('gCobertura').getContext('2d');

  if (tipo === 'dona') {
    S.charts.cobertura = new Chart(ctxCob, {
      type: 'doughnut',
      plugins: [fondoBlanco()],
      data: {
        labels: top.map(d => d.nombre),
        datasets: [{ data: top.map(d => d.activos), backgroundColor: [TEAL,'#0D9488','#14B8A6','#2DD4BF','#5EEAD4','#99F6E4','#CCFBF1',AMBAR,'#FCD34D',ROJO,'#FCA5A5',GRIS], borderWidth: 2, borderColor: '#fff' }]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12 } } } }
    });
  } else {
    const horizontal = tipo === 'barrasH' || tipo === 'tabla';
    S.charts.cobertura = new Chart(ctxCob, {
      type: 'bar',
      plugins: [fondoBlanco()],
      data: {
        labels: top.map(d => d.nombre),
        datasets: [{ label: '% cobertura', data: top.map(d => +d.cobertura.toFixed(1)),
          backgroundColor: top.map(d => colorBarra(d.cobertura)), borderRadius: 3 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        indexAxis: tipo === 'barrasH' ? 'y' : 'x',
        scales: {
          x: { max: tipo === 'barrasH' ? 100 : undefined, ticks: { callback: v => tipo === 'barrasH' ? v+'%' : v, font:{size:10} }, grid:{color:'#E2E8F0'} },
          y: { ticks: { font:{size:10} }, grid: { display: tipo !== 'barrasH' } }
        },
        plugins: { legend: { display: false } }
      }
    });
  }

  // Gráfico de estado (dona fija)
  destruirChart('estado');
  const m = calcularMetricas(S.plan);
  S.charts.estado = new Chart($('gEstado').getContext('2d'), {
    type: 'doughnut',
    plugins: [fondoBlanco()],
    data: {
      labels: ['Con foto', 'Sin foto'],
      datasets: [{ data: [m.activosConFoto, m.activosSinFoto], backgroundColor: [TEAL, GRIS], borderWidth: 2, borderColor: '#fff' }]
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: '65%',
      plugins: { legend: { position: 'bottom', labels: { font:{size:11}, boxWidth:12 } } } }
  });

  // Tabla resumen de la dimensión seleccionada
  pintarTablaResumen(top);
}

function pintarTablaResumen(datos) {
  if (!datos) {
    const dim = $('selDimension')?.value || 'familia';
    datos = (S.datosPlan?.[dim+'s'] || S.datosPlan?.[dim] || []).slice(0,50);
  }
  const filas = datos.map(d => {
    const c = d.cobertura;
    const cls = c < 60 ? 'bad' : c < 90 ? 'warn' : 'ok';
    return `<tr>
      <td>${esc(d.nombre)}</td>
      <td class="r">${num(d.activos)}</td>
      <td class="r">${num(d.conFoto)}</td>
      <td class="r">${num(d.sinFoto)}</td>
      <td class="r"><span class="pill ${cls}">${pct(c)}</span></td>
    </tr>`;
  }).join('');
  $('tbody').innerHTML = filas;
}

// ── Descarga de resumen en Excel ───────────────────────────
async function descargarResumen() {
  const dim = $('selDimension').value;
  const datos = (S.datosPlan?.[dim+'s'] || S.datosPlan?.[dim] || []);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Resumen');
  ws.columns = [
    { header: 'Nombre', key: 'n', width: 40 },
    { header: 'Activos', key: 'a', width: 12 },
    { header: 'Con foto', key: 'c', width: 12 },
    { header: 'Sin foto', key: 's', width: 12 },
    { header: 'Cobertura %', key: 'p', width: 14 },
  ];
  const fila1 = ws.getRow(1);
  fila1.eachCell(c => { c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1E293B'}}; c.font={bold:true,color:{argb:'FFFFFFFF'}}; });
  datos.forEach(d => ws.addRow({ n:d.nombre, a:d.activos, c:d.conFoto, s:d.sinFoto, p:+d.cobertura.toFixed(1) }));
  const buf = await wb.xlsx.writeBuffer();
  descargar(new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), `Resumen_${dim}.xlsx`);
}
window.descargarResumen = descargarResumen;

// ── Utilidades ─────────────────────────────────────────────
const kpi = (l,v,cls='',sub='') =>
  `<div class="kpi ${cls}"><div class="kl">${l}</div><div class="kv">${v}</div>${sub?`<div class="ks">${sub}</div>`:''}</div>`;

function bloquear(b) {
  $('btnSimular').disabled = b;
  $('btnEjecutar').disabled = b || !S.simulado;
  $('btnEjecutarBarra').disabled = b;
}
const progTxt  = t => { if($('progTxt'))  $('progTxt').textContent = t; };
const progFill = f => { if($('progFill')) $('progFill').style.width = `${Math.min(100,f*100)}%`; };
function log(t) {
  const el = $('logbox'); if(!el) return;
  el.textContent += new Date().toLocaleTimeString('es-PE') + '  ' + t + '\n';
  el.scrollTop = el.scrollHeight;
}
function descargar(blob, nombre) {
  if (!blob) return;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = nombre; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
