/**
 * NEXOVA · Organizador de Fotos v2
 * Versión simplificada — un solo Excel, detección automática
 * Adaptado para inventarios Gate Gourmet / formato NEXOVA
 */

import { celdaATexto, detectarFilaEncabezado, adivinarColumna, PISTAS,
         cargarActivos, construirPlan, calcularMetricas, agruparPorFamilia } from './motor.js';
import { soportado, pedirCarpeta, escanearFotos, ejecutarPlan, guardarEnDestino } from './archivos.js';
import { graficoEstado, graficoCobertura, generarExcel, generarPDF } from './reportes.js';

const $ = id => document.getElementById(id);
const num = n => n.toLocaleString('es-PE');

// Estado global
const S = {
  inv: { registros: [], nombre: '' },
  origen: null, destino: null,
  plan: null, simulado: false,
  blobExcel: null, blobPDF: null
};

// ── ARRANQUE ──────────────────────────────────────────────
if (!soportado()) {
  $('app').style.display = 'none';
  $('nocompat').style.display = 'grid';
} else {
  iniciar();
}

function iniciar() {
  $('zInv').onclick    = () => $('fInv').click();
  $('fInv').onchange   = e  => cargarExcel(e.target.files[0]);
  $('zOrigen').onclick = () => elegirCarpeta('origen');
  $('zDestino').onclick = () => elegirCarpeta('destino');

  // Modos copiar/mover
  $('modos').onclick = e => {
    const m = e.target.closest('.modo');
    if (!m) return;
    document.querySelectorAll('.modo').forEach(x => x.classList.remove('sel'));
    m.classList.add('sel');
  };

  $('btnSimular').onclick      = () => correr(true);
  $('btnEjecutar').onclick     = () => correr(false);
  $('btnEjecutarBarra').onclick = () => correr(false);
  $('btnExcel').onclick = () => descargar(S.blobExcel, `Conciliacion_${S.inv.nombre}.xlsx`);
  $('btnPDF').onclick   = () => descargar(S.blobPDF,   `Informe_${S.inv.nombre}.pdf`);
}

// ── WIZARD ────────────────────────────────────────────────
function irA(paso) {
  [1,2,3,4].forEach(n => {
    $(`p${n}`).classList.toggle('visible', n === paso);
    const ws = $(`ws${n}`);
    ws.classList.remove('activo','listo');
    if (n === paso) ws.classList.add('activo');
    else if (n < paso) ws.classList.add('listo');
    if (n < 4) $(`wl${n}`).classList.toggle('ok', n < paso);
  });
}
window.irA = irA;

// ── EXCEL ─────────────────────────────────────────────────
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

    // Detectar columnas automáticamente
    const colCodigo      = adivinarColumna(enc, ['BarNue','Codigo Activo','codigo activo','cod activo','codigo','cod','placa','etiqueta','id activo']);
    const colFamilia     = adivinarColumna(enc, PISTAS.familia);
    const colSubfamilia  = adivinarColumna(enc, PISTAS.subfamilia);
    const colDescripcion = adivinarColumna(enc, PISTAS.descripcion);

    if (!colCodigo) {
      mostrarAlerta('No encontré la columna de código del activo. Verifica que el Excel sea el correcto.');
      return;
    }

    S.inv = { registros, nombre: archivo.name.replace(/\.[^.]+$/,''), colCodigo, colFamilia, colSubfamilia, colDescripcion };

    // Mostrar info
    const zInv = $('zInv');
    zInv.classList.add('ok');
    zInv.querySelector('.ic').textContent = '✅';
    zInv.querySelector('.zt').textContent = archivo.name;
    zInv.querySelector('.zs').textContent = `${num(registros.length)} activos · encabezado en fila ${idx+1}`;
    $('inv-info').style.display = 'block';
    $('inv-resumen').innerHTML =
      `<b>Código:</b> ${colCodigo} &nbsp;·&nbsp; ` +
      `<b>Familia:</b> ${colFamilia||'no detectada'} &nbsp;·&nbsp; ` +
      `<b>Subfamilia:</b> ${colSubfamilia||'no detectada'}`;

    if (!colFamilia) mostrarAlerta('No detecté la columna de Familia. El organizador creará una sola carpeta por activo.');
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
function ocultarAlerta() {
  $('inv-alerta').style.display = 'none';
}

// ── CARPETAS ──────────────────────────────────────────────
async function elegirCarpeta(cual) {
  try {
    const handle = await pedirCarpeta();
    S[cual] = handle;
    const z = $(cual === 'origen' ? 'zOrigen' : 'zDestino');
    z.classList.add('ok');
    z.querySelector('.ic').textContent = '✅';
    z.querySelector('.zt').textContent = handle.name;
    z.querySelector('.zs').textContent = 'Seleccionada';
    verificarPaso2();
  } catch(e) {
    if (e.name !== 'AbortError') alert('No se pudo abrir la carpeta: ' + e.message);
  }
}

function verificarPaso2() {
  $('btn2').disabled = !(S.origen && S.destino);
}

// ── CONFIRMACION (paso 3) ─────────────────────────────────
function mostrarConfirmacion() {
  $('confirmGrid').innerHTML = `
    <div class="citem"><div class="cl">Inventario</div><div class="cv">${S.inv.nombre}</div></div>
    <div class="citem"><div class="cl">Activos</div><div class="cv">${num(S.inv.registros.length)}</div></div>
    <div class="citem"><div class="cl">Carpeta de fotos</div><div class="cv">${S.origen.name}</div></div>
    <div class="citem"><div class="cl">Carpeta destino</div><div class="cv">${S.destino.name}</div></div>
  `;
}

// ── PROCESO ───────────────────────────────────────────────
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

  try {
    // 1. Escanear fotos
    progTxt('Escaneando carpeta de fotos…');
    const fotos = await escanearFotos(S.origen, true, S.destino,
      n => progTxt(`Escaneando… ${num(n)} fotos`));
    log(`Fotos encontradas: ${num(fotos.length)}`);
    if (!fotos.length) throw new Error('No se encontraron fotos en la carpeta de origen.');

    // 2. Armar config desde detección automática
    const cfg = {
      colCodigo: S.inv.colCodigo,
      colEnlaceInv: null,
      colDescripcion: S.inv.colDescripcion,
      usarCatalogo: false,
      colFamilia: S.inv.colFamilia,
      colSubfamilia: S.inv.colSubfamilia,
      niveles: S.inv.colSubfamilia ? 2 : 1,
      moverHuerfanas: true,
      modo: document.querySelector('#modos .modo.sel').dataset.v,
      cliente: S.inv.nombre,
    };

    // 3. Cruzar datos
    progTxt('Procesando inventario…');
    const { activos, alertas } = cargarActivos(S.inv.registros, [], cfg);
    log(`Activos: ${num(activos.size)}`);
    S.plan = construirPlan(fotos, activos, alertas, cfg);

    // 4. Ejecutar si no es simulación
    if (!simulacion) {
      progFill(0);
      await ejecutarPlan(S.plan, S.destino, cfg.modo,
        (h, t) => { progFill(h/t); progTxt(`${cfg.modo === 'mover' ? 'Moviendo' : 'Copiando'} ${num(h)} de ${num(t)}…`); });
      log('Proceso completado.');
    }

    // 5. Reportes
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

    // 6. Mostrar resultado
    S.simulado = simulacion;
    pintarResultado(simulacion);
    irA(4);
    progFill(1);
    progTxt(simulacion ? 'Simulación lista.' : '✅ Proceso completado.');

    if (simulacion) {
      const m = calcularMetricas(S.plan);
      $('bfResumen').textContent = `${num(m.emparejadas)} fotos listas · cobertura ${m.cobertura.toFixed(0)}%`;
      $('barraFija').style.display = 'block';
      $('btnEjecutar').disabled = false;
    } else {
      $('barraFija').style.display = 'none';
    }

  } catch(e) {
    log('ERROR: ' + e.message);
    alert('Error: ' + e.message);
  } finally {
    bloquear(false);
  }
}

// ── RESULTADOS ────────────────────────────────────────────
function pintarResultado(simulacion) {
  const m = calcularMetricas(S.plan);
  const familias = agruparPorFamilia(S.plan);
  const cls = (v, b, med) => v < b ? 'bad' : v < med ? 'warn' : 'ok';

  $('tituloRes').textContent = simulacion
    ? 'Simulación — sin tocar archivos'
    : 'Proceso ejecutado';

  $('kpis').innerHTML = `
    ${kpi('Activos', num(m.totalActivos))}
    ${kpi('Fotos procesadas', num(m.totalFotos))}
    ${kpi('Cobertura', m.cobertura.toFixed(1)+'%', cls(m.cobertura,60,90), `${num(m.activosConFoto)} de ${num(m.totalActivos)}`)}
    ${kpi('Emparejadas', num(m.emparejadas), 'ok')}
    ${kpi('Sin código', num(m.huerfanas), m.huerfanas ? 'warn' : 'ok', '→ _SIN_CLASIFICAR')}
    ${kpi('Sin foto', num(m.activosSinFoto), m.activosSinFoto ? 'warn' : 'ok')}
  `;

  const avisos = [];
  if (m.sinFamilia > m.totalActivos * 0.2)
    avisos.push(`<div class="aviso warn">⚠️<div><b>${num(m.sinFamilia)} activos sin familia.</b> Revisa que la columna de familia esté correcta.</div></div>`);
  $('avisos').innerHTML = avisos.join('');

  // Gráficos
  graficoEstado($('gEstado'), m);
  graficoCobertura($('gCobertura'), familias);

  // Tabla
  const filas = [];
  for (const f of familias) {
    for (const s of f.subfamilias) {
      const c = s.cobertura;
      filas.push(`<tr>
        <td><b>${esc(f.familia)}</b></td><td>${esc(s.subfamilia)}</td>
        <td class="r">${num(s.activos)}</td><td class="r">${num(s.conFoto)}</td>
        <td class="r">${num(s.sinFoto)}</td>
        <td class="r"><span class="pill ${cls(c,60,90)}">${c.toFixed(0)}%</span></td>
      </tr>`);
    }
  }
  $('tbody').innerHTML = filas.join('');
}

const kpi = (l,v,cls='',sub='') =>
  `<div class="kpi ${cls}"><div class="kl">${l}</div><div class="kv">${v}</div>${sub?`<div class="ks">${sub}</div>`:''}</div>`;
const esc = t => String(t).replace(/[<>&]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c]));

// ── UTILIDADES ────────────────────────────────────────────
function bloquear(b) {
  $('btnSimular').disabled = b;
  $('btnEjecutar').disabled = b || !S.simulado;
  $('btnEjecutarBarra').disabled = b;
}
const progTxt  = t => { if ($('progTxt')) $('progTxt').textContent = t; };
const progFill = f => { if ($('progFill')) $('progFill').style.width = `${Math.min(100,f*100)}%`; };
function log(t) {
  const el = $('logbox');
  if (!el) return;
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

// Exponer irA globalmente para los botones onclick del HTML
export { irA };
