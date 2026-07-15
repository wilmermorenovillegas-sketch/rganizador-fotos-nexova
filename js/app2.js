/**
 * NEXOVA · Organizador de Fotos v4
 * Dashboard con 5 dimensiones, cuadros resumen visuales, progreso real
 */
import {
  celdaATexto, detectarFilaEncabezado, adivinarColumna, PISTAS,
  cargarActivos, construirPlan, calcularMetricas, normalizar
} from './motor.js';
import { soportado, pedirCarpeta, escanearFotos, ejecutarPlan, guardarEnDestino } from './archivos.js';
import { generarExcel, generarPDF } from './reportes.js';

const $ = id => document.getElementById(id);
const num = n => (+n).toLocaleString('es-PE');
const pct = n => (+n).toFixed(1) + '%';
const esc = t => String(t).replace(/[<>&]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c]));
const TEAL='#0F766E', AMBAR='#D97706', ROJO='#B91C1C';
const COLS_BAR=[TEAL,'#14B8A6','#0D9488','#2DD4BF','#0891B2','#7C3AED','#DB2777',AMBAR,'#DC2626','#65A30D','#CA8A04','#EA580C'];

const S = {
  inv:{registros:[],nombre:'',cols:{}},
  origen:null, destino:null,
  plan:null, simulado:false,
  blobExcel:null, blobPDF:null,
  charts:{}, tiempoInicio:null, datosDims:{},
  dimActiva:'familia'
};

// Dimensiones siempre presentes (con 0 si no existe la columna)
const DIMS_CONFIG = [
  { v:'familia',    l:'Familia',         pistas:['DESCRIPCION DE FAMILIA','descripcion de familia','familia','grupo','rubro'] },
  { v:'subfamilia', l:'Subfamilia',      pistas:['DESCRIPCION DE SUB FAMILIA','descripcion de sub familia','subfamilia','sub familia'] },
  { v:'area',       l:'Área',            pistas:['DescUbicacion','Desc Ubicacion','descripcion ubicacion','area','ubicacion'] },
  { v:'sede',       l:'Sede',            pistas:['DescCentroCosto','Desc Centro Costo','centro costo descripcion','sede','sucursal','centro de costo'] },
  { v:'estado',     l:'Estado',          pistas:['Estado','estado','condicion','situacion'] },
  { v:'linea',      l:'Línea producción',pistas:['linea','linea produccion','proceso','planta'] },
];

if (!soportado()) { $('app').style.display='none'; $('nocompat').style.display='grid'; }
else { iniciar(); }

function iniciar() {
  $('zInv').onclick = () => $('fInv').click();
  $('fInv').onchange = e => cargarExcel(e.target.files[0]);
  dragDrop('zInv', f => cargarExcel(f));
  $('zOrigen').onclick = () => elegirCarpeta('origen');
  $('zDestino').onclick = () => elegirCarpeta('destino');
  $('modos').onclick = e => {
    const m = e.target.closest('.modo'); if (!m) return;
    document.querySelectorAll('.modo').forEach(x => x.classList.remove('sel')); m.classList.add('sel');
  };
  $('btnSimular').onclick = () => correr(true);
  $('btnEjecutar').onclick = () => correr(false);
  $('btnEjecutarBarra').onclick = () => correr(false);
  $('btnExcel').onclick = () => descargar(S.blobExcel, `Conciliacion_${S.inv.nombre}.xlsx`);
  $('btnPDF').onclick = () => descargar(S.blobPDF, `Informe_${S.inv.nombre}.pdf`);
  $('selGrafico').onchange = () => pintarPanel(S.dimActiva);
}

function dragDrop(id, cb) {
  const z = $(id);
  z.addEventListener('dragover', e => { e.preventDefault(); z.classList.add('drag'); });
  z.addEventListener('dragleave', () => z.classList.remove('drag'));
  z.addEventListener('drop', e => { e.preventDefault(); z.classList.remove('drag'); const f=e.dataTransfer.files[0]; if(f) cb(f); });
}

function irA(paso) {
  [1,2,3,4].forEach(n => {
    $(`p${n}`).classList.toggle('visible', n===paso);
    const ws=$(`ws${n}`); ws.classList.remove('activo','listo');
    if(n===paso) ws.classList.add('activo'); else if(n<paso) ws.classList.add('listo');
    if(n<4) $(`wl${n}`).classList.toggle('ok', n<paso);
  });
  window.scrollTo({top:0,behavior:'smooth'});
}
window.irA = irA;

async function cargarExcel(archivo) {
  if (!archivo) return;
  try {
    const buffer = await archivo.arrayBuffer();
    const wb = XLSX.read(buffer, {type:'array', cellDates:true});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null, blankrows:true});
    const idx = detectarFilaEncabezado(filas);
    const vistos = {};
    const enc = (filas[idx]||[]).map((c,j) => {
      let n = celdaATexto(c)||`Col${j+1}`;
      if(vistos[n]) { vistos[n]++; n=`${n}(${vistos[n]})`; } else vistos[n]=1; return n;
    });
    const registros = [];
    for (let i=idx+1; i<filas.length; i++) {
      const f=filas[i]||[]; if(!f.some(c=>celdaATexto(c)!=='')) continue;
      const o={}; enc.forEach((h,j)=>{o[h]=f[j]??null;}); registros.push(o);
    }
    // Detectar columna de código
    const colCodigo = adivinarColumna(enc, ['BarNue','Codigo Activo','codigo activo','codigo','cod','placa','etiqueta']);
    if (!colCodigo) { alert('No encontré la columna de código del activo.'); return; }
    // Detectar todas las dimensiones (null si no existe)
    const cols = { codigo: colCodigo, descripcion: adivinarColumna(enc, ['DescCatalogo','descripcion','denominacion','detalle','bien']) };
    DIMS_CONFIG.forEach(d => { cols[d.v] = adivinarColumna(enc, d.pistas) || null; });

    S.inv = { registros, nombre: archivo.name.replace(/\.[^.]+$/,''), cols };

    const z = $('zInv');
    z.classList.add('ok');
    z.querySelector('.ic').textContent = '✅';
    z.querySelector('.zt').textContent = archivo.name;
    z.querySelector('.zs').textContent = `${num(registros.length)} activos`;
    $('inv-info').style.display = 'block';

    const detectadas = DIMS_CONFIG
      .map(d => `<span class="tag">${d.l}: <b>${cols[d.v] ? cols[d.v] : '—'}</b></span>`).join(' ');
    $('inv-resumen').innerHTML = `<span class="tag">Código: <b>${colCodigo}</b></span> ` + detectadas;
    $('btn1').disabled = false;
  } catch(e) { alert('Error leyendo el archivo: '+e.message); }
}

async function elegirCarpeta(cual) {
  try {
    const handle = await pedirCarpeta(); S[cual] = handle;
    const z = $(cual==='origen'?'zOrigen':'zDestino');
    z.classList.add('ok'); z.querySelector('.ic').textContent='✅';
    z.querySelector('.zt').textContent=handle.name; z.querySelector('.zs').textContent='Seleccionada';
    $('btn2').disabled = !(S.origen && S.destino);
  } catch(e) { if(e.name!=='AbortError') alert('No se pudo abrir: '+e.message); }
}

function mostrarConfirmacion() {
  $('confirmGrid').innerHTML = `
    <div class="citem"><div class="cl">Inventario</div><div class="cv">${esc(S.inv.nombre)}</div></div>
    <div class="citem"><div class="cl">Activos</div><div class="cv">${num(S.inv.registros.length)}</div></div>
    <div class="citem"><div class="cl">Origen</div><div class="cv">${esc(S.origen.name)}</div></div>
    <div class="citem"><div class="cl">Destino</div><div class="cv">${esc(S.destino.name)}</div></div>`;
}
window.mostrarConfirmacion = mostrarConfirmacion;

async function correr(simulacion) {
  if (!simulacion && !S.simulado) { alert('Primero corre la simulación.'); return; }
  if (!simulacion) {
    const modo = document.querySelector('#modos .modo.sel').dataset.v;
    if (!confirm(`Se van a ${modo.toUpperCase()} las fotos hacia:\n${S.destino.name}\n\n¿Continuar?`)) return;
  }
  bloquear(true);
  $('logbox').style.display = 'block';
  $('progWrap').style.display = 'block';
  $('completado').style.display = 'none';
  S.tiempoInicio = Date.now();

  const timer = setInterval(() => {
    const s = Math.floor((Date.now()-S.tiempoInicio)/1000);
    $('progTiempo').textContent = `${String(Math.floor(s/60)).padStart(2,'0')}:${String(s%60).padStart(2,'0')}`;
  }, 500);

  try {
    // Verificaciones tempranas con logs visibles
    log('Iniciando ' + (simulacion ? 'simulación' : 'ejecución') + '...');
    if (!S.origen) throw new Error('No se seleccionó carpeta de fotos.');
    if (!S.destino) throw new Error('No se seleccionó carpeta destino.');
    log('Carpeta origen: ' + S.origen.name);
    log('Carpeta destino: ' + S.destino.name);

    setProg(0.05, 'Escaneando carpeta de fotos…');
    const fotos = await escanearFotos(S.origen, true, S.destino,
      n => setProg(0.05+Math.min(n/5000,1)*0.1, `Escaneando… ${num(n)} fotos encontradas`));
    log(`Fotos encontradas: ${num(fotos.length)}`);
    if (!fotos.length) throw new Error('No se encontraron fotos en la carpeta de origen.');

    setProg(0.18, 'Procesando inventario…');
    const cols = S.inv.cols;
    const cfg = {
      colCodigo: cols.codigo, colEnlaceInv: null,
      colDescripcion: cols.descripcion, usarCatalogo: false,
      colFamilia: cols.familia, colSubfamilia: cols.subfamilia,
      colArea: cols.area, colSede: cols.sede,
      colEstado: cols.estado, colLinea: cols.linea,
      niveles: cols.subfamilia ? 2 : 1,
      moverHuerfanas: true,
      modo: document.querySelector('#modos .modo.sel').dataset.v,
      cliente: S.inv.nombre,
    };
    const { activos, alertas } = cargarActivos(S.inv.registros, [], cfg);
    log(`Activos cargados: ${num(activos.size)}`);
    S.plan = construirPlan(fotos, activos, alertas, cfg);
    setProg(0.25, 'Plan construido…');

    if (!simulacion) {
      let prev = 0;
      await ejecutarPlan(S.plan, S.destino, cfg.modo, (h, t) => {
        setProg(0.25+(h/t)*0.65, `${cfg.modo==='mover'?'Moviendo':'Copiando'} ${num(h)} de ${num(t)} fotos…`);
        if (h-prev >= 100) { log(`${num(h)} / ${num(t)} fotos procesadas`); prev=h; }
      });
      log('Archivos procesados. Generando reportes…');
    }

    setProg(0.92, 'Generando reportes…');
    S.blobExcel = await generarExcel(S.plan, cfg, simulacion);
    S.blobPDF   = await generarPDF(S.plan, cfg, {}, simulacion);
    if (!simulacion) {
      await guardarEnDestino(S.destino, 'Conciliacion.xlsx', S.blobExcel);
      await guardarEnDestino(S.destino, 'Informe_Cobertura.pdf', S.blobPDF);
      log('Reportes guardados en la carpeta destino.');
    }
    setProg(1, simulacion ? 'Simulación completada.' : '¡Proceso completado!');
    S.simulado = simulacion;

    const m = calcularMetricas(S.plan);
    const seg = Math.floor((Date.now()-S.tiempoInicio)/1000);
    const tStr = seg>=60 ? `${Math.floor(seg/60)} min ${seg%60} seg` : `${seg} seg`;

    if (!simulacion) {
      $('completado').style.display = 'flex';
      $('completado-msg').innerHTML = `✅ <b>${num(m.emparejadas)} fotos organizadas</b> en ${tStr}`;
      sonido();
    }

    pintarDashboard(cfg, m);
    irA(4);

    if (simulacion) {
      $('bfResumen').textContent = `${num(m.emparejadas)} listas · ${pct(m.cobertura)} cobertura`;
      $('barraFija').style.display = 'block';
      $('btnEjecutar').disabled = false;
    } else { $('barraFija').style.display = 'none'; }

  } catch(e) { log('ERROR: '+e.message); alert('Error:\n'+e.message); setProg(0,'Error.'); }
  finally { clearInterval(timer); bloquear(false); }
}

// ── Dashboard ──────────────────────────────────────────────
function pintarDashboard(cfg, m) {
  const cls = (v,b,med) => v<b?'bad':v<med?'warn':'ok';
  $('kpis').innerHTML = [
    kpi('Total activos',    num(m.totalActivos)),
    kpi('Fotos procesadas', num(m.totalFotos)),
    kpi('Cobertura',        pct(m.cobertura), cls(m.cobertura,60,90), `${num(m.activosConFoto)} con foto`),
    kpi('Emparejadas',      num(m.emparejadas), 'ok'),
    kpi('Sin código',       num(m.huerfanas),   m.huerfanas?'warn':'ok'),
    kpi('Sin foto',         num(m.activosSinFoto), m.activosSinFoto?'warn':'ok'),
  ].join('');

  // Calcular datos para TODAS las dimensiones (vacío si no hay columna)
  const dimACol = {
    familia: cfg.colFamilia, subfamilia: cfg.colSubfamilia,
    area: cfg.colArea, sede: cfg.colSede,
    estado: cfg.colEstado, linea: cfg.colLinea,
  };
  S.datosDims = {};
  DIMS_CONFIG.forEach(d => {
    const col = dimACol[d.v] || null;
    S.datosDims[d.v] = col ? calcDim(col) : [];
  });

  // Tabs — siempre mostrar todas las dimensiones (con badge si tiene datos)
  $('dimTabs').innerHTML = DIMS_CONFIG.map((d,i) => {
    const n = S.datosDims[d.v].length;
    const badge = n ? `<span class="badge">${n}</span>` : '<span class="badge empty">0</span>';
    return `<button class="dtab${i===0?' activo':''}" data-dim="${d.v}" onclick="cambiarDim('${d.v}',this)">${d.l}${badge}</button>`;
  }).join('');

  pintarPanel('familia');
}

function calcDim(col) {
  const mapa = new Map();
  for (const reg of S.inv.registros) {
    const cod = normalizar(celdaATexto(reg[S.inv.cols.codigo])); if (!cod) continue;
    const val = celdaATexto(reg[col]) || 'SIN DATO';
    if (!mapa.has(val)) mapa.set(val, {nombre:val, activos:0, conFoto:0});
    const g = mapa.get(val); g.activos++;
    if (S.plan.codigosConFoto.has(cod)) g.conFoto++;
  }
  return [...mapa.values()]
    .map(g => ({...g, sinFoto:g.activos-g.conFoto, cobertura:g.activos?(g.conFoto/g.activos)*100:0}))
    .sort((a,b) => b.activos-a.activos);
}

function pintarPanel(dim) {
  S.dimActiva = dim;
  const datos = S.datosDims[dim] || [];
  const tipo = $('selGrafico').value || 'barrasH';
  const label = DIMS_CONFIG.find(d=>d.v===dim)?.l || dim;
  $('thDim').textContent = label;

  if (datos.length === 0) {
    dest('gPrincipal'); dest('gDona');
    $('gPrincipal').getContext('2d').clearRect(0,0,9999,9999);
    $('tbody').innerHTML = `<tr><td colspan="5" style="text-align:center;color:var(--gris);padding:20px">Esta dimensión no existe en el inventario de este cliente</td></tr>`;
    $('resumenCards').innerHTML = '';
    return;
  }

  pintarGrafico('gPrincipal', datos, tipo);
  pintarDona('gDona');
  pintarTabla(datos);
  pintarCards(datos, label);
}

function cambiarDim(dim, btn) {
  document.querySelectorAll('.dtab').forEach(b => b.classList.remove('activo'));
  btn.classList.add('activo');
  pintarPanel(dim);
}
window.cambiarDim = cambiarDim;

// ── Gráficos ───────────────────────────────────────────────
const fb = { id:'fb', beforeDraw(c){ const ctx=c.ctx; ctx.save(); ctx.globalCompositeOperation='destination-over'; ctx.fillStyle='#fff'; ctx.fillRect(0,0,c.width,c.height); ctx.restore(); }};
function dest(id) { if(S.charts[id]) { S.charts[id].destroy(); delete S.charts[id]; } }

function pintarGrafico(id, datos, tipo) {
  dest(id); const top=datos.slice(0,14); const canvas=$(id); if(!canvas) return;
  if (tipo==='dona') {
    S.charts[id] = new Chart(canvas.getContext('2d'), {
      type:'doughnut', plugins:[fb],
      data:{ labels:top.map(d=>d.nombre), datasets:[{ data:top.map(d=>d.activos), backgroundColor:COLS_BAR.slice(0,top.length), borderWidth:2, borderColor:'#fff' }] },
      options:{ responsive:true, maintainAspectRatio:false, cutout:'55%', plugins:{ legend:{ position:'right', labels:{ font:{size:11}, boxWidth:12 } } } }
    }); return;
  }
  const h = tipo==='barrasH';
  S.charts[id] = new Chart(canvas.getContext('2d'), {
    type:'bar', plugins:[fb],
    data:{ labels:top.map(d=>d.nombre), datasets:[
      { label:'Con foto', data:top.map(d=>d.conFoto), backgroundColor:TEAL, borderRadius:3, barThickness:16 },
      { label:'Sin foto', data:top.map(d=>d.sinFoto), backgroundColor:'#CBD5E1', borderRadius:3, barThickness:16 },
    ]},
    options:{ responsive:true, maintainAspectRatio:false, indexAxis:h?'y':'x',
      scales:{ x:{stacked:true,ticks:{font:{size:10},color:'#475569'},grid:{color:'#F1F5F9'}}, y:{stacked:true,ticks:{font:{size:10},color:'#475569'},grid:{color:h?'#F1F5F9':'transparent'}} },
      plugins:{ legend:{ position:'bottom', labels:{ font:{size:11}, boxWidth:12 } } }
    }
  });
}

function pintarDona(id) {
  dest(id); const m=calcularMetricas(S.plan); const canvas=$(id); if(!canvas) return;
  S.charts[id] = new Chart(canvas.getContext('2d'), {
    type:'doughnut', plugins:[fb],
    data:{ labels:['Con foto','Sin foto'], datasets:[{ data:[m.activosConFoto,m.activosSinFoto], backgroundColor:[TEAL,'#CBD5E1'], borderWidth:2, borderColor:'#fff' }] },
    options:{ responsive:true, maintainAspectRatio:false, cutout:'65%', plugins:{ legend:{ position:'bottom', labels:{ font:{size:11}, boxWidth:12 } } } }
  });
}

// ── Tabla ─────────────────────────────────────────────────
function pintarTabla(datos) {
  const cls=(v,b,med)=>v<b?'bad':v<med?'warn':'ok';
  $('tbody').innerHTML = datos.map(d=>`<tr>
    <td>${esc(d.nombre)}</td>
    <td class="r">${num(d.activos)}</td>
    <td class="r" style="color:var(--teal);font-weight:700">${num(d.conFoto)}</td>
    <td class="r" style="color:var(--ambar)">${num(d.sinFoto)}</td>
    <td class="r"><span class="pill ${cls(d.cobertura,60,90)}">${pct(d.cobertura)}</span></td>
  </tr>`).join('');
}

// ── Cuadros resumen visuales ───────────────────────────────
function pintarCards(datos, label) {
  const cls=(v,b,med)=>v<b?'bad':v<med?'warn':'ok';
  const top = datos.slice(0,12);
  $('resumenCards').innerHTML = `
    <div style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:.04em;margin-bottom:10px">
      Cuadros resumen — ${label}
    </div>
    <div class="cards-grid">
      ${top.map(d => {
        const c = d.cobertura;
        const cl = cls(c,60,90);
        const barW = Math.round(c);
        return `<div class="rcard">
          <div class="rcard-nombre">${esc(d.nombre)}</div>
          <div class="rcard-nums">
            <span class="rcard-total">${num(d.activos)} activos</span>
            <span class="rcard-pct ${cl}">${pct(c)}</span>
          </div>
          <div class="rcard-bar-bg"><div class="rcard-bar-fill ${cl}" style="width:${barW}%"></div></div>
          <div class="rcard-sub">
            <span style="color:var(--teal)">✓ ${num(d.conFoto)} con foto</span>
            <span style="color:var(--ambar)">✗ ${num(d.sinFoto)} sin foto</span>
          </div>
        </div>`;
      }).join('')}
    </div>`;
}

// ── Descarga resumen Excel ─────────────────────────────────
async function descargarResumen() {
  const dim = S.dimActiva;
  const label = DIMS_CONFIG.find(d=>d.v===dim)?.l || dim;
  const datos = S.datosDims[dim] || [];
  if (!datos.length) { alert('Esta dimensión no tiene datos para exportar.'); return; }
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(label);
  ws.columns=[{header:label,key:'n',width:40},{header:'Activos',key:'a',width:12},{header:'Con foto',key:'c',width:12},{header:'Sin foto',key:'s',width:12},{header:'Cobertura %',key:'p',width:14}];
  ws.getRow(1).eachCell(c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1E293B'}};c.font={bold:true,color:{argb:'FFFFFFFF'}};});
  datos.forEach(d=>ws.addRow({n:d.nombre,a:d.activos,c:d.conFoto,s:d.sinFoto,p:+d.cobertura.toFixed(1)}));
  const buf = await wb.xlsx.writeBuffer();
  descargar(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), `Resumen_${label}.xlsx`);
}
window.descargarResumen = descargarResumen;

// ── Utilidades ─────────────────────────────────────────────
const kpi = (l,v,cls='',sub='') =>
  `<div class="kpi ${cls}"><div class="kl">${l}</div><div class="kv">${v}</div>${sub?`<div class="ks">${sub}</div>`:''}</div>`;

function sonido(){try{const ctx=new AudioContext(),o=ctx.createOscillator(),g=ctx.createGain();o.connect(g);g.connect(ctx.destination);o.frequency.setValueAtTime(880,ctx.currentTime);o.frequency.setValueAtTime(1100,ctx.currentTime+0.15);g.gain.setValueAtTime(0.3,ctx.currentTime);g.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+0.5);o.start();o.stop(ctx.currentTime+0.5);}catch{}}
function bloquear(b){$('btnSimular').disabled=b;$('btnEjecutar').disabled=b||!S.simulado;$('btnEjecutarBarra').disabled=b;}
function setProg(f,txt){if($('progFill'))$('progFill').style.width=`${Math.min(100,f*100).toFixed(0)}%`;if($('progTxt'))$('progTxt').textContent=txt;}
function log(t){const el=$('logbox');if(!el)return;el.textContent+=new Date().toLocaleTimeString('es-PE')+'  '+t+'\n';el.scrollTop=el.scrollHeight;}
function descargar(blob,nombre){if(!blob)return;const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=nombre;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
