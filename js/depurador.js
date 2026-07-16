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
 * · PDF vectorial (jsPDF + autoTable + gráficos en canvas ocultos), no captura.
 * · Reutiliza los helpers ya probados de motor.js (no se modifica motor.js).
 * ============================================================================
 */
import { celdaATexto, normalizar, detectarFilaEncabezado, adivinarColumna } from './motor.js';

const $ = (id) => document.getElementById(id);
const num = (n) => (+n || 0).toLocaleString('es-PE');
const pct1 = (n) => (+n || 0).toFixed(1) + '%';
const esc = (t) => String(t ?? '').replace(/[<>&]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c]));

const TEAL = '#0F766E';
const PALETA = ['#0F766E', '#14B8A6', '#0D9488', '#2DD4BF', '#0891B2', '#0EA5E9',
                '#6366F1', '#7C3AED', '#DB2777', '#F59E0B', '#DC2626', '#65A30D'];

/** Campos canónicos y sus pistas de detección (genéricos, no atados a SITIA). */
const CAMPOS = [
  { k: 'codigo',      l: 'Código',          key: true, hints: ['codigo inventario', 'codigo activo', 'codigo interno', 'codigo', 'cod', 'barnue', 'placa', 'etiqueta', 'id activo', 'correlativo'] },
  { k: 'descripcion', l: 'Descripción',     hints: ['descripcion del activo', 'descripcion', 'denominacion', 'detalle', 'bien', 'nombre', 'articulo', 'desccatalogo'] },
  { k: 'marca',       l: 'Marca',           hints: ['marca', 'fabricante', 'brand'] },
  { k: 'modelo',      l: 'Modelo',          hints: ['modelo', 'model'] },
  { k: 'serie',       l: 'Nro. Serie',      hints: ['nro serie', 'numero de serie', 'n serie', 'serie', 'serial'] },
  { k: 'ubicacion',   l: 'Ubicación',       hints: ['descripcion ubicacion', 'descubicacion', 'ubicacion', 'area', 'ambiente'] },
  { k: 'centro',      l: 'Centro de costo', hints: ['centro de costo', 'centro costo', 'desccentrocosto', 'cco'] },
  { k: 'responsable', l: 'Responsable',     hints: ['responsable', 'custodio', 'usuario', 'asignado'] },
  { k: 'linea',       l: 'Línea',           hints: ['linea de produccion', 'linea produccion', 'linea', 'proceso', 'planta'] },
  { k: 'familia',     l: 'Familia',         hints: ['descripcion de familia', 'familia', 'grupo', 'rubro', 'clase', 'categoria'] },
  { k: 'estado',      l: 'Estado',          hints: ['estado conservacion', 'estado de conservacion', 'estado', 'condicion', 'situacion'] },
  { k: 'color',       l: 'Color',           hints: ['color'] },
  { k: 'observacion', l: 'Observaciones',   hints: ['observaciones', 'observacion', 'nota', 'glosa'] },
];
// Dimensiones que alimentan los gráficos del resumen (si están detectadas).
const DIMS_RESUMEN = ['familia', 'estado', 'ubicacion', 'centro', 'marca', 'linea'];
// Campos cuyo texto se pasa a MAYÚSCULAS al depurar.
const A_MAYUS = new Set(['marca', 'modelo', 'serie', 'color']);

const D = {
  raw: [], headers: [], cols: {}, clean: [], nombre: '',
  maestro: null, alertas: [], M: {}, tab: 'resumen',
};
const CH = {};
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

    // Detección automática por pistas.
    D.cols = {};
    CAMPOS.forEach((c) => { D.cols[c.k] = adivinarColumna(D.headers, c.hints) || ''; });

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

function tituloCaso(s) {
  return String(s).toLowerCase().replace(/(^|\s|\(|\/|-)\S/g, (a) => a.toUpperCase());
}

function procesar() {
  D.clean = D.raw.map((row) => {
    const r = { _ch: 0 };
    CAMPOS.forEach((c) => {
      let v = gv(row, c.k);
      if (typeof v === 'string') {
        const o = v;
        v = v.trim().replace(/\s+/g, ' ');
        if (c.k === 'descripcion') v = tituloCaso(v);
        else if (A_MAYUS.has(c.k)) v = v.toUpperCase();
        if (v !== o) r._ch++;
      }
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
  for (const r of D.clean) {
    // Enriquecer a la grafía canónica del maestro cuando coincide.
    if (D.maestro) {
      const m = D.maestro.get(normalizar(r.marca));
      if (m) {
        r.marca = m.nombre;
        const mo = m.modelos.get(normalizar(r.modelo));
        if (mo) r.modelo = mo;
      }
    }
    const c = clasificarMM(r);
    r._mm = c ? c.tipo : 'ok';
    if (c) D.alertas.push({ codigo: r.codigo, descripcion: r.descripcion, marca: r.marca, modelo: r.modelo, ubicacion: r.ubicacion, motivo: c.motivo });
  }
}

function calcularMetricas() {
  const total = D.clean.length;
  const presentes = CAMPOS.filter((c) => D.cols[c.k]);
  // Completitud por campo detectado.
  const compl = presentes.map((c) => {
    const llenos = D.clean.filter((r) => String(r[c.k] ?? '').trim() !== '').length;
    return { k: c.k, l: c.l, llenos, pct: total ? (llenos / total) * 100 : 0 };
  });
  const complGlobal = compl.length ? compl.reduce((s, x) => s + x.pct, 0) / compl.length : 0;
  // Duplicados.
  const dup = (k) => {
    if (!D.cols[k]) return { n: 0, lista: [] };
    const vistos = new Map(); const rep = new Map();
    D.clean.forEach((r) => {
      const v = normalizar(r[k]); if (!v) return;
      if (vistos.has(v)) rep.set(v, (rep.get(v) || 1) + 1); else vistos.set(v, 1);
    });
    return { n: rep.size, lista: [...rep.entries()].map(([v, c]) => ({ valor: v, veces: c })).sort((a, b) => b.veces - a.veces) };
  };
  const sinCodigo = D.clean.filter((r) => !normalizar(r.codigo)).length;
  const alertasVacio = D.alertas.filter((a) => a.motivo.startsWith('Sin')).length;
  const alertasNoReg = D.alertas.length - alertasVacio;

  D.M = {
    total, presentes, compl, complGlobal,
    dupCodigo: dup('codigo'), dupSerie: dup('serie'),
    sinCodigo, cambios: D.clean.reduce((s, r) => s + (r._ch || 0), 0),
    alertasVacio, alertasNoReg, mmOk: D.clean.filter((r) => r._mm === 'ok').length,
  };
}

// ── 3) Dashboard ────────────────────────────────────────────
const TABS = [
  { k: 'resumen',  l: 'Resumen' },
  { k: 'calidad',  l: 'Calidad de datos' },
  { k: 'mm',       l: 'Marca / Modelo' },
  { k: 'base',     l: 'Base depurada' },
];

function pintarTabs() {
  $('depTabs').innerHTML = TABS.map((t) =>
    `<button class="dtab${t.k === D.tab ? ' activo' : ''}" data-t="${t.k}">${t.l}</button>`).join('');
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
  if (t === 'resumen') panelResumen();
  else if (t === 'calidad') panelCalidad();
  else if (t === 'mm') panelMM();
  else if (t === 'base') panelBase();
}

function panelResumen() {
  const dims = DIMS_RESUMEN.filter((k) => D.cols[k]);
  if (!dims.length) {
    $('depPanel').innerHTML = `<div class="aviso warn"><span>ℹ️</span><div>No se detectaron dimensiones (familia, estado, ubicación, etc.) para graficar. Revisa el mapeo de columnas.</div></div>`;
    return;
  }
  const label = (k) => CAMPOS.find((c) => c.k === k).l;
  $('depPanel').innerHTML = `<div class="graficos-grid" style="grid-template-columns:1fr 1fr">
    ${dims.map((k) => `<div class="gpanel"><div class="gtitle">Registros por ${esc(label(k).toLowerCase())}</div>
      <div class="gcanvas-wrap"><canvas id="dc_${k}" width="470" height="260"></canvas></div></div>`).join('')}
  </div>`;
  dims.forEach((k) => barra(`dc_${k}`, conteo(k).slice(0, 12)));
}

function panelCalidad() {
  const M = D.M;
  const bar = (c) => {
    const cl = semClass(c.pct);
    return `<div class="q-row">
      <div class="qn">${esc(c.l)} <span style="color:var(--gris);font-weight:400">· ${num(c.llenos)}/${num(M.total)}</span></div>
      <div class="q-bar"><div class="q-bar-fill ${cl}" style="width:${c.pct.toFixed(0)}%"></div></div>
      <div class="q-pct ${cl}">${pct1(c.pct)}</div>
    </div>`;
  };
  const dupBlock = (titulo, d) => `<div class="cx" style="background:var(--blanco);border:1px solid var(--gris-200);border-radius:10px;padding:14px">
      <div class="gtitle">${titulo}: ${num(d.n)} valores repetidos</div>
      ${d.n ? `<div class="tabla-wrap" style="margin:0"><table><thead><tr><th>Valor</th><th class="r">Veces</th></tr></thead><tbody>
        ${d.lista.slice(0, 10).map((x) => `<tr><td>${esc(x.valor)}</td><td class="r">${num(x.veces)}</td></tr>`).join('')}
      </tbody></table></div>` : '<div class="mini-note">Sin duplicados.</div>'}
    </div>`;
  $('depPanel').innerHTML = `
    <div class="res-tit" style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:.04em;margin-bottom:10px">Completitud por campo</div>
    <div class="q-list" style="margin-bottom:16px">${M.compl.map(bar).join('')}</div>
    <div class="graficos-grid" style="grid-template-columns:1fr 1fr">
      ${dupBlock('Códigos duplicados', M.dupCodigo)}
      ${dupBlock('Series duplicadas', M.dupSerie)}
    </div>
    ${M.sinCodigo ? `<div class="aviso warn" style="margin-top:12px"><span>⚠️</span><div><b>${num(M.sinCodigo)} registros sin código.</b> No podrán identificarse de forma única.</div></div>` : ''}`;
}

function panelMM() {
  const M = D.M;
  const estado = D.maestro
    ? `<b>Maestro cargado</b>${num(D.maestro.size)} marcas · valida contra el catálogo`
    : `<b>Sin maestro cargado</b>Solo se listan los activos con marca/modelo vacío. Carga un maestro para validar contra el catálogo.`;
  const topNoReg = conteoAlertas();
  $('depPanel').innerHTML = `
    <div class="maestro-box">
      <div class="mtxt">${estado}</div>
      <input type="file" id="depMaestro" accept=".xlsx,.xls,.csv,.json" hidden>
      <button class="b-desc teal" id="btnMaestro">⬆ Cargar maestro (Excel/JSON)</button>
      ${D.maestro ? '<button class="b-desc" id="btnMaestroClr">Quitar</button>' : ''}
    </div>
    <div class="kpis" style="grid-template-columns:repeat(3,1fr);margin-bottom:14px">
      ${kpi('Marca/modelo OK', num(M.mmOk), 'ok')}
      ${kpi('Con campo vacío', num(M.alertasVacio), M.alertasVacio ? 'warn' : 'ok')}
      ${kpi('No registrados', num(M.alertasNoReg), M.alertasNoReg ? 'warn' : 'ok')}
    </div>
    ${topNoReg.length ? `<div class="res-tit" style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px">Marcas más frecuentes en las alertas (candidatas al maestro)</div>
      <div class="cards-grid" style="margin-bottom:14px">${topNoReg.slice(0, 8).map((x) => `<div class="rcard"><div class="rcard-nombre">${esc(x.valor)}</div><div class="rcard-nums"><span class="rcard-total">${num(x.veces)} activos</span></div></div>`).join('')}</div>` : ''}
    <div class="tabla-wrap">
      <div style="padding:9px 13px;border-bottom:1px solid var(--gris-200);display:flex;align-items:center;gap:10px;flex-wrap:wrap">
        <input class="dep-search" id="mmSearch" placeholder="🔍 Buscar en alertas…">
        <span class="mini-note" style="margin:0" id="mmCnt"></span>
      </div>
      <div style="max-height:460px;overflow:auto"><table><thead><tr>
        <th>Código</th><th>Descripción</th><th>Marca</th><th>Modelo</th><th>Ubicación</th><th>Motivo</th>
      </tr></thead><tbody id="mmBody"></tbody></table></div>
    </div>
    <div class="mini-note">${num(D.alertas.length)} activos sin marca/modelo registrado.</div>`;

  $('btnMaestro').onclick = () => $('depMaestro').click();
  $('depMaestro').onchange = (e) => cargarMaestro(e.target.files[0]);
  if (D.maestro) $('btnMaestroClr').onclick = () => { D.maestro = null; validarMM(); calcularMetricas(); pintarTabs(); irTab('mm'); };
  const search = $('mmSearch');
  const render = () => {
    const q = search.value.toLowerCase();
    const filas = D.alertas.filter((a) => !q || [a.codigo, a.descripcion, a.marca, a.modelo, a.motivo].some((v) => String(v ?? '').toLowerCase().includes(q)));
    $('mmBody').innerHTML = filas.slice(0, 800).map((a) => `<tr>
      <td>${esc(a.codigo)}</td><td>${esc(a.descripcion)}</td><td>${esc(a.marca)}</td>
      <td>${esc(a.modelo)}</td><td>${esc(a.ubicacion)}</td>
      <td><span class="pill ${a.motivo.startsWith('Sin') ? 'warn' : 'bad'}">${esc(a.motivo)}</span></td></tr>`).join('')
      || `<tr><td colspan="6" style="text-align:center;color:var(--gris);padding:18px">Sin coincidencias.</td></tr>`;
    $('mmCnt').textContent = `${num(Math.min(filas.length, 800))} de ${num(filas.length)}`;
  };
  search.oninput = render;
  render();
}

function panelBase() {
  const cols = CAMPOS.filter((c) => D.cols[c.k]);
  $('depPanel').innerHTML = `
    <div class="tabla-wrap">
      <div style="padding:9px 13px;border-bottom:1px solid var(--gris-200);display:flex;align-items:center;gap:10px;flex-wrap:wrap">
        <input class="dep-search" id="baseSearch" placeholder="🔍 Buscar en la base…">
        <span class="mini-note" style="margin:0" id="baseCnt"></span>
      </div>
      <div style="max-height:560px;overflow:auto"><table><thead><tr>
        ${cols.map((c) => `<th>${esc(c.l)}</th>`).join('')}
      </tr></thead><tbody id="baseBody"></tbody></table></div>
    </div>
    <div class="mini-note">Se muestran hasta 800 filas; usa la búsqueda o exporta el Excel para la base completa.</div>`;
  const search = $('baseSearch');
  const render = () => {
    const q = search.value.toLowerCase();
    const filas = D.clean.filter((r) => !q || cols.some((c) => String(r[c.k] ?? '').toLowerCase().includes(q)));
    $('baseBody').innerHTML = filas.slice(0, 800).map((r) =>
      `<tr>${cols.map((c) => `<td>${esc(r[c.k])}</td>`).join('')}</tr>`).join('')
      || `<tr><td colspan="${cols.length}" style="text-align:center;color:var(--gris);padding:18px">Sin coincidencias.</td></tr>`;
    $('baseCnt').textContent = `${num(Math.min(filas.length, 800))} de ${num(filas.length)}`;
  };
  search.oninput = render;
  render();
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
function conteoAlertas() {
  const m = new Map();
  D.alertas.forEach((a) => { const v = String(a.marca ?? '').trim(); if (!v) return; m.set(v, (m.get(v) || 0) + 1); });
  return [...m.entries()].map(([valor, veces]) => ({ valor, veces })).sort((a, b) => b.veces - a.veces);
}
const semClass = (v) => (v < 60 ? 'bad' : v < 90 ? 'warn' : 'ok');
const kpi = (l, v, cls = '', sub = '') => `<div class="kpi ${cls}"><div class="kl">${l}</div><div class="kv">${v}</div>${sub ? `<div class="ks">${sub}</div>` : ''}</div>`;

// ── Gráficos (responsive:false, tamaño fijo → sin scroll infinito) ─
const fondoBlanco = { id: 'fb', beforeDraw(c) { const x = c.ctx; x.save(); x.globalCompositeOperation = 'destination-over'; x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.restore(); } };
function destruirCharts() { for (const k in CH) { CH[k]?.destroy?.(); delete CH[k]; } }
function barra(id, datos) {
  const cv = $(id); if (!cv) return;
  CH[id] = new Chart(cv.getContext('2d'), {
    type: 'bar', plugins: [fondoBlanco],
    data: { labels: datos.map((d) => d.valor), datasets: [{ data: datos.map((d) => d.veces), backgroundColor: datos.map((_, i) => PALETA[i % PALETA.length]), borderRadius: 4, barThickness: 14 }] },
    options: {
      responsive: false, maintainAspectRatio: false, animation: false, indexAxis: 'y',
      scales: { x: { beginAtZero: true, ticks: { font: { size: 10 }, color: '#475569' }, grid: { color: '#F1F5F9' } }, y: { ticks: { font: { size: 10 }, color: '#475569', autoSkip: false }, grid: { display: false } } },
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => ` ${num(c.raw)} registros` } } },
    },
  });
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
  const cols = CAMPOS.filter((c) => D.cols[c.k]);
  const wb = new ExcelJS.Workbook(); wb.creator = 'NEXOVA Suite';
  const ws = wb.addWorksheet('BASE DEPURADA');
  ws.columns = cols.map((c) => ({ header: c.l, key: c.k, width: Math.max(12, Math.min(40, c.l.length + 8)) }));
  styleHead(ws.getRow(1));
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
  D.clean.forEach((r) => { const o = {}; cols.forEach((c) => { o[c.k] = r[c.k]; }); ws.addRow(o); });
  bajar(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Base_Depurada_${D.nombre}.xlsx`);
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
  wc.columns = [{ header: 'Campo', key: 'l', width: 26 }, { header: 'Con dato', key: 'n', width: 12 }, { header: 'Total', key: 't', width: 12 }, { header: '% completitud', key: 'p', width: 14 }];
  styleHead(wc.getRow(1));
  M.compl.forEach((c) => wc.addRow({ l: c.l, n: c.llenos, t: M.total, p: +c.pct.toFixed(1) }));

  const wa = wb.addWorksheet('ALERTAS MARCA-MODELO');
  wa.columns = [{ header: 'Código', key: 'c', width: 20 }, { header: 'Descripción', key: 'd', width: 40 }, { header: 'Marca', key: 'm', width: 18 }, { header: 'Modelo', key: 'mo', width: 18 }, { header: 'Ubicación', key: 'u', width: 24 }, { header: 'Motivo', key: 'mt', width: 34 }];
  styleHead(wa.getRow(1));
  D.alertas.forEach((a) => wa.addRow({ c: a.codigo, d: a.descripcion, m: a.marca, mo: a.modelo, u: a.ubicacion, mt: a.motivo }));

  bajar(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Calidad_${D.nombre}.xlsx`);
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

  // Completitud por campo
  doc.setFillColor(15, 118, 110); doc.rect(M, y - 4, 3, 6, 'F');
  doc.setFont('helvetica', 'bold').setFontSize(12).setTextColor(30, 41, 59); doc.text('Completitud por campo', M + 6, y + 1);
  y += 6;
  doc.autoTable({
    startY: y, margin: { left: M, right: M },
    head: [['Campo', 'Con dato', 'Total', '% completitud']],
    body: m.compl.map((c) => [c.l, num(c.llenos), num(m.total), `${c.pct.toFixed(1)}%`]),
    theme: 'grid', styles: { fontSize: 8, cellPadding: 1.8 }, headStyles: { fillColor: [30, 41, 59], textColor: 255, fontSize: 8 },
    alternateRowStyles: { fillColor: [248, 250, 251] }, columnStyles: { 1: { halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' } },
  });
  y = doc.lastAutoTable.finalY + 10;

  // Gráficos de dimensiones (canvas ocultos)
  const cont = document.createElement('div'); cont.style.cssText = 'position:fixed;left:-99999px;top:0;pointer-events:none'; document.body.appendChild(cont);
  try {
    const dims = DIMS_RESUMEN.filter((k) => D.cols[k]).slice(0, 4);
    for (const k of dims) {
      if (y > 235) { doc.addPage(); y = 20; }
      const label = CAMPOS.find((c) => c.k === k).l;
      doc.setFillColor(15, 118, 110); doc.rect(M, y - 4, 3, 6, 'F');
      doc.setFont('helvetica', 'bold').setFontSize(12).setTextColor(30, 41, 59); doc.text(`Registros por ${label.toLowerCase()}`, M + 6, y + 1);
      y += 8;
      const datos = conteo(k).slice(0, 12);
      const cv = document.createElement('canvas'); cv.width = 900; cv.height = 420; cont.appendChild(cv);
      const ch = new Chart(cv.getContext('2d'), {
        type: 'bar', plugins: [fondoBlanco],
        data: { labels: datos.map((d) => d.valor), datasets: [{ data: datos.map((d) => d.veces), backgroundColor: datos.map((_, i) => PALETA[i % PALETA.length]), borderRadius: 4 }] },
        options: { responsive: false, maintainAspectRatio: false, animation: false, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true, ticks: { color: '#475569' } }, y: { ticks: { color: '#475569', font: { size: 11 } } } } },
      });
      await new Promise((r) => requestAnimationFrame(r));
      doc.addImage(cv.toDataURL('image/png', 1), 'PNG', M, y, W - M * 2, 52); ch.destroy();
      y += 58;
    }

    if (D.alertas.length) {
      if (y > 220) { doc.addPage(); y = 20; }
      doc.setFillColor(193, 30, 58); doc.rect(M, y - 4, 3, 6, 'F');
      doc.setFont('helvetica', 'bold').setFontSize(12).setTextColor(30, 41, 59); doc.text('Activos sin marca/modelo registrado', M + 6, y + 1);
      y += 6;
      doc.autoTable({
        startY: y, margin: { left: M, right: M },
        head: [['Código', 'Descripción', 'Marca', 'Modelo', 'Motivo']],
        body: D.alertas.slice(0, 300).map((a) => [a.codigo, a.descripcion, a.marca, a.modelo, a.motivo]),
        theme: 'striped', styles: { fontSize: 7.5, cellPadding: 1.6 }, headStyles: { fillColor: [30, 41, 59], textColor: 255, fontSize: 8 },
        columnStyles: { 0: { cellWidth: 26, fontStyle: 'bold' } },
      });
      if (D.alertas.length > 300) {
        doc.setFont('helvetica', 'italic').setFontSize(8).setTextColor(148, 163, 184);
        doc.text(`Se muestran 300 de ${num(D.alertas.length)}. El listado completo está en el Excel de calidad.`, M, doc.lastAutoTable.finalY + 6);
      }
    }
    bajar(doc.output('blob'), `Informe_Calidad_${D.nombre}.pdf`);
  } finally {
    cont.remove();
  }
}
