/**
 * ============================================================================
 * catalogo.js — Catálogo fotográfico de inventario en PowerPoint (.pptx)
 * ----------------------------------------------------------------------------
 * Estructura: portada con mosaico de fotos · contenido · separador por familia
 * · láminas con 2 activos cada una (foto principal + 2 miniaturas + ficha).
 *
 * Fotos SIN deformar: cada foto se compone en un canvas con la proporción
 * exacta de su marco — la imagen completa (contain) sobre un fondo desenfocado
 * de la misma foto — y se inserta con ese mismo tamaño. Se respeta la rotación
 * EXIF de las fotos de celular.
 *
 * Foto principal (analizarCatalogo): el nombre del archivo no dice qué muestra
 * cada foto, así que se decide por la IMAGEN, en este orden de confianza:
 *   1. Lector de códigos de barras (ZXing): si una foto contiene el código del
 *      propio activo es la etiqueta (certeza 'codigo'); otras lecturas son
 *      placas de datos del fabricante. Se reintenta girada ±30°/±60°.
 *   2. Densidad de barras + color/variedad para el resto ('imagen'/'revisar').
 * El resultado se revisa en pantalla antes de generar (app2.js) y llega aquí
 * como opts.seleccion.
 *
 * Usa PptxGenJS y ZXing (globales, vendorizados). No toca motor/archivos/reportes.
 * ============================================================================
 */
import { LOGO_AQUARIUS } from './logo-aquarius.js';

// Paleta Aquarius (entregable al cliente)
const C = {
  navy: '0F2B47', azul: '1E5A8E', accent: '2D7DD2', cielo: 'A7C4E0',
  texto: '1A2B3D', gris: '5A6B7E', gris2: '8899AA', borde: 'DFE6EE', fondo: 'F0F4F8', blanco: 'FFFFFF',
};
const FONT = 'Arial';
const DPI = 150;                       // resolución de composición de fotos
const esc = (t) => String(t ?? '').trim();
const corto = (t, n) => { const s = esc(t); return s.length > n ? s.slice(0, n - 1) + '…' : s; };

// ── Procesamiento de fotos ──────────────────────────────────

/** Lee la foto (rotación EXIF correcta), la reduce y calcula su "score" de toma del bien. */
async function procesarFoto(handle) {
  const file = await handle.getFile();
  const bmp = await createImageBitmap(file, { imageOrientation: 'from-image' });
  const sc = Math.min(1, 1100 / Math.max(bmp.width, bmp.height));
  const cv = document.createElement('canvas');
  cv.width = Math.max(1, Math.round(bmp.width * sc));
  cv.height = Math.max(1, Math.round(bmp.height * sc));
  cv.getContext('2d').drawImage(bmp, 0, 0, cv.width, cv.height);
  bmp.close?.();

  // Score: color/complejidad alto = bien; mucho blanco (papel) = etiqueta/datos
  const S = 64, ac = document.createElement('canvas'); ac.width = S; ac.height = S;
  const actx = ac.getContext('2d'); actx.drawImage(cv, 0, 0, S, S);
  const px = actx.getImageData(0, 0, S, S).data;
  let n = S * S, sr = 0, sy = 0, sr2 = 0, sy2 = 0, white = 0;
  for (let i = 0; i < px.length; i += 4) {
    const r = px[i], g = px[i + 1], b = px[i + 2];
    const rg = r - g, yb = 0.5 * (r + g) - b;
    sr += rg; sy += yb; sr2 += rg * rg; sy2 += yb * yb;
    if (r > 222 && g > 222 && b > 222) white++;
  }
  const mr = sr / n, my = sy / n;
  const sdr = Math.sqrt(Math.max(0, sr2 / n - mr * mr));
  const sdy = Math.sqrt(Math.max(0, sy2 / n - my * my));
  const colorful = Math.sqrt(sdr * sdr + sdy * sdy) + 0.3 * Math.sqrt(mr * mr + my * my);
  const barras = densidadBarras(cv);
  return { cv, score: colorful * (1 - 0.7 * (white / n)), barras, etiqueta: barras >= UMBRAL_BARRAS };
}

/**
 * Densidad de "barras": bordes fuertes con continuidad vertical (o horizontal,
 * si la foto está girada) en la mejor ventana de la imagen. Las etiquetas con
 * código de barras y las placas de datos dan valores altos; la toma del bien,
 * bajos. Calibrado con fotos reales de inventario: umbral 0.10.
 */
const UMBRAL_BARRAS = 0.10;
function densidadBarras(src) {
  const sc = 400 / Math.max(src.width, src.height);
  const w = Math.max(8, Math.round(src.width * sc)), h = Math.max(8, Math.round(src.height * sc));
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const ctx = c.getContext('2d'); ctx.drawImage(src, 0, 0, w, h);
  const d = ctx.getImageData(0, 0, w, h).data, g = new Float32Array(w * h);
  for (let p = 0, k = 0; p < d.length; p += 4, k++) g[k] = 0.299 * d[p] + 0.587 * d[p + 1] + 0.114 * d[p + 2];
  const L = 14, T = 30, mv = new Uint8Array(w * h), mh = new Uint8Array(w * h);
  for (let y = 1; y < h - 1 - L; y++) for (let x = 1; x < w - 1; x++) {
    const a = g[y * w + x + 1] - g[y * w + x - 1];
    if (Math.abs(a) <= T) continue;
    let ok = true;
    for (let k = 4; k <= L; k += 5) { const b = g[(y + k) * w + x + 1] - g[(y + k) * w + x - 1]; if (Math.abs(b) < T || Math.sign(b) !== Math.sign(a)) { ok = false; break; } }
    if (ok) mv[y * w + x] = 1;
  }
  for (let y = 1; y < h - 1; y++) for (let x = 1; x < w - 1 - L; x++) {
    const a = g[(y + 1) * w + x] - g[(y - 1) * w + x];
    if (Math.abs(a) <= T) continue;
    let ok = true;
    for (let k = 4; k <= L; k += 5) { const b = g[(y + 1) * w + x + k] - g[(y - 1) * w + x + k]; if (Math.abs(b) < T || Math.sign(b) !== Math.sign(a)) { ok = false; break; } }
    if (ok) mh[y * w + x] = 1;
  }
  const ws = Math.max(4, Math.round(Math.min(w, h) * 0.25)), paso = Math.max(1, Math.round(ws / 3));
  const mejor = (m) => {
    let mx = 0;
    for (let y = 0; y + ws <= h; y += paso) for (let x = 0; x + ws <= w; x += paso) {
      let s = 0; for (let yy = y; yy < y + ws; yy += 2) for (let xx = x; xx < x + ws; xx += 2) s += m[yy * w + xx];
      mx = Math.max(mx, s / ((ws / 2) * (ws / 2)));
    }
    return mx;
  };
  return Math.max(mejor(mv), mejor(mh));
}

/** Orden: primero fotos sin código de barras (toma del bien), por color/variedad. */
export function ordenarFotos(fotos) {
  return fotos.sort((a, b) => (a.etiqueta - b.etiqueta) || (b.score - a.score));
}
export { procesarFoto };

/**
 * Compone la foto en un marco de wIn×hIn pulgadas: fondo desenfocado (cover)
 * + foto completa (contain), esquinas redondeadas sobre `bg`. Nunca deforma.
 */
function componer(src, wIn, hIn, bg = '#FFFFFF', radioIn = 0.1) {
  const w = Math.round(wIn * DPI), h = Math.round(hIn * DPI);
  const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
  const c = cv.getContext('2d');
  c.fillStyle = bg; c.fillRect(0, 0, w, h);
  c.save();
  const r = Math.round(radioIn * DPI);
  c.beginPath();
  if (c.roundRect) c.roundRect(0, 0, w, h, r); else c.rect(0, 0, w, h);
  c.clip();
  const sw = src.width, sh = src.height;
  const sCover = Math.max(w / sw, h / sh);
  if ('filter' in c) {
    c.filter = 'blur(16px) brightness(0.82) saturate(1.15)';
    const cw = sw * sCover + 40, ch = sh * sCover + 40;
    c.drawImage(src, (w - cw) / 2, (h - ch) / 2, cw, ch);
    c.filter = 'none';
  } else { c.fillStyle = '#E4EAF0'; c.fillRect(0, 0, w, h); }
  const sFit = Math.min(w / sw, h / sh);
  const dw = sw * sFit, dh = sh * sFit;
  c.drawImage(src, (w - dw) / 2, (h - dh) / 2, dw, dh);
  c.restore();
  return cv.toDataURL('image/jpeg', 0.82);
}

/** Lee la foto reducida (orientación EXIF correcta), sin análisis. */
async function leerFoto(handle) {
  const bmp = await createImageBitmap(await handle.getFile(), { imageOrientation: 'from-image' });
  const sc = Math.min(1, 1100 / Math.max(bmp.width, bmp.height));
  const cv = document.createElement('canvas');
  cv.width = Math.max(1, Math.round(bmp.width * sc));
  cv.height = Math.max(1, Math.round(bmp.height * sc));
  cv.getContext('2d').drawImage(bmp, 0, 0, cv.width, cv.height);
  bmp.close?.();
  return cv;
}

// ── Lector de códigos de barras (ZXing, local) ──────────────
let lector = null;
/** Intenta leer un código de barras/QR en la imagen (girada "ang" grados). Devuelve el texto o null. */
function leerCodigo(src, ang = 0) {
  const Z = window.ZXing;
  if (!Z) return null;
  if (!lector) {
    lector = new Z.MultiFormatReader();
    const h = new Map(); h.set(Z.DecodeHintType.TRY_HARDER, true); lector.setHints(h);
  }
  const r = (ang * Math.PI) / 180, w = src.width, h = src.height;
  const W = Math.ceil(Math.abs(w * Math.cos(r)) + Math.abs(h * Math.sin(r)));
  const H = Math.ceil(Math.abs(w * Math.sin(r)) + Math.abs(h * Math.cos(r)));
  const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
  const c = cv.getContext('2d');
  c.fillStyle = '#FFFFFF'; c.fillRect(0, 0, W, H);
  c.translate(W / 2, H / 2); c.rotate(r); c.drawImage(src, -w / 2, -h / 2);
  const d = c.getImageData(0, 0, W, H).data, lum = new Uint8ClampedArray(W * H);
  for (let p = 0, k = 0; p < d.length; p += 4, k++) lum[k] = (d[p] * 77 + d[p + 1] * 150 + d[p + 2] * 29) >> 8;
  try {
    return lector.decode(new Z.BinaryBitmap(new Z.HybridBinarizer(new Z.RGBLuminanceSource(lum, W, H)))).getText();
  } catch { return null; }
}

/** Miniatura JPEG pequeña para la pantalla de revisión. */
function miniatura(cv, max = 240) {
  const sc = Math.min(1, max / Math.max(cv.width, cv.height));
  const m = document.createElement('canvas');
  m.width = Math.round(cv.width * sc); m.height = Math.round(cv.height * sc);
  m.getContext('2d').drawImage(cv, 0, 0, m.width, m.height);
  return m.toDataURL('image/jpeg', 0.72);
}

/**
 * Analiza todas las fotos y propone la principal de cada activo.
 * @returns {{familias, total, totalFotos, seleccion: Map<codigo,
 *   {principal:number, certeza:'codigo'|'imagen'|'revisar', thumbs:string[], tipos:string[]}>}}
 *   tipos[i]: 'etiqueta' (código propio leído), 'datos' (otro código leído),
 *             'probable-etiqueta' (muchas barras), 'foto'.
 */
export async function analizarCatalogo(plan, onProgress) {
  const { familias, total, totalFotos } = agrupar(plan);
  const seleccion = new Map();
  let hechos = 0;
  for (const fam of familias) for (const reg of fam.activos) {
    const cod = String(reg.activo.codigo).toUpperCase();
    const esPropio = (t) => !!t && t.toUpperCase().includes(cod);
    const fotos = [];
    for (let i = 0; i < reg.items.length; i++) {
      let x = { idx: i, cv: null, score: -1, barras: 0, etiqueta: false, texto: null };
      try {
        const p = await procesarFoto(reg.items[i].handle);
        x = { ...x, ...p, texto: leerCodigo(p.cv) };
      } catch { /* formato no legible (HEIC…) */ }
      fotos.push(x);
      onProgress?.(++hechos, totalFotos);
    }
    // Si no se leyó la etiqueta, reintentar girando las fotos con pinta de etiqueta
    if (!fotos.some((x) => esPropio(x.texto))) {
      const cands = fotos.filter((x) => x.cv && !x.texto && x.barras >= 0.04).sort((a, b) => b.barras - a.barras);
      buscar: for (const c of cands) {
        for (const a of [30, -30, 60, -60]) {
          const t = leerCodigo(c.cv, a);
          if (t) { c.texto = t; if (esPropio(t)) break buscar; break; }
        }
      }
    }
    const propia = fotos.some((x) => esPropio(x.texto));
    const legibles = fotos.filter((x) => x.cv);
    const libres = legibles.filter((x) => !x.texto);
    const orden = ordenarFotos((libres.length ? libres : legibles).slice());
    const top = orden[0];
    const otraMarcada = top && fotos.some((x) => x !== top && (x.texto || x.etiqueta));
    const certeza = propia && libres.length === 1 ? 'codigo'
      : top && !top.etiqueta && otraMarcada ? 'imagen' : 'revisar';
    seleccion.set(reg.activo.codigo, {
      principal: top ? top.idx : 0,
      certeza,
      thumbs: fotos.map((x) => (x.cv ? miniatura(x.cv) : null)),
      tipos: fotos.map((x) => (esPropio(x.texto) ? 'etiqueta' : x.texto ? 'datos' : x.etiqueta ? 'probable-etiqueta' : 'foto')),
    });
    fotos.forEach((x) => { x.cv = null; });
  }
  return { familias, total, totalFotos, seleccion };
}

/** Lee las fotos del activo en el orden elegido: principal primero, luego las demás. */
async function fotosDeActivo(reg, sel, tick) {
  const principal = sel?.principal ?? 0;
  const orden = [principal, ...reg.items.map((_, i) => i).filter((i) => i !== principal)];
  const fotos = [];
  for (const i of orden) {
    try { fotos.push({ cv: await leerFoto(reg.items[i].handle) }); } catch { /* ilegible */ }
    tick();
  }
  return fotos;
}

// ── Agrupación ──────────────────────────────────────────────
function agrupar(plan) {
  const porCodigo = new Map();
  for (const it of plan.items) {
    if (!it.activo) continue;
    const k = it.activo.codigo;
    if (!porCodigo.has(k)) porCodigo.set(k, { activo: it.activo, items: [] });
    porCodigo.get(k).items.push(it);
  }
  const fam = new Map();
  for (const reg of porCodigo.values()) {
    const f = esc(reg.activo.familia) || 'SIN FAMILIA';
    const s = esc(reg.activo.subfamilia) || 'SIN SUBFAMILIA';
    if (!fam.has(f)) fam.set(f, new Map());
    if (!fam.get(f).has(s)) fam.get(f).set(s, []);
    fam.get(f).get(s).push(reg);
  }
  // Familias por cantidad; dentro, subfamilias por cantidad y activos por código
  const familias = [...fam.entries()].map(([f, subs]) => {
    const subsOrd = [...subs.entries()].sort((a, b) => b[1].length - a[1].length);
    const activos = subsOrd.flatMap(([, arr]) =>
      arr.sort((a, b) => String(a.activo.codigo).localeCompare(String(b.activo.codigo), 'es')));
    const conSub = activos.some((r) => esc(r.activo.subfamilia));
    return { familia: f, nSubs: subs.size, conSub, activos, nFotos: activos.reduce((a, r) => a + r.items.length, 0) };
  }).sort((a, b) => b.activos.length - a.activos.length);
  const totalFotos = familias.reduce((a, f) => a + f.nFotos, 0);
  return { familias, total: porCodigo.size, totalFotos };
}

// ── Layout (pulgadas, 16:9 — 13.333 × 7.5) ─────────────────
const W = 13.333, H = 7.5;
const P = { top: 1.02, h: 5.92, mx: 0.45, gap: 0.33, pad: 0.2, fotoH: 3.6 };
P.w = (W - P.mx * 2 - P.gap) / 2;

function encabezado(slide, titulo, sub, cliente) {
  slide.background = { color: C.fondo };
  slide.addShape('rect', { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.navy } });
  slide.addShape('rect', { x: P.mx, y: 0.36, w: 0.08, h: 0.42, fill: { color: C.accent } });
  slide.addText([
    { text: esc(titulo).toUpperCase(), options: { bold: true, color: C.navy } },
    ...(sub ? [{ text: `   ·   ${sub}`, options: { color: C.gris } }] : []),
  ], { x: P.mx + 0.2, y: 0.32, w: 9, h: 0.5, fontSize: 16, fontFace: FONT, valign: 'middle' });
  slide.addText(corto(cliente, 58).toUpperCase(), { x: W - 6.4, y: 0.32, w: 5.95, h: 0.5, fontSize: 8, color: C.gris2, fontFace: FONT, align: 'right', valign: 'middle', charSpacing: 1, fit: 'shrink' });
}

function pie(slide, n, oscuro = false) {
  slide.addText([
    { text: 'AQUARIUS CONSULTING', options: { bold: true, color: oscuro ? C.blanco : C.navy } },
    { text: '   ·   Catálogo fotográfico de activos fijos', options: { color: C.gris2 } },
  ], { x: P.mx, y: H - 0.42, w: 8, h: 0.3, fontSize: 8, fontFace: FONT, charSpacing: 1 });
  slide.addText(String(n).padStart(2, '0'), { x: W - 1.2, y: H - 0.44, w: 0.75, h: 0.32, fontSize: 10, bold: true, color: C.accent, fontFace: FONT, align: 'right' });
}

/** Tarjeta de un activo en el panel izquierdo (lado 0) o derecho (lado 1). */
function tarjeta(slide, x, reg, fotos, extra) {
  const { activo, items } = reg;
  const y = P.top, pad = P.pad;
  const iw = P.w - pad * 2;

  slide.addShape('roundRect', {
    x, y, w: P.w, h: P.h, rectRadius: 0.12, fill: { color: C.blanco }, line: { color: C.borde, width: 0.75 },
    shadow: { type: 'outer', blur: 10, offset: 3, angle: 90, color: '1E293B', opacity: 0.12 },
  });

  // ── Fotos ──
  const fy = y + pad;
  const thumbs = fotos.slice(1, 3);
  const mainW = thumbs.length ? iw * 0.62 : iw;
  const main = fotos[0];
  if (main) slide.addImage({ data: componer(main.cv, mainW, P.fotoH), x: x + pad, y: fy, w: mainW, h: P.fotoH });
  else {
    slide.addShape('roundRect', { x: x + pad, y: fy, w: mainW, h: P.fotoH, rectRadius: 0.1, fill: { color: 'E9EEF3' } });
    slide.addText('Sin foto disponible', { x: x + pad, y: fy, w: mainW, h: P.fotoH, align: 'center', valign: 'middle', fontSize: 11, color: C.gris2, fontFace: FONT });
  }
  if (thumbs.length) {
    const tx = x + pad + mainW + 0.12, tw = iw - mainW - 0.12;
    const th = thumbs.length === 2 ? (P.fotoH - 0.12) / 2 : P.fotoH;
    thumbs.forEach((t, i) => slide.addImage({ data: componer(t.cv, tw, th), x: tx, y: fy + i * (th + 0.12), w: tw, h: th }));
  }

  // Etiqueta con el código sobre la foto (estilo "precio")
  const cod = esc(activo.codigo).toUpperCase();
  const tagW = Math.min(mainW - 0.3, 0.42 + cod.length * 0.115);
  slide.addShape('rect', { x: x + pad, y: fy + 0.22, w: tagW, h: 0.4, fill: { color: C.accent } });
  slide.addText(cod, { x: x + pad, y: fy + 0.22, w: tagW, h: 0.4, fontSize: 13, bold: true, color: C.blanco, fontFace: FONT, align: 'center', valign: 'middle' });
  // Contador de fotos
  const nF = items.length;
  slide.addShape('roundRect', { x: x + pad + mainW - 0.95, y: fy + P.fotoH - 0.42, w: 0.8, h: 0.28, rectRadius: 0.14, fill: { color: C.navy, transparency: 20 } });
  slide.addText(`${nF} ${nF === 1 ? 'foto' : 'fotos'}`, { x: x + pad + mainW - 0.95, y: fy + P.fotoH - 0.42, w: 0.8, h: 0.28, fontSize: 8, bold: true, color: C.blanco, fontFace: FONT, align: 'center', valign: 'middle' });

  // ── Ficha ──
  let ty = fy + P.fotoH + 0.16;
  const eyebrow = [esc(activo.familia), esc(activo.subfamilia)].filter(Boolean).join('  ·  ');
  slide.addText(corto(eyebrow, 70).toUpperCase(), { x: x + pad, y: ty, w: iw, h: 0.24, fontSize: 8.5, bold: true, color: C.accent, fontFace: FONT, charSpacing: 1.5 });
  ty += 0.24;
  slide.addText(corto(activo.descripcion || 'Activo sin descripción', 64).toUpperCase(), {
    x: x + pad, y: ty, w: iw, h: 0.62, fontSize: 17, bold: true, italic: true, color: C.navy, fontFace: FONT, valign: 'top', lineSpacingMultiple: 0.9,
  });
  ty += 0.64;
  if (extra?.ubicacion) {
    slide.addText(`Ubicación: ${corto(extra.ubicacion, 70)}`, { x: x + pad, y: ty, w: iw, h: 0.22, fontSize: 8.5, color: C.gris, fontFace: FONT });
  }

  // Fila de especificaciones (Marca · Modelo · Serie · Estado)
  const sy = y + P.h - 0.78;
  slide.addShape('line', { x: x + pad, y: sy, w: iw, h: 0, line: { color: C.borde, width: 0.75 } });
  const specs = [['MARCA', extra?.marca], ['MODELO', extra?.modelo], ['SERIE', extra?.serie], ['ESTADO', extra?.estado]];
  const cw = iw / specs.length;
  specs.forEach(([lab, v], i) => {
    const cx = x + pad + i * cw;
    if (i) slide.addShape('line', { x: cx, y: sy + 0.14, w: 0, h: 0.46, line: { color: C.borde, width: 0.75 } });
    slide.addText(lab, { x: cx + 0.08, y: sy + 0.1, w: cw - 0.12, h: 0.2, fontSize: 7, bold: true, color: C.gris2, fontFace: FONT, charSpacing: 1.5 });
    slide.addText(corto(v || '—', 18).toUpperCase(), { x: cx + 0.08, y: sy + 0.32, w: cw - 0.12, h: 0.3, fontSize: 10.5, bold: true, color: v ? C.texto : C.gris2, fontFace: FONT });
  });
}

/**
 * @param {Object}   plan        plan del organizador
 * @param {string}   cliente     nombre del inventario
 * @param {Function} onProgress  (fotosHechas, fotosTotal) => void
 * @param {Object}   opts        { extras: Map<codigo,{marca,modelo,serie,estado,ubicacion}>,
 *                                 seleccion: Map<codigo,{principal}> (de analizarCatalogo, revisada) }
 */
export async function generarCatalogoPPT(plan, cliente, onProgress, opts = {}) {
  const extras = opts.extras || new Map();
  const seleccion = opts.seleccion || (await analizarCatalogo(plan)).seleccion;
  const { familias, total, totalFotos } = agrupar(plan);
  if (!total) throw new Error('No hay activos con foto para el catálogo.');

  // Inventarios grandes: un archivo por cada MAX activos (PptxGenJS arma todo en memoria).
  const partes = partir(familias, opts.maxActivos || MAX_ACTIVOS_POR_ARCHIVO);
  let hechos = 0;
  const tick = () => { hechos++; onProgress?.(hechos, totalFotos); };
  for (let i = 0; i < partes.length; i++) {
    const pptx = await armarParte(partes[i], { cliente, extras, seleccion, tick, parte: i + 1, partes: partes.length });
    const sufijo = partes.length > 1 ? `_parte${i + 1}de${partes.length}` : '';
    await pptx.writeFile({ fileName: `Catalogo_Fotografico_${cliente}${sufijo}.pptx` });
  }
  return partes.length;
}

const MAX_ACTIVOS_POR_ARCHIVO = 350;

/** Reparte las familias en grupos de hasta `max` activos; una familia puede continuar en la parte siguiente. */
function partir(familias, max) {
  const partes = [];
  let actual = [], n = 0;
  for (const f of familias) {
    let desde = 0;
    while (desde < f.activos.length) {
      const toma = f.activos.slice(desde, desde + (max - n));
      actual.push({ ...f, activos: toma, nFotos: toma.reduce((a, r) => a + r.items.length, 0), offset: desde, totalFam: f.activos.length, nFotosFam: f.nFotos });
      desde += toma.length; n += toma.length;
      if (n >= max) { partes.push(actual); actual = []; n = 0; }
    }
  }
  if (actual.length) partes.push(actual);
  return partes;
}

async function armarParte(familias, { cliente, extras, seleccion, tick, parte, partes }) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Aquarius Consulting';
  pptx.title = `Catálogo fotográfico — ${cliente}`;
  const total = familias.reduce((a, f) => a + f.activos.length, 0);
  const totalFotos = familias.reduce((a, f) => a + f.nFotos, 0);
  // Caché: portada y separadores "miran" fotos ya procesadas; la tarjeta las consume y libera.
  const cache = new Map();
  const ver = async (reg) => {
    if (!cache.has(reg.activo.codigo)) cache.set(reg.activo.codigo, await fotosDeActivo(reg, seleccion.get(reg.activo.codigo), tick));
    return cache.get(reg.activo.codigo);
  };
  const fotosDe = async (reg) => { const f = await ver(reg); cache.delete(reg.activo.codigo); return f; };
  const fecha = new Date().toLocaleDateString('es-PE', { day: '2-digit', month: 'long', year: 'numeric' });
  let nSlide = 0;

  // ── Portada con mosaico ──
  const cover = pptx.addSlide(); nSlide++;
  cover.background = { color: C.navy };
  const mosaico = familias.flatMap((f) => f.activos).slice(0, 6);
  const mx0 = 7.25, my0 = 0.55, mw = 2.75, mh = 2.05, mg = 0.13;
  for (let i = 0; i < mosaico.length; i++) {
    const f = await ver(mosaico[i]);
    if (!f[0]) continue;
    const col = i % 2, row = Math.floor(i / 2);
    cover.addImage({ data: componer(f[0].cv, mw, mh, '#0F2B47', 0.08), x: mx0 + col * (mw + mg), y: my0 + row * (mh + mg), w: mw, h: mh });
  }
  cover.addShape('rect', { x: 0, y: 0, w: 0.16, h: H, fill: { color: C.accent } });
  cover.addShape('roundRect', { x: 0.7, y: 0.55, w: 2.5, h: 1.32, rectRadius: 0.12, fill: { color: C.blanco }, shadow: { type: 'outer', blur: 12, offset: 4, angle: 90, color: '000000', opacity: 0.25 } });
  cover.addImage({ data: LOGO_AQUARIUS, x: 0.95, y: 0.68, w: 2.0, h: 1.06, sizing: { type: 'contain', w: 2.0, h: 1.06 } });
  cover.addText('CATÁLOGO', { x: 0.7, y: 2.35, w: 6.4, h: 0.9, fontSize: 48, bold: true, color: C.blanco, fontFace: FONT, charSpacing: 2 });
  cover.addText('FOTOGRÁFICO', { x: 0.7, y: 3.1, w: 6.4, h: 0.9, fontSize: 48, bold: true, italic: true, color: '7FB3E8', fontFace: FONT, charSpacing: 2 });
  cover.addShape('rect', { x: 0.72, y: 4.12, w: 1.1, h: 0.06, fill: { color: C.accent } });
  cover.addText(partes > 1 ? `INVENTARIO DE ACTIVOS FIJOS   ·   PARTE ${parte} DE ${partes}` : 'INVENTARIO DE ACTIVOS FIJOS',
    { x: 0.7, y: 4.3, w: 6.4, h: 0.4, fontSize: 14, color: C.cielo, fontFace: FONT, charSpacing: 3 });
  cover.addText(corto(cliente, 60).toUpperCase(), { x: 0.7, y: 4.85, w: 6.4, h: 0.5, fontSize: 15, bold: true, color: C.blanco, fontFace: FONT });
  const stats = [[total, 'ACTIVOS'], [totalFotos, 'FOTOGRAFÍAS'], [familias.length, 'FAMILIAS']];
  stats.forEach(([v, l], i) => {
    cover.addText(Number(v).toLocaleString('es-PE'), { x: 0.7 + i * 2.05, y: 5.55, w: 1.95, h: 0.6, fontSize: 30, bold: true, color: C.blanco, fontFace: FONT });
    cover.addText(l, { x: 0.7 + i * 2.05, y: 6.12, w: 1.95, h: 0.3, fontSize: 9, color: C.cielo, fontFace: FONT, charSpacing: 2 });
  });
  cover.addText(`${fecha}   ·   ATF-PR-01-FO-03 / Ver.02`, { x: 0.7, y: 6.85, w: 6.4, h: 0.3, fontSize: 9, color: C.gris2, fontFace: FONT });

  // ── Contenido ──
  const idx = pptx.addSlide(); nSlide++;
  encabezado(idx, 'Contenido', `${familias.length} familias${partes > 1 ? ` · parte ${parte} de ${partes}` : ''}`, cliente);
  idx.addTable([
    [{ text: 'N°', options: { bold: true, color: C.blanco, fill: { color: C.navy } } },
     { text: 'FAMILIA', options: { bold: true, color: C.blanco, fill: { color: C.navy } } },
     { text: 'SUBFAMILIAS', options: { bold: true, color: C.blanco, fill: { color: C.navy }, align: 'right' } },
     { text: 'ACTIVOS', options: { bold: true, color: C.blanco, fill: { color: C.navy }, align: 'right' } },
     { text: 'FOTOS', options: { bold: true, color: C.blanco, fill: { color: C.navy }, align: 'right' } }],
    ...familias.slice(0, 14).map((f, i) => [
      { text: String(i + 1) }, { text: f.familia.toUpperCase(), options: { bold: true, color: C.navy } },
      { text: f.conSub ? String(f.nSubs) : '—', options: { align: 'right' } },
      { text: f.activos.length.toLocaleString('es-PE'), options: { align: 'right', bold: true } },
      { text: f.nFotos.toLocaleString('es-PE'), options: { align: 'right' } }]),
  ], { x: P.mx, y: 1.15, w: W - P.mx * 2, colW: [0.7, 7.6, 1.5, 1.32, 1.32], fontSize: 11, fontFace: FONT, color: C.texto, border: { type: 'solid', pt: 0.5, color: C.borde }, fill: { color: C.blanco }, rowH: 0.36 });
  pie(idx, nSlide);

  // ── Familias ──
  for (const f of familias) {
    const sep = pptx.addSlide(); nSlide++;
    sep.background = { color: C.blanco };
    sep.addShape('rect', { x: 0, y: 0, w: 5.6, h: H, fill: { color: C.navy } });
    sep.addShape('rect', { x: 5.6, y: 0, w: 0.12, h: H, fill: { color: C.accent } });
    sep.addText(f.offset ? 'FAMILIA · CONTINUACIÓN' : 'FAMILIA', { x: 0.7, y: 2.4, w: 4.6, h: 0.4, fontSize: 12, color: C.cielo, fontFace: FONT, charSpacing: 4 });
    sep.addText(f.familia.toUpperCase(), { x: 0.7, y: 2.85, w: 4.7, h: 1.9, fontSize: 32, bold: true, color: C.blanco, fontFace: FONT, valign: 'top' });
    sep.addShape('rect', { x: 0.72, y: 4.85, w: 0.9, h: 0.06, fill: { color: C.accent } });
    const cifras = [[f.totalFam, 'ACTIVOS'], [f.nFotosFam, 'FOTOGRAFÍAS'], ...(f.conSub ? [[f.nSubs, 'SUBFAMILIAS']] : [])];
    cifras.forEach(([v, l], i) => {
      sep.addText(Number(v).toLocaleString('es-PE'), { x: 6.3 + i * 2.25, y: 0.55, w: 2.15, h: 0.85, fontSize: 40, bold: true, color: C.navy, fontFace: FONT });
      sep.addText(l, { x: 6.33 + i * 2.25, y: 1.38, w: 2.15, h: 0.3, fontSize: 9.5, color: C.gris, fontFace: FONT, charSpacing: 2 });
    });
    // Franja de fotos representativas (2×2)
    const muestra = f.activos.slice(0, 4);
    const gw = 3.33, gh = 2.42, g0x = 6.3, g0y = 2.05, gg = 0.16;
    for (let i = 0; i < muestra.length; i++) {
      const fo = await ver(muestra[i]);
      if (!fo[0]) continue;
      sep.addImage({ data: componer(fo[0].cv, gw, gh, '#FFFFFF', 0.1), x: g0x + (i % 2) * (gw + gg), y: g0y + Math.floor(i / 2) * (gh + gg), w: gw, h: gh });
    }
    pie(sep, nSlide, true);

    // Láminas: 2 activos por lámina (la numeración respeta el total de la familia)
    for (let p = 0; p < f.activos.length; p += 2) {
      const par = f.activos.slice(p, p + 2);
      const slide = pptx.addSlide(); nSlide++;
      encabezado(slide, f.familia, `${f.offset + Math.min(p + 2, f.activos.length)} de ${f.totalFam}`, cliente);
      for (let i = 0; i < par.length; i++) {
        const fotos = await fotosDe(par[i]);
        // Si queda un solo activo, se centra
        const x = par.length === 1 ? (W - P.w) / 2 : P.mx + i * (P.w + P.gap);
        tarjeta(slide, x, par[i], fotos, extras.get(par[i].activo.codigo));
      }
      pie(slide, nSlide);
    }
  }
  return pptx;
}
