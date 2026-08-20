/**
 * ============================================================================
 * catalogo.js — Catálogo fotográfico de inventario en PowerPoint (.pptx)
 * ----------------------------------------------------------------------------
 * A partir del plan del organizador arma un PPT tipo catálogo: portada
 * Aquarius, un separador por familia, y diapositivas con una tarjeta por
 * activo (agrupadas por familia/subfamilia — igual que las carpetas).
 *
 * Multi-foto: si el activo tiene 2-3 fotos, muestra la TOMA PANORÁMICA del
 * bien en grande y las demás (etiqueta / datos) como miniaturas. Como el
 * nombre del archivo no dice qué contiene cada foto, se identifica por la
 * IMAGEN: la más "colorida/compleja" es el bien; las de papel blanco con
 * barras/texto (etiqueta, placa de datos) van como miniaturas.
 *
 * Usa PptxGenJS (global, vendorizado). No toca motor/archivos/reportes.
 * ============================================================================
 */
import { LOGO_AQUARIUS } from './logo-aquarius.js';

const NAVY = '0F2B47', ACCENT = '2D7DD2', GRIS = '5A6B7E', GRIS2 = '8899AA', BORDE = 'E2E8F0';
const esc = (t) => String(t ?? '');

/**
 * Lee una foto: produce el JPEG reducido (dataURL) para el PPT y un "score"
 * de qué tan probable es que sea la toma del bien (más color = más probable;
 * mucho blanco = etiqueta/datos → score bajo).
 */
async function procesarFoto(handle) {
  const file = await handle.getFile();
  const bmp = await createImageBitmap(file);
  // Imagen para mostrar (máx 680px)
  const sc = Math.min(1, 680 / Math.max(bmp.width, bmp.height));
  const w = Math.max(1, Math.round(bmp.width * sc));
  const h = Math.max(1, Math.round(bmp.height * sc));
  const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
  cv.getContext('2d').drawImage(bmp, 0, 0, w, h);
  const data = cv.toDataURL('image/jpeg', 0.72);
  // Miniatura 64×64 para analizar color y proporción de blanco
  const S = 64, ac = document.createElement('canvas'); ac.width = S; ac.height = S;
  const actx = ac.getContext('2d'); actx.drawImage(bmp, 0, 0, S, S);
  const px = actx.getImageData(0, 0, S, S).data;
  bmp.close?.();
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
  const score = colorful * (1 - 0.7 * (white / n));   // alto = bien; bajo = etiqueta/datos
  return { data, score };
}

/** Agrupa los activos emparejados por familia→subfamilia, conservando TODAS sus fotos. */
function agrupar(plan) {
  const porCodigo = new Map();
  for (const it of plan.items) {
    if (!it.activo) continue;
    if (!porCodigo.has(it.activo.codigo)) porCodigo.set(it.activo.codigo, { activo: it.activo, items: [] });
    porCodigo.get(it.activo.codigo).items.push(it);
  }
  const fam = new Map();
  for (const reg of porCodigo.values()) {
    const f = (String(reg.activo.familia || '').trim()) || 'SIN FAMILIA';
    const s = (String(reg.activo.subfamilia || '').trim()) || 'SIN SUBFAMILIA';
    if (!fam.has(f)) fam.set(f, new Map());
    const sf = fam.get(f);
    if (!sf.has(s)) sf.set(s, []);
    sf.get(s).push(reg);
  }
  const cnt = (sf) => [...sf.values()].reduce((a, arr) => a + arr.length, 0);
  const famOrd = [...fam.entries()].sort((a, b) => cnt(b[1]) - cnt(a[1]));
  for (const [, sf] of famOrd) {
    for (const [, arr] of sf) arr.sort((a, b) => String(a.activo.codigo).localeCompare(String(b.activo.codigo), 'es'));
  }
  const totalFotos = [...porCodigo.values()].reduce((a, r) => a + r.items.length, 0);
  return { fam: famOrd.map(([f, sf]) => [f, [...sf.entries()].sort((a, b) => b[1].length - a[1].length)]), total: porCodigo.size, totalFotos };
}

// Layout (pulgadas, 13.33×7.5 — 16:9): 3 columnas × 2 filas = 6 tarjetas/slide
const L = { W: 13.33, H: 7.5, mx: 0.5, cols: 3, gap: 0.35, top: 1.5, cardH: 2.5, raise: 0.25, photoH: 1.4, thumbH: 0.52, thumbW: 0.74 };
const cardW = () => (L.W - L.mx * 2 - L.gap * (L.cols - 1)) / L.cols;

function pie(slide, cliente) {
  slide.addText([
    { text: 'Aquarius Consulting SAC', options: { bold: true } },
    { text: `   ·   Catálogo fotográfico   ·   ${cliente}`, options: {} },
  ], { x: L.mx, y: L.H - 0.4, w: L.W - 1, h: 0.3, fontSize: 8, color: GRIS2, fontFace: 'Arial' });
}

function encabezado(slide, familia, subfam) {
  slide.addShape('rect', { x: L.mx, y: 1.12, w: 0.12, h: 0.42, fill: { color: ACCENT } });
  slide.addText([
    { text: esc(familia).toUpperCase(), options: { bold: true, color: NAVY } },
    { text: `   ›   ${esc(subfam)}`, options: { color: GRIS } },
  ], { x: L.mx + 0.22, y: 1.1, w: L.W - 3, h: 0.46, fontSize: 15, fontFace: 'Arial', valign: 'middle' });
}

function imgOrBox(slide, data, o) {
  if (data) slide.addImage({ data, x: o.x, y: o.y, w: o.w, h: o.h, sizing: { type: 'cover', w: o.w, h: o.h }, shadow: o.shadow });
  else {
    slide.addShape('roundRect', { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.03, fill: { color: 'EEF1F5' }, line: { color: BORDE, width: 1 } });
    slide.addText('Sin vista', { x: o.x, y: o.y, w: o.w, h: o.h, align: 'center', valign: 'middle', fontSize: 8, color: GRIS2, fontFace: 'Arial' });
  }
}

function tarjeta(slide, col, row, activo, principal, thumbs, nFotos) {
  const w = cardW();
  const x = L.mx + col * (w + L.gap);
  const photoTop = L.top + row * (L.cardH + L.raise + 0.28);
  const cardTop = photoTop + L.raise;

  slide.addShape('roundRect', {
    x, y: cardTop, w, h: L.cardH, rectRadius: 0.06,
    fill: { color: 'FFFFFF' }, line: { color: BORDE, width: 1 },
    shadow: { type: 'outer', blur: 6, offset: 3, angle: 90, color: GRIS2, opacity: 0.35 },
  });
  // Foto principal (bien) superpuesta y elevada
  imgOrBox(slide, principal, { x: x + 0.14, y: photoTop, w: w - 0.28, h: L.photoH, shadow: { type: 'outer', blur: 5, offset: 2, angle: 90, color: GRIS2, opacity: 0.45 } });

  // Título (código) + descripción
  const ty = cardTop + L.photoH - L.raise + 0.06;
  slide.addText(esc(activo.codigo).toUpperCase(), { x: x + 0.16, y: ty, w: w - 0.32, h: 0.26, fontSize: 12, bold: true, color: NAVY, fontFace: 'Arial' });
  if (activo.descripcion) slide.addText(esc(activo.descripcion), { x: x + 0.16, y: ty + 0.28, w: w - 0.32, h: 0.3, fontSize: 8.5, color: GRIS, fontFace: 'Arial' });

  // Miniaturas (etiqueta / datos)
  const thumbY = cardTop + L.cardH - L.thumbH - 0.14;
  thumbs.slice(0, 2).forEach((t, i) => {
    imgOrBox(slide, t, { x: x + 0.16 + i * (L.thumbW + 0.12), y: thumbY, w: L.thumbW, h: L.thumbH, shadow: { type: 'outer', blur: 3, offset: 1, angle: 90, color: GRIS2, opacity: 0.35 } });
  });
  if (nFotos > 3) slide.addText(`+${nFotos - 3}`, { x: x + 0.16 + 2 * (L.thumbW + 0.12), y: thumbY, w: 0.6, h: L.thumbH, valign: 'middle', fontSize: 9, color: GRIS2, fontFace: 'Arial' });
  // Pie de tarjeta (subfamilia)
  slide.addText(esc(activo.subfamilia || activo.familia || '').toUpperCase(), { x: w - 1.9 + x, y: thumbY + 0.02, w: 1.74, h: 0.24, align: 'right', fontSize: 7.5, bold: true, color: ACCENT, fontFace: 'Arial' });
}

/**
 * @param {Object} plan   plan del organizador
 * @param {string} cliente nombre del inventario
 * @param {Function} onProgress (fotosHechas, fotosTotal) => void
 */
export async function generarCatalogoPPT(plan, cliente, onProgress) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Aquarius Consulting';
  pptx.title = `Catálogo fotográfico — ${cliente}`;

  const { fam, total, totalFotos } = agrupar(plan);
  const fecha = new Date().toLocaleDateString('es-PE', { day: '2-digit', month: 'long', year: 'numeric' });

  // ── Portada ──
  const cover = pptx.addSlide();
  cover.background = { color: NAVY };
  cover.addShape('rect', { x: 0, y: 5.55, w: L.W, h: 0.08, fill: { color: ACCENT } });
  cover.addImage({ data: LOGO_AQUARIUS, x: L.mx, y: 0.5, w: 2.2, h: 1.2, sizing: { type: 'contain', w: 2.2, h: 1.2 } });
  cover.addText('CATÁLOGO FOTOGRÁFICO', { x: L.mx, y: 2.5, w: 11, h: 0.9, fontSize: 40, bold: true, color: 'FFFFFF', fontFace: 'Arial' });
  cover.addText('DE INVENTARIO DE ACTIVOS FIJOS', { x: L.mx, y: 3.45, w: 11, h: 0.6, fontSize: 20, color: 'A7C4E0', fontFace: 'Arial' });
  cover.addText([
    { text: esc(cliente).toUpperCase(), options: { bold: true, fontSize: 16, color: 'FFFFFF' } },
    { text: `\n${total.toLocaleString('es-PE')} activos · ${totalFotos.toLocaleString('es-PE')} fotos   ·   ${fecha}`, options: { fontSize: 12, color: 'A7C4E0' } },
  ], { x: L.mx, y: 4.4, w: 11, h: 1, fontFace: 'Arial' });
  cover.addText('ATF-PR-01-FO-03 / Ver.02   ·   Generado con NEXOVA Suite', { x: L.mx, y: 6.9, w: 11, h: 0.3, fontSize: 9, color: GRIS2, fontFace: 'Arial' });

  let hechos = 0;
  for (const [familia, subs] of fam) {
    const nFam = subs.reduce((a, [, arr]) => a + arr.length, 0);
    const sep = pptx.addSlide();
    sep.addShape('rect', { x: 0, y: 0, w: 0.35, h: L.H, fill: { color: NAVY } });
    sep.addShape('rect', { x: L.mx, y: 3.0, w: 0.16, h: 1.0, fill: { color: ACCENT } });
    sep.addText(esc(familia).toUpperCase(), { x: L.mx + 0.35, y: 3.0, w: 11.5, h: 1.0, fontSize: 34, bold: true, color: NAVY, fontFace: 'Arial', valign: 'middle' });
    sep.addText(`${nFam.toLocaleString('es-PE')} activos   ·   ${subs.length} subfamilia(s)`, { x: L.mx + 0.37, y: 4.05, w: 11, h: 0.4, fontSize: 13, color: GRIS, fontFace: 'Arial' });
    pie(sep, cliente);

    for (const [subfam, activos] of subs) {
      for (let p = 0; p < activos.length; p += 6) {
        const slide = pptx.addSlide();
        encabezado(slide, familia, subfam);
        pie(slide, cliente);
        const chunk = activos.slice(p, p + 6);
        for (let idx = 0; idx < chunk.length; idx++) {
          const { activo, items } = chunk[idx];
          // Procesar TODAS las fotos del activo y ordenar por "score" (bien primero)
          const fotos = [];
          for (const it of items) {
            let f = { data: null, score: -1 };
            try { f = await procesarFoto(it.handle); } catch { /* HEIC u otro */ }
            hechos++; onProgress?.(hechos, totalFotos);
            fotos.push(f);
          }
          fotos.sort((a, b) => b.score - a.score);
          const principal = fotos[0]?.data || null;
          const thumbs = fotos.slice(1).map((f) => f.data);
          tarjeta(slide, idx % L.cols, Math.floor(idx / L.cols), activo, principal, thumbs, items.length);
        }
      }
    }
  }

  await pptx.writeFile({ fileName: `Catalogo_Fotografico_${cliente}.pptx` });
}
