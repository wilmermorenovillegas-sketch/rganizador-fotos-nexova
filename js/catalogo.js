/**
 * ============================================================================
 * catalogo.js — Catálogo fotográfico de inventario en PowerPoint (.pptx)
 * ----------------------------------------------------------------------------
 * Genera, a partir del plan del organizador, un PPT tipo catálogo: portada
 * Aquarius, un separador por familia, y diapositivas con tarjetas (una por
 * activo) agrupadas por familia/subfamilia — igual que las carpetas. Cada
 * tarjeta lleva la foto "superpuesta" (elevada sobre el recuadro).
 *
 * Usa PptxGenJS (global, vendorizado). No toca motor/archivos/reportes.
 * ============================================================================
 */
import { LOGO_AQUARIUS } from './logo-aquarius.js';

// Paleta Aquarius
const NAVY = '0F2B47', AZUL = '1E5A8E', ACCENT = '2D7DD2', GRIS = '5A6B7E', GRIS2 = '8899AA', BORDE = 'E2E8F0';

/** Reduce una imagen a JPEG pequeño (dataURL) para no inflar el PPT. */
async function imgDataURL(handle, max = 720) {
  const file = await handle.getFile();
  const bmp = await createImageBitmap(file);
  const scale = Math.min(1, max / Math.max(bmp.width, bmp.height));
  const w = Math.max(1, Math.round(bmp.width * scale));
  const h = Math.max(1, Math.round(bmp.height * scale));
  const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
  cv.getContext('2d').drawImage(bmp, 0, 0, w, h);
  bmp.close?.();
  return cv.toDataURL('image/jpeg', 0.72);
}

/** Agrupa los activos emparejados (uno por código, foto principal) por familia→subfamilia. */
function agrupar(plan) {
  const porCodigo = new Map();
  for (const it of plan.items) {
    if (!it.activo) continue;
    if (!porCodigo.has(it.activo.codigo)) porCodigo.set(it.activo.codigo, { activo: it.activo, item: it });
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
  // Ordenar familias y subfamilias por cantidad (desc); tarjetas por código.
  const cnt = (sf) => [...sf.values()].reduce((a, arr) => a + arr.length, 0);
  const famOrd = [...fam.entries()].sort((a, b) => cnt(b[1]) - cnt(a[1]));
  for (const [, sf] of famOrd) {
    for (const [, arr] of sf) arr.sort((a, b) => String(a.activo.codigo).localeCompare(String(b.activo.codigo), 'es'));
  }
  const total = porCodigo.size;
  return { fam: famOrd.map(([f, sf]) => [f, [...sf.entries()].sort((a, b) => b[1].length - a[1].length)]), total };
}

const esc = (t) => String(t ?? '');

// Layout (pulgadas, 13.33 x 7.5 — 16:9)
const L = { W: 13.33, H: 7.5, mx: 0.5, cols: 4, gap: 0.3, top: 1.55, cardH: 2.35, rowGap: 0.55, raise: 0.3, photoH: 1.4 };
const cardW = () => (L.W - L.mx * 2 - L.gap * (L.cols - 1)) / L.cols;

function pie(slide, cliente, pptx) {
  slide.addText([
    { text: 'Aquarius Consulting SAC', options: { bold: true } },
    { text: `   ·   Catálogo fotográfico   ·   ${cliente}`, options: {} },
  ], { x: L.mx, y: L.H - 0.42, w: L.W - 1, h: 0.3, fontSize: 8, color: GRIS2, fontFace: 'Arial' });
}

function encabezado(slide, familia, subfam) {
  slide.addShape('rect', { x: L.mx, y: 1.18, w: 0.12, h: 0.42, fill: { color: ACCENT } });
  slide.addText([
    { text: esc(familia).toUpperCase(), options: { bold: true, color: NAVY } },
    { text: `   ›   ${esc(subfam)}`, options: { color: GRIS } },
  ], { x: L.mx + 0.22, y: 1.15, w: L.W - 3, h: 0.48, fontSize: 15, fontFace: 'Arial', valign: 'middle' });
}

function tarjeta(slide, col, row, activo, data) {
  const w = cardW();
  const x = L.mx + col * (w + L.gap);
  const photoTop = L.top + row * (L.cardH + L.rowGap + L.raise);
  const cardTop = photoTop + L.raise;

  // Cuerpo de la tarjeta
  slide.addShape('roundRect', {
    x, y: cardTop, w, h: L.cardH, rectRadius: 0.06,
    fill: { color: 'FFFFFF' }, line: { color: BORDE, width: 1 },
    shadow: { type: 'outer', blur: 6, offset: 3, angle: 90, color: GRIS2, opacity: 0.35 },
  });
  // Foto superpuesta (elevada sobre el borde superior)
  const pw = w - 0.28, px = x + 0.14;
  if (data) {
    slide.addImage({
      data, x: px, y: photoTop, w: pw, h: L.photoH, sizing: { type: 'cover', w: pw, h: L.photoH },
      shadow: { type: 'outer', blur: 5, offset: 2, angle: 90, color: GRIS2, opacity: 0.45 },
    });
  } else {
    slide.addShape('roundRect', { x: px, y: photoTop, w: pw, h: L.photoH, rectRadius: 0.04, fill: { color: 'EEF1F5' }, line: { color: BORDE, width: 1 } });
    slide.addText('Sin vista previa', { x: px, y: photoTop, w: pw, h: L.photoH, align: 'center', valign: 'middle', fontSize: 9, color: GRIS2, fontFace: 'Arial' });
  }
  // Título (código) + descripción
  const ty = photoTop + L.photoH + 0.06;
  slide.addText(esc(activo.codigo).toUpperCase(), { x: x + 0.16, y: ty, w: w - 0.32, h: 0.28, fontSize: 12, bold: true, color: NAVY, fontFace: 'Arial' });
  if (activo.descripcion) {
    slide.addText(esc(activo.descripcion), { x: x + 0.16, y: ty + 0.3, w: w - 0.32, h: 0.5, fontSize: 8.5, color: GRIS, fontFace: 'Arial', valign: 'top', lineSpacingMultiple: 0.9 });
  }
  // Pie de tarjeta (subfamilia/familia) con acento
  slide.addText(esc(activo.subfamilia || activo.familia || '').toUpperCase(), { x: x + 0.16, y: cardTop + L.cardH - 0.34, w: w - 0.32, h: 0.26, fontSize: 8, bold: true, color: ACCENT, fontFace: 'Arial' });
}

/**
 * @param {Object} plan   plan del organizador (con items emparejados)
 * @param {string} cliente nombre del inventario
 * @param {Function} onProgress (hechos, total) => void
 */
export async function generarCatalogoPPT(plan, cliente, onProgress) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Aquarius Consulting';
  pptx.title = `Catálogo fotográfico — ${cliente}`;

  const { fam, total } = agrupar(plan);
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
    { text: `\n${total.toLocaleString('es-PE')} activos con registro fotográfico   ·   ${fecha}`, options: { fontSize: 12, color: 'A7C4E0' } },
  ], { x: L.mx, y: 4.4, w: 11, h: 1, fontFace: 'Arial' });
  cover.addText('ATF-PR-01-FO-03 / Ver.02   ·   Generado con NEXOVA Suite', { x: L.mx, y: 6.9, w: 11, h: 0.3, fontSize: 9, color: GRIS2, fontFace: 'Arial' });

  let hechos = 0;
  for (const [familia, subs] of fam) {
    const nFam = subs.reduce((a, [, arr]) => a + arr.length, 0);
    // Separador de familia
    const sep = pptx.addSlide();
    sep.addShape('rect', { x: 0, y: 0, w: 0.35, h: L.H, fill: { color: NAVY } });
    sep.addShape('rect', { x: L.mx, y: 3.0, w: 0.16, h: 1.0, fill: { color: ACCENT } });
    sep.addText(esc(familia).toUpperCase(), { x: L.mx + 0.35, y: 3.0, w: 11.5, h: 1.0, fontSize: 34, bold: true, color: NAVY, fontFace: 'Arial', valign: 'middle' });
    sep.addText(`${nFam.toLocaleString('es-PE')} activos   ·   ${subs.length} subfamilia(s)`, { x: L.mx + 0.37, y: 4.05, w: 11, h: 0.4, fontSize: 13, color: GRIS, fontFace: 'Arial' });
    pie(sep, cliente, pptx);

    for (const [subfam, activos] of subs) {
      for (let p = 0; p < activos.length; p += 8) {
        const slide = pptx.addSlide();
        encabezado(slide, familia, subfam);
        pie(slide, cliente, pptx);
        const chunk = activos.slice(p, p + 8);
        for (let idx = 0; idx < chunk.length; idx++) {
          const { activo, item } = chunk[idx];
          let data = null;
          try { data = await imgDataURL(item.handle); } catch { /* HEIC u otro: placeholder */ }
          hechos++; onProgress?.(hechos, total);
          tarjeta(slide, idx % L.cols, Math.floor(idx / L.cols), activo, data);
        }
      }
    }
  }

  await pptx.writeFile({ fileName: `Catalogo_Fotografico_${cliente}.pptx` });
}
