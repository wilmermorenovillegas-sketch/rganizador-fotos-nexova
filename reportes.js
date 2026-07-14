/**
 * ============================================================================
 * reportes.js — GRAFICOS · EXCEL · INFORME PDF
 * ============================================================================
 *
 * Aqui es donde la herramienta deja de ser una utilidad interna y se convierte
 * en un ENTREGABLE FACTURABLE: el informe PDF de cobertura fotografica es un
 * documento de control de calidad que se le entrega al cliente.
 * ============================================================================
 */

import { calcularMetricas, agruparPorFamilia, normalizar } from './motor.js';

// Paleta institucional NEXOVA
const TEAL = '#0F766E';
const TEAL_CLARO = '#0D9488';
const SLATE = '#1E293B';
const GRIS = '#94A3B8';
const AMBAR = '#D97706';
const ROJO = '#B91C1C';

const rgb = (hex) => [
  parseInt(hex.slice(1, 3), 16),
  parseInt(hex.slice(3, 5), 16),
  parseInt(hex.slice(5, 7), 16),
];

const num = (n) => n.toLocaleString('es-PE');

// ---------------------------------------------------------------------------
// GRAFICOS
// ---------------------------------------------------------------------------

/** Pinta el canvas de blanco: sin esto, el PNG sale transparente y el PDF lo ennegrece. */
const fondoBlanco = {
  id: 'fondoBlanco',
  beforeDraw(c) {
    const ctx = c.ctx;
    ctx.save();
    ctx.globalCompositeOperation = 'destination-over';
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, c.width, c.height);
    ctx.restore();
  },
};

const BASE = {
  responsive: false,
  animation: false,           // sin animacion: el canvas queda listo para exportar a PDF
  plugins: {
    legend: { labels: { font: { family: 'DM Sans, sans-serif', size: 12 }, color: SLATE } },
  },
};

const registro = new Map();   // reusar canvas entre re-renders sin fugas

function nuevoChart(canvas, config) {
  registro.get(canvas)?.destroy();
  const chart = new Chart(canvas.getContext('2d'), config);
  registro.set(canvas, chart);
  return chart;
}

/** 1 · Dona: fotos emparejadas vs fotos sin codigo reconocible. */
export function graficoEstado(canvas, m) {
  return nuevoChart(canvas, {
    type: 'doughnut',
    plugins: [fondoBlanco],
    data: {
      labels: ['Fotos emparejadas', 'Fotos sin código'],
      datasets: [{
        data: [m.emparejadas, m.huerfanas],
        backgroundColor: [TEAL, GRIS],
        borderWidth: 2,
        borderColor: '#FFFFFF',
      }],
    },
    options: {
      ...BASE,
      cutout: '62%',
      plugins: { ...BASE.plugins, legend: { ...BASE.plugins.legend, position: 'bottom' } },
    },
  });
}

/** 2 · Cobertura fotografica por familia. El KPI que le importa al cliente. */
export function graficoCobertura(canvas, familias) {
  const datos = familias.slice(0, 12);
  return nuevoChart(canvas, {
    type: 'bar',
    plugins: [fondoBlanco],
    data: {
      labels: datos.map((f) => f.familia),
      datasets: [{
        label: '% de activos con foto',
        data: datos.map((f) => +f.cobertura.toFixed(1)),
        // Semaforo: rojo bajo 60%, ambar bajo 90%, teal si esta sano.
        backgroundColor: datos.map((f) =>
          f.cobertura < 60 ? ROJO : f.cobertura < 90 ? AMBAR : TEAL),
        borderRadius: 3,
      }],
    },
    options: {
      ...BASE,
      indexAxis: 'y',
      scales: {
        x: { max: 100, ticks: { callback: (v) => v + '%', color: SLATE },
             grid: { color: '#E2E8F0' } },
        y: { ticks: { color: SLATE, font: { size: 11 } }, grid: { display: false } },
      },
      plugins: { ...BASE.plugins, legend: { display: false } },
    },
  });
}

/** 3 · Activos con foto vs sin foto, por familia (barras apiladas). */
export function graficoActivos(canvas, familias) {
  const datos = familias.slice(0, 12);
  return nuevoChart(canvas, {
    type: 'bar',
    plugins: [fondoBlanco],
    data: {
      labels: datos.map((f) => f.familia),
      datasets: [
        { label: 'Con foto', data: datos.map((f) => f.conFoto),
          backgroundColor: TEAL, borderRadius: 3 },
        { label: 'Sin foto', data: datos.map((f) => f.sinFoto),
          backgroundColor: '#CBD5E1', borderRadius: 3 },
      ],
    },
    options: {
      ...BASE,
      scales: {
        x: { stacked: true, ticks: { color: SLATE, font: { size: 10 }, maxRotation: 40 },
             grid: { display: false } },
        y: { stacked: true, ticks: { color: SLATE }, grid: { color: '#E2E8F0' } },
      },
      plugins: { ...BASE.plugins, legend: { ...BASE.plugins.legend, position: 'bottom' } },
    },
  });
}

/** Exporta un canvas de Chart.js a PNG para incrustarlo en el PDF. */
async function aImagen(canvas) {
  await new Promise((r) => requestAnimationFrame(r));
  return canvas.toDataURL('image/png', 1.0);
}

// ---------------------------------------------------------------------------
// EXCEL DE CONCILIACION (ExcelJS — 6 hojas)
// ---------------------------------------------------------------------------

function encabezar(ws, columnas) {
  ws.columns = columnas;
  const fila = ws.getRow(1);
  fila.eachCell((c) => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    c.border = { bottom: { style: 'medium', color: { argb: 'FF0F766E' } } };
    c.alignment = { vertical: 'middle' };
  });
  fila.height = 20;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columnas.length } };
}

export async function generarExcel(plan, cfg, simulacion) {
  const m = calcularMetricas(plan);
  const familias = agruparPorFamilia(plan);
  const wb = new ExcelJS.Workbook();
  wb.creator = 'NEXOVA · AssetControl';
  wb.created = new Date();

  // --- RESUMEN
  const ws = wb.addWorksheet('RESUMEN');
  ws.getColumn(1).width = 34;
  ws.getColumn(2).width = 20;
  ws.mergeCells('A1:B1');
  ws.getCell('A1').value = 'REPORTE DE CONCILIACIÓN FOTOGRÁFICA';
  ws.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FF1E293B' } };
  ws.getCell('A2').value = simulacion
    ? 'SIMULACIÓN — no se modificó ningún archivo'
    : `EJECUTADO — modo ${cfg.modo.toUpperCase()}`;
  ws.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF0F766E' } };
  ws.getCell('A3').value = `Cliente: ${cfg.cliente || '—'}   ·   ${new Date().toLocaleString('es-PE')}`;
  ws.getCell('A3').font = { size: 10, color: { argb: 'FF64748B' } };

  const metricas = [
    ['Activos en el inventario', m.totalActivos],
    ['Fotos encontradas', m.totalFotos],
    ['Fotos emparejadas', m.emparejadas],
    ['Fotos sin código', m.huerfanas],
    ['Activos CON foto', m.activosConFoto],
    ['Activos SIN foto', m.activosSinFoto],
    ['Cobertura fotográfica', `${m.cobertura.toFixed(1)}%`],
    ['Activos sin familia asignada', m.sinFamilia],
    ['Errores de archivo', m.errores],
  ];
  metricas.forEach(([k, v], i) => {
    const f = ws.getRow(5 + i);
    f.getCell(1).value = k;
    f.getCell(2).value = v;
    f.getCell(1).font = { bold: true, color: { argb: 'FF1E293B' } };
    f.getCell(2).font = { bold: true, size: 12, color: { argb: 'FF0F766E' } };
    f.getCell(2).alignment = { horizontal: 'right' };
    [1, 2].forEach((c) => {
      f.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    });
  });

  // --- DETALLE
  const wsD = wb.addWorksheet('DETALLE');
  encabezar(wsD, [
    { header: 'Archivo', key: 'a', width: 30 },
    { header: 'Código', key: 'c', width: 18 },
    { header: 'Descripción', key: 'd', width: 42 },
    { header: 'Familia', key: 'f', width: 26 },
    { header: 'Subfamilia', key: 's', width: 26 },
    { header: 'Método de match', key: 'm', width: 18 },
    { header: 'Estado', key: 'e', width: 16 },
    { header: 'Ruta destino', key: 'r', width: 60 },
  ]);
  for (const it of plan.items) {
    wsD.addRow({
      a: it.nombre,
      c: it.activo?.codigo ?? '',
      d: it.activo?.descripcion ?? '',
      f: it.activo?.familia ?? '',
      s: it.activo?.subfamilia ?? '',
      m: it.metodo,
      e: it.estado,
      r: it.carpetas ? `${it.carpetas.join('/')}/${it.nombreDestino}` : '',
    });
  }

  // --- FOTOS SIN CODIGO
  const wsH = wb.addWorksheet('FOTOS SIN CODIGO');
  encabezar(wsH, [
    { header: 'Archivo', key: 'a', width: 40 },
    { header: 'Ruta de origen', key: 'r', width: 70 },
  ]);
  for (const it of m.listaHuerfanas) wsH.addRow({ a: it.nombre, r: it.ruta });

  // --- ACTIVOS SIN FOTO  (entregable de control de calidad para el cliente)
  const wsA = wb.addWorksheet('ACTIVOS SIN FOTO');
  encabezar(wsA, [
    { header: 'Código', key: 'c', width: 20 },
    { header: 'Descripción', key: 'd', width: 50 },
    { header: 'Familia', key: 'f', width: 28 },
    { header: 'Subfamilia', key: 's', width: 28 },
  ]);
  for (const a of m.listaActivosSinFoto) {
    wsA.addRow({ c: a.codigo, d: a.descripcion, f: a.familia, s: a.subfamilia });
  }

  // --- POR FAMILIA
  const wsF = wb.addWorksheet('POR FAMILIA');
  encabezar(wsF, [
    { header: 'Familia', key: 'f', width: 30 },
    { header: 'Subfamilia', key: 's', width: 30 },
    { header: 'Activos', key: 'a', width: 12 },
    { header: 'Con foto', key: 'c', width: 12 },
    { header: 'Sin foto', key: 'x', width: 12 },
    { header: 'Cobertura', key: 'p', width: 12 },
    { header: 'Fotos', key: 'n', width: 10 },
  ]);
  for (const fam of familias) {
    for (const sub of fam.subfamilias) {
      wsF.addRow({
        f: fam.familia, s: sub.subfamilia, a: sub.activos, c: sub.conFoto,
        x: sub.sinFoto, p: `${sub.cobertura.toFixed(1)}%`, n: sub.fotos,
      });
    }
  }

  // --- ALERTAS
  const wsAl = wb.addWorksheet('ALERTAS');
  encabezar(wsAl, [
    { header: 'Tipo de alerta', key: 't', width: 45 },
    { header: 'Dato', key: 'd', width: 70 },
  ]);
  for (const [t, d] of plan.alertas) wsAl.addRow({ t, d });

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

// ---------------------------------------------------------------------------
// INFORME PDF (jsPDF + autoTable)
// ---------------------------------------------------------------------------

const A4 = { ancho: 210, alto: 297, margen: 16 };

function pie(doc, cliente) {
  const total = doc.internal.getNumberOfPages();
  for (let p = 1; p <= total; p++) {
    doc.setPage(p);
    if (p === 1) continue;                       // la portada no lleva pie
    doc.setDrawColor(...rgb('#E2E8F0'));
    doc.setLineWidth(0.3);
    doc.line(A4.margen, A4.alto - 14, A4.ancho - A4.margen, A4.alto - 14);
    doc.setFont('helvetica', 'normal').setFontSize(7.5).setTextColor(...rgb(GRIS));
    doc.text('NEXOVA Software Empresarial · nexova.pe · +51 949 287 897',
             A4.margen, A4.alto - 9);
    doc.text(`Página ${p} de ${total}`, A4.ancho - A4.margen, A4.alto - 9, { align: 'right' });
    if (cliente) {
      doc.text(cliente, A4.ancho / 2, A4.alto - 9, { align: 'center' });
    }
  }
}

function tituloSeccion(doc, y, texto) {
  doc.setFillColor(...rgb(TEAL));
  doc.rect(A4.margen, y - 4, 3, 6, 'F');                  // barrita teal de acento
  doc.setFont('helvetica', 'bold').setFontSize(13).setTextColor(...rgb(SLATE));
  doc.text(texto, A4.margen + 7, y + 1);
  return y + 10;
}

/** Tarjeta de KPI: numero grande sobre fondo perla. */
function tarjetaKPI(doc, x, y, ancho, etiqueta, valor, color = SLATE) {
  doc.setFillColor(...rgb('#F1F5F9'));
  doc.roundedRect(x, y, ancho, 20, 1.5, 1.5, 'F');
  doc.setFont('helvetica', 'normal').setFontSize(7.5).setTextColor(...rgb('#64748B'));
  doc.text(etiqueta.toUpperCase(), x + 4, y + 6);
  doc.setFont('helvetica', 'bold').setFontSize(15).setTextColor(...rgb(color));
  doc.text(String(valor), x + 4, y + 15);
}

export async function generarPDF(plan, cfg, canvases, simulacion) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  const m = calcularMetricas(plan);
  const familias = agruparPorFamilia(plan);
  const cliente = cfg.cliente || 'Cliente no especificado';
  const hoy = new Date();

  // ===================== PORTADA =====================
  doc.setFillColor(...rgb(SLATE));
  doc.rect(0, 0, A4.ancho, 120, 'F');
  doc.setFillColor(...rgb(TEAL));
  doc.rect(0, 118, A4.ancho, 2.5, 'F');

  // Isotipo hexagonal + wordmark
  doc.setFillColor(...rgb(TEAL));
  const cx = A4.margen + 6, cy = 32, r = 6;
  const hex = Array.from({ length: 6 }, (_, i) => {
    const ang = (Math.PI / 3) * i - Math.PI / 6;
    return [cx + r * Math.cos(ang), cy + r * Math.sin(ang)];
  });
  doc.lines(
    hex.slice(1).concat([hex[0]]).map(([x, y], i) => {
      const [px, py] = i === 0 ? hex[0] : hex[i];
      return [x - px, y - py];
    }),
    hex[0][0], hex[0][1], [1, 1], 'F'
  );
  doc.setFont('helvetica', 'bold').setFontSize(22).setTextColor(255, 255, 255);
  doc.text('NEXOVA', cx + 12, cy + 3);
  doc.setFont('helvetica', 'normal').setFontSize(7.5).setTextColor(...rgb(TEAL_CLARO));
  doc.text('S O F T W A R E   E M P R E S A R I A L', cx + 12.5, cy + 9);

  doc.setFont('helvetica', 'bold').setFontSize(26).setTextColor(255, 255, 255);
  doc.text('INFORME DE COBERTURA', A4.margen, 74);
  doc.text('FOTOGRÁFICA', A4.margen, 86);
  doc.setFont('helvetica', 'normal').setFontSize(10).setTextColor(...rgb('#94A3B8'));
  doc.text('Inventario de activos fijos · Control de calidad del registro fotográfico',
           A4.margen, 97);

  doc.setFont('helvetica', 'bold').setFontSize(11).setTextColor(...rgb(SLATE));
  doc.text('CLIENTE', A4.margen, 140);
  doc.setFont('helvetica', 'normal').setFontSize(14);
  doc.text(cliente, A4.margen, 148);

  const codigo = `NXV-ICF-${hoy.getFullYear()}${String(hoy.getMonth() + 1).padStart(2, '0')}${String(hoy.getDate()).padStart(2, '0')}`;
  doc.setFont('helvetica', 'bold').setFontSize(9).setTextColor(...rgb('#64748B'));
  doc.text('INFORME N°', A4.margen, 164);
  doc.text('FECHA DE EMISIÓN', A4.margen + 60, 164);
  doc.text('ESTADO', A4.margen + 130, 164);
  doc.setFont('helvetica', 'normal').setFontSize(10).setTextColor(...rgb(SLATE));
  doc.text(codigo, A4.margen, 170);
  doc.text(hoy.toLocaleDateString('es-PE'), A4.margen + 60, 170);
  doc.setTextColor(...rgb(simulacion ? AMBAR : TEAL));
  doc.text(simulacion ? 'SIMULACIÓN' : 'EJECUTADO', A4.margen + 130, 170);

  // Titular: la cobertura, en grande. Es el numero que el cliente quiere ver.
  const colorCob = m.cobertura < 60 ? ROJO : m.cobertura < 90 ? AMBAR : TEAL;
  doc.setFillColor(...rgb('#F1F5F9'));
  doc.roundedRect(A4.margen, 188, A4.ancho - A4.margen * 2, 40, 2, 2, 'F');
  doc.setFillColor(...rgb(colorCob));
  doc.rect(A4.margen, 188, 3, 40, 'F');
  doc.setFont('helvetica', 'normal').setFontSize(9).setTextColor(...rgb('#64748B'));
  doc.text('COBERTURA FOTOGRÁFICA DEL INVENTARIO', A4.margen + 10, 200);
  doc.setFont('helvetica', 'bold').setFontSize(34).setTextColor(...rgb(colorCob));
  doc.text(`${m.cobertura.toFixed(1)}%`, A4.margen + 10, 216);
  doc.setFont('helvetica', 'normal').setFontSize(9.5).setTextColor(...rgb(SLATE));
  doc.text(`${num(m.activosConFoto)} de ${num(m.totalActivos)} activos cuentan con registro fotográfico.`,
           A4.margen + 55, 211);
  doc.text(`Quedan ${num(m.activosSinFoto)} activos pendientes de fotografiar.`,
           A4.margen + 55, 217);

  doc.setFont('helvetica', 'normal').setFontSize(7.5).setTextColor(...rgb(GRIS));
  doc.text('Documento generado automáticamente por NEXOVA AssetControl · Organizador Fotográfico de Inventarios',
           A4.margen, 272);
  doc.text('Los archivos fotográficos fueron procesados localmente. Ninguna imagen fue transmitida a servidores externos.',
           A4.margen, 277);

  // ===================== 1 · RESUMEN EJECUTIVO =====================
  doc.addPage();
  let y = tituloSeccion(doc, 24, '1. Resumen ejecutivo');

  const anchoT = (A4.ancho - A4.margen * 2 - 6) / 3;
  tarjetaKPI(doc, A4.margen, y, anchoT, 'Activos en inventario', num(m.totalActivos));
  tarjetaKPI(doc, A4.margen + anchoT + 3, y, anchoT, 'Fotos procesadas', num(m.totalFotos));
  tarjetaKPI(doc, A4.margen + (anchoT + 3) * 2, y, anchoT, 'Cobertura', `${m.cobertura.toFixed(1)}%`, colorCob);
  y += 24;
  tarjetaKPI(doc, A4.margen, y, anchoT, 'Activos con foto', num(m.activosConFoto), TEAL);
  tarjetaKPI(doc, A4.margen + anchoT + 3, y, anchoT, 'Activos sin foto', num(m.activosSinFoto), m.activosSinFoto ? AMBAR : TEAL);
  tarjetaKPI(doc, A4.margen + (anchoT + 3) * 2, y, anchoT, 'Fotos sin código', num(m.huerfanas), m.huerfanas ? AMBAR : TEAL);
  y += 30;

  if (canvases.estado) {
    doc.addImage(await aImagen(canvases.estado), 'PNG', A4.margen, y, 78, 58);
  }
  doc.setFont('helvetica', 'bold').setFontSize(9.5).setTextColor(...rgb(SLATE));
  doc.text('Lectura del resultado', A4.margen + 88, y + 8);
  doc.setFont('helvetica', 'normal').setFontSize(8.5).setTextColor(...rgb('#475569'));
  const parrafo = [
    `De las ${num(m.totalFotos)} fotografías procesadas, ${num(m.emparejadas)} fueron asociadas`,
    `correctamente a un activo del inventario y ${num(m.huerfanas)} no corresponden a`,
    `ningún código registrado.`,
    '',
    m.activosSinFoto > 0
      ? `Quedan ${num(m.activosSinFoto)} activos sin registro fotográfico. El detalle`
      : 'Todos los activos del inventario cuentan con registro fotográfico.',
    m.activosSinFoto > 0 ? 'completo se lista en la sección 3 de este informe.' : '',
  ];
  parrafo.forEach((t, i) => t && doc.text(t, A4.margen + 88, y + 16 + i * 5));
  y += 66;

  if (m.sinFamilia > 0) {
    doc.setFillColor(...rgb('#FEF3C7'));
    doc.roundedRect(A4.margen, y, A4.ancho - A4.margen * 2, 16, 1.5, 1.5, 'F');
    doc.setFillColor(...rgb(AMBAR));
    doc.rect(A4.margen, y, 2.5, 16, 'F');
    doc.setFont('helvetica', 'bold').setFontSize(8.5).setTextColor(...rgb('#92400E'));
    doc.text(`OBSERVACIÓN · ${num(m.sinFamilia)} activos (${m.pctSinFamilia.toFixed(0)}%) no tienen familia asignada en el catálogo.`,
             A4.margen + 7, y + 7);
    doc.setFont('helvetica', 'normal').setFontSize(8);
    doc.text('Revisar la correspondencia entre el código de catálogo del inventario y la maestra de familias.',
             A4.margen + 7, y + 12.5);
  }

  // ===================== 2 · COBERTURA POR FAMILIA =====================
  doc.addPage();
  y = tituloSeccion(doc, 24, '2. Cobertura por familia');

  if (canvases.cobertura) {
    doc.addImage(await aImagen(canvases.cobertura), 'PNG', A4.margen, y, A4.ancho - A4.margen * 2, 62);
    y += 70;
  }

  doc.autoTable({
    startY: y,
    margin: { left: A4.margen, right: A4.margen },
    head: [['Familia', 'Activos', 'Con foto', 'Sin foto', 'Cobertura', 'Fotos']],
    body: familias.map((f) => [
      f.familia, num(f.activos), num(f.conFoto), num(f.sinFoto),
      `${f.cobertura.toFixed(1)}%`, num(f.fotos),
    ]),
    theme: 'grid',
    styles: { fontSize: 8.5, cellPadding: 2.2, lineColor: rgb('#E2E8F0'), lineWidth: 0.1 },
    headStyles: { fillColor: rgb(SLATE), textColor: 255, fontStyle: 'bold', fontSize: 8.5 },
    alternateRowStyles: { fillColor: rgb('#F8FAFB') },
    columnStyles: {
      1: { halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' },
      4: { halign: 'right', fontStyle: 'bold' }, 5: { halign: 'right' },
    },
    didParseCell: (d) => {
      if (d.section === 'body' && d.column.index === 4) {
        const v = parseFloat(d.cell.raw);
        d.cell.styles.textColor = v < 60 ? rgb(ROJO) : v < 90 ? rgb(AMBAR) : rgb(TEAL);
      }
    },
  });

  // ===================== 3 · ACTIVOS SIN REGISTRO =====================
  if (m.activosSinFoto > 0) {
    doc.addPage();
    y = tituloSeccion(doc, 24, '3. Activos sin registro fotográfico');
    doc.setFont('helvetica', 'normal').setFontSize(8.5).setTextColor(...rgb('#475569'));
    doc.text(`${num(m.activosSinFoto)} activos requieren ser fotografiados para cerrar el levantamiento.`,
             A4.margen, y);
    y += 6;

    const LIMITE = 400;   // un PDF de 5,000 filas no lo lee nadie: eso va en el Excel
    const lista = m.listaActivosSinFoto.slice(0, LIMITE);
    doc.autoTable({
      startY: y,
      margin: { left: A4.margen, right: A4.margen },
      head: [['Código', 'Descripción', 'Familia', 'Subfamilia']],
      body: lista.map((a) => [a.codigo, a.descripcion, a.familia, a.subfamilia]),
      theme: 'striped',
      styles: { fontSize: 7.5, cellPadding: 1.8 },
      headStyles: { fillColor: rgb(SLATE), textColor: 255, fontSize: 8 },
      alternateRowStyles: { fillColor: rgb('#F8FAFB') },
      columnStyles: { 0: { fontStyle: 'bold', cellWidth: 26 } },
    });
    if (m.activosSinFoto > LIMITE) {
      doc.setFont('helvetica', 'italic').setFontSize(8).setTextColor(...rgb(GRIS));
      doc.text(`Se muestran los primeros ${LIMITE} de ${num(m.activosSinFoto)}. El listado completo está en el Excel de conciliación, hoja "ACTIVOS SIN FOTO".`,
               A4.margen, doc.lastAutoTable.finalY + 6, { maxWidth: A4.ancho - A4.margen * 2 });
    }
  }

  // ===================== 4 · ALERTAS =====================
  if (plan.alertas.length) {
    doc.addPage();
    y = tituloSeccion(doc, 24, '4. Alertas y observaciones técnicas');
    doc.autoTable({
      startY: y,
      margin: { left: A4.margen, right: A4.margen },
      head: [['Tipo', 'Detalle']],
      body: plan.alertas.slice(0, 300),
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 2 },
      headStyles: { fillColor: rgb(SLATE), textColor: 255 },
      columnStyles: { 0: { cellWidth: 62, fontStyle: 'bold' } },
    });
  }

  pie(doc, cliente);
  return doc.output('blob');
}

export { calcularMetricas, agruparPorFamilia };
