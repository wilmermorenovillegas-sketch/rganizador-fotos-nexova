/** Criterio de aceptacion: los mismos 14 casos que valido el motor en Python. */
import { IndiceCodigos, cargarActivos, construirPlan, calcularMetricas, agruparPorFamilia } from './js/motor.js';

const filasInv = [
  { 'Código Activo': 'EC-001',  'Cód. Catálogo': 'CAT-100', 'Descripción': 'Laptop Dell Latitude' },
  { 'Código Activo': 'EC-002',  'Cód. Catálogo': 'CAT-100', 'Descripción': 'Laptop HP ProBook' },
  { 'Código Activo': 'MQ-0050', 'Cód. Catálogo': 'CAT-200', 'Descripción': 'Compresora industrial' },
  { 'Código Activo': 'MOB-10',  'Cód. Catálogo': 'CAT-300', 'Descripción': 'Escritorio gerencial' },
  { 'Código Activo': 'MOB-102', 'Cód. Catálogo': 'CAT-300', 'Descripción': 'Silla ergonómica' },
  { 'Código Activo': 'TR2024',  'Cód. Catálogo': 'CAT-400', 'Descripción': 'Camioneta Hilux' },
  { 'Código Activo': 1001,      'Cód. Catálogo': 'CAT-100', 'Descripción': 'Monitor LG 27' },  // numérico
  { 'Código Activo': 'AB-01',   'Cód. Catálogo': 'CAT-300', 'Descripción': 'Archivador A' },
  { 'Código Activo': 'AB-1',    'Cód. Catálogo': 'CAT-300', 'Descripción': 'Archivador B' },
  { 'Código Activo': 'EC-999',  'Cód. Catálogo': 'CAT-100', 'Descripción': 'Servidor (sin foto)' },
];
const filasCat = [
  { 'Código': 'CAT-100', 'Familia': 'EQUIPOS DE COMPUTO',     'Subfamilia': 'Computo Personal' },
  { 'Código': 'CAT-200', 'Familia': 'MAQUINARIAS Y EQUIPOS',  'Subfamilia': 'Equipos de Planta' },
  { 'Código': 'CAT-300', 'Familia': 'MUEBLES Y ENSERES',      'Subfamilia': 'Mobiliario de Oficina' },
  { 'Código': 'CAT-400', 'Familia': 'UNIDADES DE TRANSPORTE', 'Subfamilia': 'Vehiculos Livianos' },
];
const cfg = {
  colCodigo: 'Código Activo', colEnlaceInv: 'Cód. Catálogo', colDescripcion: 'Descripción',
  usarCatalogo: true, colEnlaceCat: 'Código', colFamilia: 'Familia', colSubfamilia: 'Subfamilia',
  niveles: 3, moverHuerfanas: true,
};

const ESPERADO = [
  ['EC-001.jpg',              'EC-001',  'exacto'],
  ['EC-001-1.jpg',            'EC-001',  'exacto+sufijo'],
  ['EC-001-2.jpg',            'EC-001',  'exacto+sufijo'],
  ['EC-001 (3).jpg',          'EC-001',  'exacto+sufijo'],
  ['ec002.JPG',               'EC-002',  'normalizado'],
  ['MQ-0050_2.png',           'MQ-0050', 'exacto+sufijo'],
  ['MOB-10-2.jpg',            'MOB-10',  'exacto+sufijo'],   // <-- EL CASO CRITICO
  ['MOB-102.jpg',             'MOB-102', 'exacto'],
  ['TR2024.jpg',              'TR2024',  'exacto'],
  ['TR2024-1.jpg',            'TR2024',  'exacto+sufijo'],
  ['1001.jpg',                '1001',    'exacto'],
  ['AB-01.jpg',               'AB-01',   'exacto'],
  ['IMG_20240115_103245.jpg', null,      'SIN MATCH'],
  ['foto sin codigo.jpg',     null,      'SIN MATCH'],
];

const { activos, alertas } = cargarActivos(filasInv, filasCat, cfg);
const fotos = ESPERADO.map(([n]) => ({ nombre: n, ruta: n, handle: null }));
const plan = construirPlan(fotos, activos, alertas, cfg);

let fallos = 0;
console.log('%s %s %s %s', 'ARCHIVO'.padEnd(26), 'ESPERADO'.padEnd(10), 'OBTENIDO'.padEnd(10), 'MÉTODO');
console.log('-'.repeat(78));
plan.items.forEach((it, i) => {
  const [archivo, codEsp, metEsp] = ESPERADO[i];
  const codObt = it.activo ? it.activo.codigo : null;
  const ok = String(codObt) === String(codEsp) && it.metodo === metEsp;
  if (!ok) fallos++;
  console.log('%s %s %s %s %s', archivo.padEnd(26), String(codEsp).padEnd(10),
              String(codObt).padEnd(10), it.metodo.padEnd(16), ok ? '✓' : '✗ FALLO');
});

const m = calcularMetricas(plan);
console.log('\nMétricas: %d emparejadas, %d huérfanas, cobertura %s%%, sin familia %d',
            m.emparejadas, m.huerfanas, m.cobertura.toFixed(1), m.sinFamilia);
console.log('Alertas:', plan.alertas.map(a => a.join(': ')).join(' | '));
console.log('\nPor familia:');
for (const f of agruparPorFamilia(plan)) {
  console.log('  %s — %d activos, %d con foto (%s%%), %d fotos',
    f.familia.padEnd(24), f.activos, f.conFoto, f.cobertura.toFixed(0), f.fotos);
}

const okGlobal = fallos === 0 && m.emparejadas === 12 && m.huerfanas === 2
                 && Math.round(m.cobertura) === 80 && m.sinFamilia === 0;
console.log('\n%s', okGlobal ? '✅ 14/14 CASOS OK — motor validado' : `❌ ${fallos} FALLOS`);
process.exit(okGlobal ? 0 : 1);
