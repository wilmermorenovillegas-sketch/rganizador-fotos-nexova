// Pruebas de detección de columnas y lectura de Excel (js/shared/). Ejecutar: node --test "tests/*.test.mjs"
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { detectarColumnas, puntuarHoja } from '../js/shared/deteccion.js';
import { leerTabla, recortarRango } from '../js/shared/excel-io.js';

globalThis.XLSX = createRequire(import.meta.url)('../vendor/xlsx.full.min.js');
const XLSX = globalThis.XLSX;

const CAMPOS = [
  { k: 'codigo', hints: ['barra nueva', 'barnueva', 'codigo de barras', 'codbar', 'codigo activo', 'codigo', 'cod'] },
  { k: 'descripcion', hints: ['descripcion de catalogo', 'desccatalogo', 'descripcion', 'denominacion'] },
  { k: 'cod_ubic', hints: ['codigo de ubicacion', 'codigo ubicacion', 'cod ubicacion'] },
  { k: 'ubicacion', hints: ['descripcion de ubicacion', 'descripcion ubicacion', 'ubicacion'] },
  { k: 'familia', hints: ['descripcion de familia', 'familia'] },
  { k: 'estado', hints: ['estado de conservacion', 'est con', 'estcon', 'estado'] },
  { k: 'color', hints: ['color'] },
  { k: 'area', hints: ['area'] },
  { k: 'sede', hints: ['sede', 'local'] },
  { k: 'observacion', hints: ['observaciones', 'nota'] },
];

test('DIVEMOTOR: el código es BARRA NUEVA, no CODIGO DE UBICACIÓN', () => {
  const h = ['N°', 'BARRA NUEVA', 'LOCAL', 'ÁREA', 'CODIGO DE UBICACIÓN', 'DESCRIPCION DE UBICACIÓN', 'DESCRIPCION DE FAMILIA', 'DESCRIPCION DE CATALOGO', 'ESTADO'];
  const c = detectarColumnas(h, CAMPOS);
  assert.equal(c.codigo, 'BARRA NUEVA');
  assert.equal(c.cod_ubic, 'CODIGO DE UBICACIÓN');
  assert.equal(c.ubicacion, 'DESCRIPCION DE UBICACIÓN');
  assert.equal(c.descripcion, 'DESCRIPCION DE CATALOGO');
  assert.equal(c.familia, 'DESCRIPCION DE FAMILIA');
  assert.equal(c.estado, 'ESTADO');
});

test('encabezados cortos (N°, C, E, A, S) no se asignan por coincidencia parcial', () => {
  const c = detectarColumnas(['N°', 'C', 'E', 'A', 'S', 'CodBar'], CAMPOS);
  assert.equal(c.observacion, '');
  assert.equal(c.color, '');
  assert.equal(c.estado, '');
  assert.equal(c.area, '');
  assert.equal(c.sede, '');
  assert.equal(c.codigo, 'CodBar');
});

test('Est_con se reconoce como estado de conservación', () => {
  assert.equal(detectarColumnas(['BarNueva', 'Est_con'], CAMPOS).estado, 'Est_con');
});

test('una columna nunca se asigna a dos campos', () => {
  const c = detectarColumnas(['Codigo'], CAMPOS);
  const usadas = Object.values(c).filter(Boolean);
  assert.equal(new Set(usadas).size, usadas.length);
});

function libro(hojas) {
  const wb = XLSX.utils.book_new();
  for (const [nombre, aoa] of Object.entries(hojas)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), nombre);
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}

test('recortarRango evita el rango inflado hasta la fila 1.048.576', () => {
  const ws = XLSX.utils.aoa_to_sheet([['A', 'B'], [1, 2]]);
  ws['!ref'] = 'A1:AZ1048576';
  recortarRango(ws);
  assert.equal(ws['!ref'], 'A1:B2');
  const t0 = Date.now();
  const filas = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true });
  assert.equal(filas.length, 2);
  assert.ok(Date.now() - t0 < 200);
});

test('leerTabla elige la hoja de datos aunque la primera sea una portada', () => {
  const buf = libro({
    PORTADA: [['INFORME DE INVENTARIO'], ['Cliente: X']],
    BASE: [['N°', 'BARRA NUEVA', 'DESCRIPCION DE FAMILIA'], [1, 'MQ-1', 'EQUIPOS'], [2, 'MQ-2', 'MUEBLES']],
  });
  const t = leerTabla(buf, (h) => puntuarHoja(h, CAMPOS));
  assert.equal(t.hoja, 'BASE');
  assert.equal(t.registros.length, 2);
  assert.equal(t.registros[1]['BARRA NUEVA'], 'MQ-2');
});

test('leerTabla salta filas vacías y desduplica encabezados', () => {
  const buf = libro({ H: [['Marca', 'Marca', '', 'Nota'], ['HP', 'Dell', 1, 2], [], ['LG', 'Sony', 3, 4]] });
  const t = leerTabla(buf);
  assert.deepEqual(t.headers, ['Marca', 'Marca(2)', 'Col3', 'Nota']);
  assert.equal(t.registros.length, 2);
});
