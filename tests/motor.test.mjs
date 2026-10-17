// Pruebas del motor de emparejamiento (js/motor.js). Ejecutar: node --test "tests/*.test.mjs"
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  celdaATexto, normalizar, claveFlexible, sanearCarpeta, partirNombre, esFoto,
  detectarFilaEncabezado, IndiceCodigos, construirPlan, calcularMetricas, CARPETA_HUERFANAS,
} from '../js/motor.js';

const activos = (...cods) => new Map(cods.map((c) => [normalizar(c), { codigo: c, familia: 'F', subfamilia: 'S', descripcion: '' }]));
const buscar = (cods, nombre) => new IndiceCodigos(activos(...cods).values()).buscar(nombre);

test('celdaATexto convierte números de Excel sin decimales espurios', () => {
  assert.equal(celdaATexto(1001), '1001');
  assert.equal(celdaATexto(null), '');
  assert.equal(celdaATexto('  MQ-1  '), 'MQ-1');
  assert.equal(celdaATexto(true), 'SI');
});

test('normalizar quita tildes y espacios, conserva guiones', () => {
  assert.equal(normalizar('  cámara   de  vídeo-1 '), 'CAMARA DE VIDEO-1');
});

test('claveFlexible iguala EC-001, "ec 001" y EC1', () => {
  assert.equal(claveFlexible('EC-001'), claveFlexible('ec 001'));
  assert.equal(claveFlexible('EC-001'), claveFlexible('EC1'));
});

test('sufijos de foto múltiple: _1, _2, -3, .4, (5)', () => {
  for (const n of ['MQ-00100_1', 'MQ-00100_2', 'MQ-00100-3', 'MQ-00100.4', 'MQ-00100 (5)']) {
    const r = buscar(['MQ-00100'], n);
    assert.equal(r.activo?.codigo, 'MQ-00100', n);
    assert.match(r.metodo, /exacto/);
  }
});

test('caso crítico: MOB-10-2 es la foto 2 de MOB-10, no el activo MOB-102', () => {
  assert.equal(buscar(['MOB-10', 'MOB-102'], 'MOB-10-2').activo.codigo, 'MOB-10');
  assert.equal(buscar(['MOB-10', 'MOB-102'], 'MOB-102').activo.codigo, 'MOB-102');
});

test('no recorta de más: MQ-00100 no se convierte en MQ', () => {
  assert.equal(IndiceCodigos.recortarSufijo('MQ-00100'), '');
  assert.equal(IndiceCodigos.recortarSufijo('EC001'), '');
});

test('coincidencia flexible con ceros a la izquierda y sin guiones', () => {
  const r = buscar(['EC-001'], 'ec1');
  assert.equal(r.activo.codigo, 'EC-001');
  assert.match(r.metodo, /normalizado/);
});

test('claves flexibles ambiguas nunca se adivinan', () => {
  // EC-01 y EC-1 colapsan a la misma clave flexible "EC1"
  const r = buscar(['EC-01', 'EC-1'], 'ec 0001');
  assert.equal(r.activo, null);
});

test('foto sin código conocido queda SIN MATCH', () => {
  assert.equal(buscar(['MQ-1'], 'ZZ-999').metodo, 'SIN MATCH');
});

test('sanearCarpeta: caracteres inválidos y punto final de Windows', () => {
  assert.equal(sanearCarpeta('Muebles/Enseres: oficina.'), 'Muebles_Enseres_ oficina');
  assert.equal(sanearCarpeta('   '), 'SIN DATO');
});

test('partirNombre y esFoto', () => {
  assert.deepEqual(partirNombre('MQ-1_2.JPEG'), { base: 'MQ-1_2', ext: 'jpeg' });
  assert.equal(esFoto('a.heic'), true);
  assert.equal(esFoto('informe.pdf'), false);
});

test('detectarFilaEncabezado salta logo/título del Excel del cliente', () => {
  const filas = [['INVENTARIO FISICO'], [], ['N°', 'BARRA NUEVA', 'FAMILIA', 'MARCA'], [1, 'MQ-1', 'EQUIPOS', 'HP']];
  assert.equal(detectarFilaEncabezado(filas), 2);
});

test('construirPlan: homónimos no se pisan y las huérfanas van a la carpeta de sin clasificar', () => {
  const fotos = [
    { nombre: 'MQ-1_1.jpg', ruta: 'a/MQ-1_1.jpg' },
    { nombre: 'MQ-1_1.jpg', ruta: 'b/MQ-1_1.jpg' },
    { nombre: 'XX.jpg', ruta: 'XX.jpg' },
  ];
  const plan = construirPlan(fotos, activos('MQ-1'), [], { niveles: 1, moverHuerfanas: true });
  assert.deepEqual(plan.items.map((i) => i.nombreDestino), ['MQ-1_1.jpg', 'MQ-1_1__2.jpg', 'XX.jpg']);
  assert.deepEqual(plan.items[2].carpetas, [CARPETA_HUERFANAS]);
  const m = calcularMetricas(plan);
  assert.equal(m.emparejadas, 2);
  assert.equal(m.huerfanas, 1);
  assert.equal(m.cobertura, 100);
});
