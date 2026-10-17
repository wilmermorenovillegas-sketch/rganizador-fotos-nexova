// Pruebas del maestro de marcas/modelos (js/shared/maestro-nube.js). Ejecutar: node --test "tests/*.test.mjs"
import test from 'node:test';
import assert from 'node:assert/strict';
import { construirMaestro } from '../js/shared/maestro-nube.js';

test('agrupa modelos por marca sin importar mayúsculas, tildes ni espacios', () => {
  const m = construirMaestro([
    { marca: 'HP', modelo: 'ProDesk 400' }, { marca: ' hp ', modelo: 'PRODESK  400' },
    { marca: 'HP', modelo: 'EliteBook 840' }, { marca: 'Epson', modelo: null },
  ]);
  assert.equal(m.size, 2);
  assert.equal(m.get('HP').modelos.size, 2);
  assert.equal(m.get('EPSON').modelos.size, 0);
});

test('el maestro local se suma al de la nube sin duplicar', () => {
  const nube = construirMaestro([{ marca: 'DELL', modelo: null }]);
  const todo = construirMaestro([{ marca: 'Dell', modelo: 'Latitude 5420' }, { marca: 'Miray', modelo: '' }], nube);
  assert.equal(todo.size, 2);
  assert.equal(todo.get('DELL').nombre, 'DELL');
  assert.ok(todo.get('DELL').modelos.has('LATITUDE 5420'));
});

test('filas sin marca se ignoran', () => {
  assert.equal(construirMaestro([{ marca: '', modelo: 'X1' }, { marca: null, modelo: 'Y' }]).size, 0);
});
