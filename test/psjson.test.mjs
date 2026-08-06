import { test } from 'node:test'
import assert from 'node:assert/strict'
import { cleanPsJson, parsePsArray } from '../out/psJson.js'

test('parsePsArray parsea JSON con controles crudos (C0, DEL, C1)', () => {
  const raw = '[{"subject":"hola mundo","body":"linea1\nlinea2"}]'
  const arr = parsePsArray(raw)
  assert.equal(arr.length, 1)
  assert.ok(arr[0].subject.startsWith('hola'))
})

test('cleanPsJson deja intactos los escapes validos', () => {
  const s = cleanPsJson('{"a":"b\\nc"}')
  assert.equal(s, '{"a":"b\\nc"}')
})

test('parsePsArray vacio => []', () => {
  assert.deepEqual(parsePsArray(''), [])
  assert.deepEqual(parsePsArray('[]'), [])
})

test('parsePsArray envuelve objeto suelto', () => {
  assert.equal(parsePsArray('{"x":1}').length, 1)
})
