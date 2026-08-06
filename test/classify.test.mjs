import { test } from 'node:test'
import assert from 'node:assert/strict'
import { classifyPriority, labelText } from '../out/priorityClassify.js'

const me = { name: 'Batres, Rogelio', email: 'rogelio@empresa.com', tokens: ['@Batres'] }

test('directed cuando vas en Para por email', () => {
  const c = classifyPriority({ to: 'Rogelio <rogelio@empresa.com>', cc: '' }, me)
  assert.equal(c.label, 'directed')
  assert.equal(c.needsAction, true)
})

test('directed por nombre en Para', () => {
  const c = classifyPriority({ to: 'Batres, Rogelio; Otro' }, me)
  assert.equal(c.label, 'directed')
})

test('mentioned cuando te taggean en el cuerpo', () => {
  const c = classifyPriority({ to: 'equipo@empresa.com', bodySnippet: 'Hola @Batres, puedes revisar?' }, me)
  assert.equal(c.label, 'mentioned')
  assert.equal(c.needsAction, true)
})

test('informative cuando solo vas en copia', () => {
  const c = classifyPriority({ to: 'otro@empresa.com', cc: 'rogelio@empresa.com' }, me)
  assert.equal(c.label, 'informative')
  assert.equal(c.needsAction, false)
  assert.equal(c.onlyCc, true)
})

test('labelText traduce', () => {
  assert.equal(labelText('directed'), 'Para ti')
  assert.equal(labelText('informative'), 'Informativo')
})
