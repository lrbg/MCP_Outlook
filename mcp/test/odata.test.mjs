import { test } from 'node:test'
import assert from 'node:assert/strict'
import { odataEscape, buildMailFilter, buildQuery, dayRange } from '../lib/odata.mjs'

test('odataEscape duplica comillas simples', () => {
  assert.equal(odataEscape("O'Brien"), "O''Brien")
  assert.equal(odataEscape(''), '')
  assert.equal(odataEscape(null), '')
})

test('buildMailFilter combina condiciones con and', () => {
  const f = buildMailFilter({ unreadOnly: true, from: 'jefe@corp.com', subject: 'presupuesto' })
  assert.ok(f.includes('isRead eq false'))
  assert.ok(f.includes("contains(from/emailAddress/address,'jefe@corp.com')"))
  assert.ok(f.includes("contains(subject,'presupuesto')"))
  assert.equal(f.split(' and ').length, 3)
})

test('buildMailFilter vacio si no hay filtros', () => {
  assert.equal(buildMailFilter({}), '')
})

test('buildMailFilter escapa comillas en el asunto', () => {
  const f = buildMailFilter({ subject: "O'Brien" })
  assert.ok(f.includes("O''Brien"))
})

test('buildQuery arma query string y omite vacios', () => {
  assert.equal(buildQuery({ $top: 5, $skip: '' }), '?$top=5')
  assert.equal(buildQuery({}), '')
})

test('dayRange devuelve un dia completo UTC', () => {
  const { start, end } = dayRange('2026-08-05T15:00:00Z', 0)
  assert.equal(start, '2026-08-05T00:00:00.000Z')
  assert.equal(end, '2026-08-06T00:00:00.000Z')
})

test('dayRange con offset', () => {
  const { start } = dayRange('2026-08-05T00:00:00Z', 1)
  assert.equal(start, '2026-08-06T00:00:00.000Z')
})
