import { test } from 'node:test'
import assert from 'node:assert/strict'
import { computeKeySenders, upsertEntry, trimEntries, toMarkdown } from '../out/bitacoraCore.js'

test('computeKeySenders cuenta y ordena por frecuencia', () => {
  const ks = computeKeySenders([
    { sender: 'Ana' }, { sender: 'Beto' }, { sender: 'Ana' }, { sender: 'Ana' }, { sender: 'Beto' },
  ])
  assert.equal(ks[0].name, 'Ana')
  assert.equal(ks[0].count, 3)
  assert.equal(ks[1].name, 'Beto')
  assert.equal(ks[1].count, 2)
})

test('computeKeySenders usa senderEmail si no hay nombre y respeta topN', () => {
  const ks = computeKeySenders([{ senderEmail: 'x@y.com' }, {}], 1)
  assert.equal(ks.length, 1)
})

test('upsertEntry reemplaza la entrada del mismo dia y ordena desc', () => {
  const a = { date: '2026-08-05', ranAt: '', unreadCount: 1, keySenders: [], notesMarkdown: '' }
  const b = { date: '2026-08-06', ranAt: '', unreadCount: 2, keySenders: [], notesMarkdown: '' }
  const bDup = { date: '2026-08-06', ranAt: '', unreadCount: 9, keySenders: [], notesMarkdown: '' }
  let list = upsertEntry([a], b)
  list = upsertEntry(list, bDup)
  assert.equal(list.length, 2)
  assert.equal(list[0].date, '2026-08-06')
  assert.equal(list[0].unreadCount, 9)
})

test('trimEntries deja los mas recientes', () => {
  const entries = ['2026-08-01', '2026-08-02', '2026-08-03'].map(d => ({ date: d, ranAt: '', unreadCount: 0, keySenders: [], notesMarkdown: '' }))
  const t = trimEntries(entries, 2)
  assert.equal(t.length, 2)
  assert.equal(t[0].date, '2026-08-03')
})

test('toMarkdown arma tabla + notas por dia', () => {
  const md = toMarkdown([{ date: '2026-08-06', ranAt: '2026-08-06 08:00', unreadCount: 4, keySenders: [{ name: 'Ana', count: 2 }], notesMarkdown: '- Revisar KPI' }])
  assert.ok(md.includes('| Fecha | No leidos | Remitentes clave |'))
  assert.ok(md.includes('| 2026-08-06 | 4 | Ana (2) |'))
  assert.ok(md.includes('Revisar KPI'))
})
