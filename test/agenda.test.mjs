import { test } from 'node:test'
import assert from 'node:assert/strict'
import { groupByDay, findConflicts, freeSlotsByDay } from '../out/agendaCore.js'

const M = (subject, start, end) => ({ subject, start, end })

test('groupByDay agrupa y ordena', () => {
  const g = groupByDay([
    M('B', '2026-08-06 11:00', '2026-08-06 12:00'),
    M('A', '2026-08-06 09:00', '2026-08-06 10:00'),
    M('C', '2026-08-05 09:00', '2026-08-05 10:00'),
  ])
  assert.equal(g.length, 2)
  assert.equal(g[0].date, '2026-08-05')
  assert.deepEqual(g[1].items.map(x => x.subject), ['A', 'B'])
})

test('findConflicts detecta empalmes', () => {
  const c = findConflicts([
    M('A', '2026-08-06 09:00', '2026-08-06 10:00'),
    M('B', '2026-08-06 09:30', '2026-08-06 10:30'),
    M('C', '2026-08-06 11:00', '2026-08-06 12:00'),
  ])
  assert.equal(c.length, 1)
  assert.deepEqual(c[0].map(x => x.subject), ['A', 'B'])
})

test('freeSlotsByDay respeta horario laboral y ocupados', () => {
  const [day] = freeSlotsByDay([M('A', '2026-08-06 10:00', '2026-08-06 11:00')], 9, 18, 30)
  assert.equal(day.date, '2026-08-06')
  assert.deepEqual(day.slots, [{ start: '09:00', end: '10:00' }, { start: '11:00', end: '18:00' }])
})

test('freeSlotsByDay sin reuniones = dia completo', () => {
  const r = freeSlotsByDay([M('A', '2026-08-06 09:00', '2026-08-06 09:00')], 9, 18, 30)
  assert.equal(r[0].slots.length, 1)
  assert.deepEqual(r[0].slots[0], { start: '09:00', end: '18:00' })
})
