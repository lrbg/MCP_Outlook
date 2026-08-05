import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mergeBusy, freeSlots, findConflicts, eventToBusy } from '../lib/freebusy.mjs'

const H = (h) => Date.UTC(2026, 7, 5, h, 0, 0) // 2026-08-05 hh:00 UTC

test('mergeBusy une intervalos solapados', () => {
  const merged = mergeBusy([
    { start: H(9), end: H(10) },
    { start: H(9.5 | 0), end: H(11) },
    { start: H(13), end: H(14) },
  ])
  assert.equal(merged.length, 2)
  assert.equal(merged[0].start, H(9))
  assert.equal(merged[0].end, H(11))
})

test('freeSlots devuelve huecos entre ocupados', () => {
  const busy = [
    { start: H(9), end: H(10) },
    { start: H(12), end: H(13) },
  ]
  const slots = freeSlots(H(8), H(18), busy, 30)
  // 8-9, 10-12, 13-18
  assert.equal(slots.length, 3)
  assert.equal(slots[0].start, H(8))
  assert.equal(slots[0].end, H(9))
  assert.equal(slots[1].start, H(10))
  assert.equal(slots[1].end, H(12))
})

test('freeSlots respeta la duracion minima', () => {
  const busy = [{ start: H(9), end: H(9) + 50 * 60000 }] // ocupa 9:00-9:50
  const slots = freeSlots(H(9), H(10), busy, 30)
  // solo quedan 10 min (9:50-10:00) < 30 -> sin huecos
  assert.equal(slots.length, 0)
})

test('freeSlots sin ocupados = toda la ventana', () => {
  const slots = freeSlots(H(9), H(11), [], 30)
  assert.equal(slots.length, 1)
  assert.equal(slots[0].start, H(9))
  assert.equal(slots[0].end, H(11))
})

test('findConflicts detecta solapes', () => {
  const events = [
    { start: { dateTime: new Date(H(9)).toISOString() }, end: { dateTime: new Date(H(10)).toISOString() }, subject: 'A' },
    { start: { dateTime: new Date(H(9.5 | 0)).toISOString() }, end: { dateTime: new Date(H(11)).toISOString() }, subject: 'B' },
    { start: { dateTime: new Date(H(12)).toISOString() }, end: { dateTime: new Date(H(13)).toISOString() }, subject: 'C' },
  ]
  const conflicts = findConflicts(events)
  assert.equal(conflicts.length, 1)
  assert.deepEqual(conflicts[0].map(e => e.subject), ['A', 'B'])
})

test('eventToBusy acepta objeto Graph y string', () => {
  const b1 = eventToBusy({ start: { dateTime: '2026-08-05T09:00:00Z' }, end: { dateTime: '2026-08-05T10:00:00Z' } })
  assert.ok(Number.isFinite(b1.start) && b1.end > b1.start)
})
