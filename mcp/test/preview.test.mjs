import { test } from 'node:test'
import assert from 'node:assert/strict'
import { buildPreview, isConfirmed } from '../lib/preview.mjs'
import { capabilities } from '../config.mjs'

test('buildPreview incluye la accion y omite campos vacios', () => {
  const p = buildPreview('Enviar correo', { 'Para': 'a@b.com', 'CC': '', 'Asunto': 'Hola' })
  assert.ok(p.includes('Enviar correo'))
  assert.ok(p.includes('Para: a@b.com'))
  assert.ok(p.includes('Asunto: Hola'))
  assert.ok(!p.includes('CC:'))
  assert.ok(p.includes('confirm: true'))
})

test('isConfirmed solo con true explicito', () => {
  assert.equal(isConfirmed({ confirm: true }), true)
  assert.equal(isConfirmed({ confirm: false }), false)
  assert.equal(isConfirmed({}), false)
  assert.equal(isConfirmed({ confirm: 'true' }), false)
})

test('capabilities deriva permisos correctamente', () => {
  const full = capabilities(['Mail.ReadWrite', 'Mail.Send', 'Calendars.ReadWrite', 'Chat.ReadWrite'])
  assert.deepEqual(full, {
    readMail: true, draftMail: true, sendMail: true,
    readCalendar: true, writeCalendar: true, readChat: true, sendChat: true,
  })
})

test('capabilities solo lectura no permite escritura', () => {
  const ro = capabilities(['Mail.Read', 'Calendars.Read'])
  assert.equal(ro.readMail, true)
  assert.equal(ro.draftMail, false)
  assert.equal(ro.sendMail, false)
  assert.equal(ro.writeCalendar, false)
  assert.equal(ro.readChat, false)
})

test('capabilities vacio = todo en false', () => {
  const none = capabilities([])
  assert.ok(Object.values(none).every(v => v === false))
})
