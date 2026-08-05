import { test } from 'node:test'
import assert from 'node:assert/strict'
import { classifyEmail, sortByPriority } from '../lib/triage.mjs'

const ME = 'rogelio@corp.com'

test('correo urgente dirigido a mi = prioridad alta y accion', () => {
  const r = classifyEmail({
    subject: 'URGENTE: necesito tu aprobacion hoy',
    bodyPreview: 'Puedes confirmarme?',
    isRead: false,
    toRecipients: [{ address: ME }],
  }, ME)
  assert.equal(r.priority, 'alta')
  assert.equal(r.needsMyAction, true)
})

test('estar solo en CC baja la prioridad y no marca accion', () => {
  const r = classifyEmail({
    subject: 'FYI presupuesto',
    bodyPreview: 'Los dejo en copia, sin accion.',
    isRead: false,
    toRecipients: [{ address: 'otro@corp.com' }],
    ccRecipients: [{ address: ME }],
  }, ME)
  assert.equal(r.needsMyAction, false)
  assert.notEqual(r.priority, 'alta')
})

test('correo informativo sin urgencia = baja', () => {
  const r = classifyEmail({
    subject: 'Boletin semanal',
    bodyPreview: 'Novedades de la empresa.',
    isRead: true,
    toRecipients: [{ address: 'lista@corp.com' }],
  }, ME)
  assert.equal(r.priority, 'baja')
})

test('importancia alta de Graph sube la urgencia', () => {
  const r = classifyEmail({
    subject: 'Revision',
    bodyPreview: 'Cuando puedas.',
    importance: 'high',
    isRead: false,
    toRecipients: [{ address: ME }],
  }, ME)
  assert.ok(['alta', 'media'].includes(r.priority))
})

test('sortByPriority ordena alta > media > baja', () => {
  const items = [
    { priority: 'baja', receivedDateTime: '2026-08-05T10:00:00Z' },
    { priority: 'alta', receivedDateTime: '2026-08-05T09:00:00Z' },
    { priority: 'media', receivedDateTime: '2026-08-05T11:00:00Z' },
  ]
  const sorted = sortByPriority(items)
  assert.deepEqual(sorted.map(x => x.priority), ['alta', 'media', 'baja'])
})

test('acepta el formato emailAddress.address de Graph', () => {
  const r = classifyEmail({
    subject: 'Puedes revisar?',
    bodyPreview: 'Gracias',
    toRecipients: [{ emailAddress: { address: ME } }],
  }, ME)
  assert.equal(r.needsMyAction, true)
})
