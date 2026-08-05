/**
 * Clasificacion de bandeja (triage) — logica pura, testeable sin red.
 * Toma correos ya normalizados y devuelve prioridad + si piden accion tuya.
 */

const URGENT_WORDS = [
  'urgente', 'urgent', 'asap', 'hoy mismo', 'lo antes posible', 'deadline',
  'vence', 'vencimiento', 'critico', 'critical', 'importante', 'prioridad',
]
const ACTION_WORDS = [
  'puedes', 'podrias', 'necesito', 'favor de', 'por favor', 'me confirmas',
  'confirmame', 'revisa', 'aprueba', 'aprobar', 'firma', 'firmar', 'responde',
  'agenda', 'agendar', 'envia', 'enviar', '?', 'pending', 'action required',
  'please', 'could you', 'can you', 'need',
]

function scoreOf(text, words) {
  const t = (text || '').toLowerCase()
  let n = 0
  for (const w of words) { if (t.includes(w)) { n++ } }
  return n
}

/**
 * Clasifica un correo. `me` = direccion del usuario (para saber si va dirigido a el).
 * Devuelve { priority: 'alta'|'media'|'baja', needsMyAction: bool, reasons: [] }.
 */
export function classifyEmail(email, me = '') {
  const subject = email.subject || ''
  const preview = email.bodyPreview || email.preview || ''
  const blob = `${subject}\n${preview}`
  const reasons = []

  let urgency = scoreOf(blob, URGENT_WORDS)
  if (email.importance === 'high') { urgency += 2; reasons.push('marcado importancia alta') }
  if (email.isRead === false) { reasons.push('no leido') }

  const actionScore = scoreOf(blob, ACTION_WORDS)
  const meLower = (me || '').toLowerCase()
  const toMeDirectly = (email.toRecipients || []).some(
    r => (r.address || r.emailAddress?.address || '').toLowerCase() === meLower,
  )
  const inCcOnly = !toMeDirectly && (email.ccRecipients || []).some(
    r => (r.address || r.emailAddress?.address || '').toLowerCase() === meLower,
  )

  const needsMyAction = actionScore > 0 && (toMeDirectly || (email.toRecipients || []).length <= 2)
  if (needsMyAction) { reasons.push('parece pedir accion tuya') }
  if (inCcOnly) { reasons.push('estas solo en copia (CC)') }

  let priority = 'baja'
  if (urgency >= 2 || (urgency >= 1 && needsMyAction)) { priority = 'alta' }
  else if (urgency >= 1 || needsMyAction) { priority = 'media' }

  // Estar solo en CC baja la prioridad un escalon.
  if (inCcOnly && priority === 'alta') { priority = 'media' }
  else if (inCcOnly && priority === 'media') { priority = 'baja' }

  return { priority, needsMyAction: needsMyAction && !inCcOnly, reasons }
}

/** Ordena correos clasificados: alta > media > baja, y dentro por fecha desc. */
export function sortByPriority(items) {
  const rank = { alta: 0, media: 1, baja: 2 }
  return [...items].sort((a, b) => {
    const r = (rank[a.priority] ?? 3) - (rank[b.priority] ?? 3)
    if (r !== 0) { return r }
    return String(b.receivedDateTime || '').localeCompare(String(a.receivedDateTime || ''))
  })
}
