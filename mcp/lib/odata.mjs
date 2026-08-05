/**
 * Helpers puros para armar consultas OData de Microsoft Graph.
 * Sin dependencias: testeables con `node --test`.
 */

/** Escapa comillas simples para literales OData ('' es la comilla escapada). */
export function odataEscape(value) {
  return String(value ?? '').replace(/'/g, "''")
}

/** Convierte un objeto de params en un query string OData ($filter, $top, etc.). */
export function buildQuery(params = {}) {
  const parts = []
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null || v === '') { continue }
    parts.push(`${k}=${encodeURIComponent(v)}`)
  }
  return parts.length ? '?' + parts.join('&') : ''
}

/**
 * Construye el $filter para listar correos. Combina condiciones con 'and'.
 * - unreadOnly: isRead eq false
 * - from: contiene remitente (por direccion)
 * - subject: contiene asunto
 * - since / until: recibidos en el rango (ISO 8601)
 */
export function buildMailFilter({ unreadOnly, from, subject, since, until } = {}) {
  const conds = []
  if (unreadOnly) { conds.push('isRead eq false') }
  if (from) { conds.push(`contains(from/emailAddress/address,'${odataEscape(from)}')`) }
  if (subject) { conds.push(`contains(subject,'${odataEscape(subject)}')`) }
  if (since) { conds.push(`receivedDateTime ge ${since}`) }
  if (until) { conds.push(`receivedDateTime le ${until}`) }
  return conds.join(' and ')
}

/** Rango [inicio, fin) de un dia (ISO) en un offset de dias desde hoy. */
export function dayRange(baseIsoDate, offsetDays = 0) {
  const d = new Date(baseIsoDate)
  d.setUTCDate(d.getUTCDate() + offsetDays)
  const start = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), 0, 0, 0))
  const end = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() + 1, 0, 0, 0))
  return { start: start.toISOString(), end: end.toISOString() }
}
