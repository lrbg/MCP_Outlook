/**
 * Calculo de huecos libres a partir de eventos ocupados — logica pura.
 * Todas las horas se manejan como epoch ms para no arrastrar zonas horarias.
 */

/** Normaliza un evento de Graph a { start, end } en epoch ms. */
export function eventToBusy(ev) {
  const start = Date.parse(ev.start?.dateTime || ev.start)
  const end = Date.parse(ev.end?.dateTime || ev.end)
  return { start, end }
}

/** Une intervalos ocupados solapados. Entrada/salida: [{start,end}] en ms. */
export function mergeBusy(busy) {
  const valid = busy.filter(b => Number.isFinite(b.start) && Number.isFinite(b.end) && b.end > b.start)
  const sorted = [...valid].sort((a, b) => a.start - b.start)
  const out = []
  for (const b of sorted) {
    const last = out[out.length - 1]
    if (last && b.start <= last.end) { last.end = Math.max(last.end, b.end) }
    else { out.push({ ...b }) }
  }
  return out
}

/**
 * Devuelve los huecos libres dentro de [windowStart, windowEnd] que no chocan
 * con los intervalos ocupados y duran al menos `minMinutes`.
 * Todos los parametros de tiempo en epoch ms; salida igual.
 */
export function freeSlots(windowStart, windowEnd, busy, minMinutes = 30) {
  const minMs = minMinutes * 60 * 1000
  const merged = mergeBusy(busy)
  const slots = []
  let cursor = windowStart
  for (const b of merged) {
    if (b.end <= windowStart || b.start >= windowEnd) { continue }
    const gapEnd = Math.min(b.start, windowEnd)
    if (gapEnd - cursor >= minMs) { slots.push({ start: cursor, end: gapEnd }) }
    cursor = Math.max(cursor, Math.min(b.end, windowEnd))
  }
  if (windowEnd - cursor >= minMs) { slots.push({ start: cursor, end: windowEnd }) }
  return slots
}

/** Detecta eventos que se solapan entre si (conflictos de agenda). */
export function findConflicts(events) {
  const busy = events.map((e, i) => ({ ...eventToBusy(e), i }))
    .filter(b => Number.isFinite(b.start) && Number.isFinite(b.end))
    .sort((a, b) => a.start - b.start)
  const conflicts = []
  for (let i = 0; i < busy.length; i++) {
    for (let j = i + 1; j < busy.length; j++) {
      if (busy[j].start >= busy[i].end) { break }
      conflicts.push([events[busy[i].i], events[busy[j].i]])
    }
  }
  return conflicts
}
