/**
 * Logica pura de agenda (sin vscode): agrupar reuniones por dia, detectar
 * conflictos (empalmes) y calcular huecos libres. Testeable con node --test.
 */

export interface Meeting {
  subject?: string
  start?: string
  end?: string
  location?: string
  organizer?: string
  attendees?: number
}

export interface DayGroup { date: string; items: Meeting[] }

/** Agrupa por fecha (yyyy-MM-dd del inicio) y ordena dias y reuniones ascendente. */
export function groupByDay(meetings: Meeting[]): DayGroup[] {
  const map = new Map<string, Meeting[]>()
  for (const m of meetings) {
    const date = (m.start || '').slice(0, 10)
    if (!date) { continue }
    if (!map.has(date)) { map.set(date, []) }
    map.get(date)!.push(m)
  }
  return [...map.entries()]
    .map(([date, items]) => ({ date, items: items.sort((a, b) => String(a.start).localeCompare(String(b.start))) }))
    .sort((a, b) => a.date.localeCompare(b.date))
}

/** Pares de reuniones que se empalman en el tiempo. */
export function findConflicts(meetings: Meeting[]): [Meeting, Meeting][] {
  const withTs = meetings
    .map(m => ({ m, s: Date.parse(m.start || ''), e: Date.parse(m.end || '') }))
    .filter(x => Number.isFinite(x.s) && Number.isFinite(x.e) && x.e > x.s)
    .sort((a, b) => a.s - b.s)
  const out: [Meeting, Meeting][] = []
  for (let i = 0; i < withTs.length; i++) {
    for (let j = i + 1; j < withTs.length; j++) {
      if (withTs[j].s >= withTs[i].e) { break }
      out.push([withTs[i].m, withTs[j].m])
    }
  }
  return out
}

/**
 * Huecos libres por dia en horario laboral [startHour, endHour), de al menos
 * `minMinutes`. Devuelve { date, slots: [{start,end}] } con horas ISO locales.
 */
export function freeSlotsByDay(meetings: Meeting[], startHour = 9, endHour = 18, minMinutes = 30): { date: string; slots: { start: string; end: string }[] }[] {
  const minMs = minMinutes * 60 * 1000
  return groupByDay(meetings).map(({ date, items }) => {
    const [y, mo, d] = date.split('-').map(Number)
    const dayStart = new Date(y, mo - 1, d, startHour, 0, 0).getTime()
    const dayEnd = new Date(y, mo - 1, d, endHour, 0, 0).getTime()
    const busy = items
      .map(m => ({ s: Date.parse(m.start || ''), e: Date.parse(m.end || '') }))
      .filter(b => Number.isFinite(b.s) && Number.isFinite(b.e))
      .sort((a, b) => a.s - b.s)
    const slots: { start: string; end: string }[] = []
    let cursor = dayStart
    for (const b of busy) {
      if (b.e <= dayStart || b.s >= dayEnd) { continue }
      const gapEnd = Math.min(b.s, dayEnd)
      if (gapEnd - cursor >= minMs) { slots.push({ start: hhmm(cursor), end: hhmm(gapEnd) }) }
      cursor = Math.max(cursor, Math.min(b.e, dayEnd))
    }
    if (dayEnd - cursor >= minMs) { slots.push({ start: hhmm(cursor), end: hhmm(dayEnd) }) }
    return { date, slots }
  })
}

function hhmm(ts: number): string {
  const d = new Date(ts)
  const p = (n: number) => String(n).padStart(2, '0')
  return `${p(d.getHours())}:${p(d.getMinutes())}`
}
