/**
 * Logica pura de la bitacora de revisiones de bandeja (sin dependencias de
 * vscode), para poder probarla con `node --test`.
 */

export interface KeySender { name: string; count: number }

export interface DayEntry {
  /** Fecha del dia, yyyy-MM-dd. */
  date: string
  /** Cuando se corrio la revision, yyyy-MM-dd HH:mm. */
  ranAt: string
  unreadCount: number
  keySenders: KeySender[]
  /** Notas generadas por Copilot (markdown). */
  notesMarkdown: string
}

export interface RawEmail {
  subject?: string
  sender?: string
  senderEmail?: string
  received?: string
  unread?: boolean
  preview?: string
  to?: string
}

/** Remitentes mas frecuentes entre los correos dados. */
export function computeKeySenders(emails: RawEmail[], topN = 5): KeySender[] {
  const counts = new Map<string, number>()
  for (const e of emails) {
    const name = (e.sender || e.senderEmail || '(desconocido)').trim()
    counts.set(name, (counts.get(name) || 0) + 1)
  }
  return [...counts.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, topN)
}

/**
 * Inserta o reemplaza la entrada de un dia (una por fecha) y devuelve la lista
 * ordenada de la mas reciente a la mas antigua.
 */
export function upsertEntry(entries: DayEntry[], entry: DayEntry): DayEntry[] {
  const rest = entries.filter(e => e.date !== entry.date)
  return [entry, ...rest].sort((a, b) => b.date.localeCompare(a.date))
}

/** Recorta la bitacora a los ultimos N dias (para no crecer sin fin). */
export function trimEntries(entries: DayEntry[], keep = 60): DayEntry[] {
  return [...entries].sort((a, b) => b.date.localeCompare(a.date)).slice(0, keep)
}

function sendersText(ks: KeySender[]): string {
  return ks.map(s => `${s.name} (${s.count})`).join(', ') || '-'
}

/** Exporta la bitacora completa como Markdown: tabla resumen + notas por dia. */
export function toMarkdown(entries: DayEntry[]): string {
  const sorted = [...entries].sort((a, b) => b.date.localeCompare(a.date))
  const lines: string[] = []
  lines.push('# Bitacora de revisiones de bandeja', '')
  lines.push('| Fecha | No leidos | Remitentes clave |')
  lines.push('|---|---|---|')
  for (const e of sorted) {
    lines.push(`| ${e.date} | ${e.unreadCount} | ${sendersText(e.keySenders)} |`)
  }
  lines.push('', '## Notas por dia', '')
  for (const e of sorted) {
    lines.push(`### ${e.date}  (${e.unreadCount} no leidos, corrida ${e.ranAt})`, '')
    lines.push(e.notesMarkdown?.trim() || '_Sin notas._', '')
  }
  return lines.join('\n')
}
