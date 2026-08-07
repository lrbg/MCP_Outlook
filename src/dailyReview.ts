import * as vscode from 'vscode'
import { DayEntry, computeKeySenders, upsertEntry, trimEntries } from './bitacoraCore'
import { getRecent } from './outlookRead'
import { summarizeInbox } from './copilot'

const FILE = 'bitacora.json'

function fileUri(context: vscode.ExtensionContext): vscode.Uri {
  return vscode.Uri.joinPath(context.globalStorageUri, FILE)
}

export async function loadEntries(context: vscode.ExtensionContext): Promise<DayEntry[]> {
  try {
    const buf = await vscode.workspace.fs.readFile(fileUri(context))
    const j = JSON.parse(Buffer.from(buf).toString('utf8'))
    return Array.isArray(j) ? j : []
  } catch { return [] }
}

async function saveEntries(context: vscode.ExtensionContext, entries: DayEntry[]): Promise<void> {
  await vscode.workspace.fs.createDirectory(context.globalStorageUri)
  await vscode.workspace.fs.writeFile(fileUri(context), Buffer.from(JSON.stringify(entries, null, 2), 'utf8'))
}

function today(): string {
  const d = new Date()
  const p = (n: number) => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`
}
function nowStamp(): string {
  const d = new Date()
  const p = (n: number) => String(n).padStart(2, '0')
  return `${today()} ${p(d.getHours())}:${p(d.getMinutes())}`
}

/**
 * Corre la revision del dia: lee no-leidos (COM), pide notas a Copilot y guarda
 * la entrada en la bitacora. Devuelve la entrada creada.
 */
export async function runDailyReview(context: vscode.ExtensionContext): Promise<DayEntry> {
  const max = vscode.workspace.getConfiguration('m365').get<number>('dailyReview.maxEmails', 40)
  const inbox = getRecent(6, max, 'ReceivedTime')
  const sent = getRecent(5, max, 'SentOn')

  let notes = ''
  try {
    notes = await vscode.window.withProgress(
      { location: vscode.ProgressLocation.Notification, title: 'Revision de correo: Copilot esta resumiendo…' },
      (_p, token) => summarizeInbox(inbox, sent, token),
    )
  } catch (e: any) {
    notes = `_No se pudieron generar notas con Copilot: ${e?.message || e}_`
  }

  const entry: DayEntry = {
    date: today(),
    ranAt: nowStamp(),
    unreadCount: inbox.filter(e => e.unread).length,
    keySenders: computeKeySenders(inbox),
    notesMarkdown: notes,
  }

  const entries = trimEntries(upsertEntry(await loadEntries(context), entry))
  await saveEntries(context, entries)
  await context.globalState.update('m365.lastReviewDate', entry.date)
  await context.globalState.update('m365.lastReviewAt', Date.now())
  return entry
}

/** ¿Ya se corrio la revision hoy? */
export function reviewDoneToday(context: vscode.ExtensionContext): boolean {
  return context.globalState.get<string>('m365.lastReviewDate') === today()
}

/**
 * Programador de SONDEO: corre la revision cada `pollMinutes` minutos, solo en
 * los dias y la ventana horaria configurados. Se llama en un tick frecuente
 * (cada minuto); aqui decide si toca correr. Asi el agente sondea correos nuevos
 * y los que ya respondiste con la frecuencia elegida.
 */
export async function maybeRunScheduled(context: vscode.ExtensionContext, onDone: () => void): Promise<void> {
  const c = vscode.workspace.getConfiguration('m365')
  if (!c.get<boolean>('dailyReview.enabled', true)) { return }
  if (process.platform !== 'win32') { return }

  const now = new Date()
  const days = c.get<number[]>('dailyReview.days', [1, 2, 3, 4, 5])
  if (!days.includes(now.getDay())) { return } // getDay: 0=Dom .. 6=Sab

  const start = c.get<number>('dailyReview.startHour', 7)
  const end = c.get<number>('dailyReview.endHour', 3)
  const h = now.getHours()
  const inWindow = start <= end ? (h >= start && h < end) : (h >= start || h < end)
  if (!inWindow) { return }

  const poll = c.get<number>('dailyReview.pollMinutes', 30)
  const lastAt = context.globalState.get<number>('m365.lastReviewAt', 0)
  if (Date.now() - lastAt < poll * 60 * 1000) { return }

  try { await runDailyReview(context); onDone() } catch { /* se reintenta al proximo tick */ }
}
