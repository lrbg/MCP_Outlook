import * as vscode from 'vscode'
import { DayEntry, computeKeySenders, upsertEntry, trimEntries } from './bitacoraCore'
import { getUnread } from './outlookRead'
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
  const max = vscode.workspace.getConfiguration('m365').get<number>('dailyReview.maxEmails', 30)
  const emails = getUnread(max)

  let notes = ''
  try {
    notes = await vscode.window.withProgress(
      { location: vscode.ProgressLocation.Notification, title: 'Revision de bandeja: Copilot esta resumiendo…' },
      (_p, token) => summarizeInbox(emails, token),
    )
  } catch (e: any) {
    notes = `_No se pudieron generar notas con Copilot: ${e?.message || e}_`
  }

  const entry: DayEntry = {
    date: today(),
    ranAt: nowStamp(),
    unreadCount: emails.length,
    keySenders: computeKeySenders(emails),
    notesMarkdown: notes,
  }

  const entries = trimEntries(upsertEntry(await loadEntries(context), entry))
  await saveEntries(context, entries)
  await context.globalState.update('m365.lastReviewDate', entry.date)
  return entry
}

/** ¿Ya se corrio la revision hoy? */
export function reviewDoneToday(context: vscode.ExtensionContext): boolean {
  return context.globalState.get<string>('m365.lastReviewDate') === today()
}

/**
 * Programador: si esta habilitado, no se ha corrido hoy y ya paso la hora
 * configurada, corre la revision. Se llama al activar y cada cierto rato.
 */
export async function maybeRunScheduled(context: vscode.ExtensionContext, onDone: () => void): Promise<void> {
  const c = vscode.workspace.getConfiguration('m365')
  if (!c.get<boolean>('dailyReview.enabled', true)) { return }
  if (process.platform !== 'win32') { return }
  if (reviewDoneToday(context)) { return }
  const hour = c.get<number>('dailyReview.hour', 8)
  if (new Date().getHours() < hour) { return }
  try { await runDailyReview(context); onDone() } catch { /* se reintenta al proximo tick */ }
}
