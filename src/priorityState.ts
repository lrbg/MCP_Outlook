import * as vscode from 'vscode'

/**
 * Estado por correo prioritario que fija el usuario: 'handled' (atendido) o
 * 'dismissed' (no requiere respuesta / informativo). Persistido por EntryID en
 * globalStorage. Se limpian los ids con mas de 45 dias para no crecer sin fin.
 */
export type MailStatus = 'handled' | 'dismissed'

interface StoreShape { [id: string]: { status: MailStatus; at: number } }

const FILE = 'priorityState.json'

function uri(context: vscode.ExtensionContext): vscode.Uri {
  return vscode.Uri.joinPath(context.globalStorageUri, FILE)
}

export async function loadStatus(context: vscode.ExtensionContext): Promise<Record<string, MailStatus>> {
  try {
    const raw = JSON.parse(Buffer.from(await vscode.workspace.fs.readFile(uri(context))).toString('utf8')) as StoreShape
    const out: Record<string, MailStatus> = {}
    for (const [id, v] of Object.entries(raw)) { out[id] = v.status }
    return out
  } catch { return {} }
}

async function readRaw(context: vscode.ExtensionContext): Promise<StoreShape> {
  try { return JSON.parse(Buffer.from(await vscode.workspace.fs.readFile(uri(context))).toString('utf8')) }
  catch { return {} }
}

async function writeRaw(context: vscode.ExtensionContext, store: StoreShape): Promise<void> {
  await vscode.workspace.fs.createDirectory(context.globalStorageUri)
  await vscode.workspace.fs.writeFile(uri(context), Buffer.from(JSON.stringify(store, null, 2), 'utf8'))
}

export async function setStatus(context: vscode.ExtensionContext, id: string, status: MailStatus, now: number): Promise<void> {
  const store = await readRaw(context)
  store[id] = { status, at: now }
  prune(store, now)
  await writeRaw(context, store)
}

export async function clearStatus(context: vscode.ExtensionContext, id: string): Promise<void> {
  const store = await readRaw(context)
  delete store[id]
  await writeRaw(context, store)
}

function prune(store: StoreShape, now: number): void {
  const cutoff = now - 45 * 24 * 60 * 60 * 1000
  for (const [id, v] of Object.entries(store)) { if (v.at < cutoff) { delete store[id] } }
}
