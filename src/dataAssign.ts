import * as vscode from 'vscode'

/**
 * Asignaciones de solicitudes de datos sinteticos: por id de correo, a quien se
 * asigno, cuando y su estatus. Persistido en globalStorage.
 */
export type ReqStatus = 'seguimiento' | 'finalizada'
export interface Assignment { member: string; assignedAt: number; status: ReqStatus }
export type AssignMap = Record<string, Assignment>

function uri(context: vscode.ExtensionContext, file: string): vscode.Uri {
  return vscode.Uri.joinPath(context.globalStorageUri, file)
}

export async function loadAssignments(context: vscode.ExtensionContext, file: string): Promise<AssignMap> {
  try {
    return JSON.parse(Buffer.from(await vscode.workspace.fs.readFile(uri(context, file))).toString('utf8'))
  } catch { return {} }
}

async function save(context: vscode.ExtensionContext, file: string, map: AssignMap): Promise<void> {
  await vscode.workspace.fs.createDirectory(context.globalStorageUri)
  await vscode.workspace.fs.writeFile(uri(context, file), Buffer.from(JSON.stringify(map, null, 2), 'utf8'))
}

export async function assign(context: vscode.ExtensionContext, file: string, id: string, member: string, now: number): Promise<void> {
  const map = await loadAssignments(context, file)
  const prev = map[id]
  map[id] = { member, assignedAt: prev && prev.member === member ? prev.assignedAt : now, status: prev?.status || 'seguimiento' }
  await save(context, file, map)
}

export async function setStatus(context: vscode.ExtensionContext, file: string, id: string, status: ReqStatus): Promise<void> {
  const map = await loadAssignments(context, file)
  if (map[id]) { map[id].status = status; await save(context, file, map) }
}

export async function unassign(context: vscode.ExtensionContext, file: string, id: string): Promise<void> {
  const map = await loadAssignments(context, file)
  delete map[id]
  await save(context, file, map)
}
