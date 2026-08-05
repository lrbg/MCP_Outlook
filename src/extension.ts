import * as vscode from 'vscode'
import {
  McpServerEntry, McpFileCorruptError,
  userMcpPathFromGlobalStorage, mergeMcpServers, hasMcpServer,
} from './mcpConfig'
import { buildM365Runtime, registerM365McpProvider, M365McpProvider } from './mcpProvider'
import { signIn, refreshSilent } from './auth'

let mcpProvider: M365McpProvider | undefined
let refreshTimer: ReturnType<typeof setInterval> | undefined

export function activate(context: vscode.ExtensionContext) {
  mcpProvider = registerM365McpProvider(context)

  const reg = (id: string, fn: (...a: any[]) => any) =>
    context.subscriptions.push(vscode.commands.registerCommand(id, fn))

  reg('m365.signIn', async () => {
    try {
      const s = await signIn(context)
      await buildM365Runtime(context)
      mcpProvider?.refresh()
      await vscode.commands.executeCommand('m365.registerMcpServer', { silent: true })
      vscode.window.showInformationMessage(
        `Sesion iniciada como ${s.account}. Permisos: ${s.scopes.join(', ') || '(ninguno)'}. ` +
        'Registra el MCP con "M365: Registrar servidor MCP" si es la primera vez.',
      )
    } catch (e: any) {
      vscode.window.showErrorMessage(`No se pudo iniciar sesion: ${e?.message || e}`)
    }
  })

  reg('m365.refresh', async () => {
    const s = await refreshSilent(context)
    await buildM365Runtime(context)
    mcpProvider?.refresh()
    await vscode.commands.executeCommand('m365.registerMcpServer', { silent: true })
    if (s) {
      vscode.window.showInformationMessage(`Sesion refrescada (${s.account}).`)
    } else {
      vscode.window.showWarningMessage('No hay sesion activa. Ejecuta "M365: Iniciar sesion".')
    }
  })

  reg('m365.status', async () => {
    const rt = await buildM365Runtime(context)
    if (!rt.hasSession) {
      vscode.window.showWarningMessage('Sin sesion de Microsoft. Ejecuta "M365: Iniciar sesion".')
      return
    }
    vscode.window.showInformationMessage(
      `Microsoft 365 conectado. Permisos concedidos: ${rt.scopes.join(', ') || '(ninguno)'}.`,
    )
  })

  reg('m365.registerMcpServer', async (opts?: { silent?: boolean }) => {
    const silent = !!opts?.silent
    const c = vscode.workspace.getConfiguration('m365')

    const rt = await buildM365Runtime(context)
    const entry: McpServerEntry = { command: 'node', args: [rt.serverPath], env: rt.env }

    const ws = vscode.workspace.workspaceFolders?.[0]
    const wsFile = ws ? vscode.Uri.joinPath(ws.uri, '.vscode', 'mcp.json') : undefined
    const userFile = vscode.Uri.file(userMcpPathFromGlobalStorage(context.globalStorageUri.fsPath))

    if (silent) {
      // Refresca donde ya este registrado, sin preguntar ni crear archivos nuevos.
      for (const f of [userFile, wsFile]) {
        if (f && await mcpFileHasM365(f)) {
          try { await writeMcpEntry(f, entry) } catch { /* se avisara al registrar a mano */ }
        }
      }
      return
    }

    let scope = c.get<string>('mcpScope', 'ask')
    if (scope !== 'user' && scope !== 'workspace') {
      if (!wsFile) {
        scope = 'user'
      } else {
        const pick = await vscode.window.showQuickPick(
          [
            { label: 'Todo VS Code', description: 'Disponible en cualquier ventana', scope: 'user' },
            { label: 'Solo este proyecto', description: 'Escribe .vscode/mcp.json en la carpeta abierta', scope: 'workspace' },
          ],
          { title: 'M365: donde registro el servidor MCP?', ignoreFocusOut: true },
        )
        if (!pick) { return }
        scope = pick.scope
        await c.update('mcpScope', scope, vscode.ConfigurationTarget.Global)
      }
    }

    const target = (scope === 'workspace' && wsFile) ? wsFile : userFile
    try {
      await writeMcpEntry(target, entry)
      const donde = target === userFile
        ? 'MCP registrado para todo VS Code.'
        : 'MCP registrado en .vscode/mcp.json de este proyecto.'
      vscode.window.showInformationMessage(
        `${donde} Presiona "Start" (o "Restart") para usarlo en Copilot Chat.`,
      )
      vscode.window.showTextDocument(target)
    } catch (e: any) {
      if (e instanceof McpFileCorruptError) {
        const abrir = 'Abrir archivo'
        const r = await vscode.window.showErrorMessage(
          `No toque ${target.fsPath} porque no se pudo leer su contenido (${e.message}). Corrigelo y vuelve a registrar.`,
          abrir,
        )
        if (r === abrir) { vscode.window.showTextDocument(target) }
      } else {
        vscode.window.showErrorMessage(`No se pudo registrar el MCP: ${e?.message || e}`)
      }
    }
  })

  // Escribe la config al arrancar (reusa la sesion si ya existe) y refresca el
  // token cada 40 min mientras VS Code este abierto, para minimizar 401 por token
  // expirado (opcion B: token en archivo).
  buildM365Runtime(context).catch(() => { /* sin sesion aun */ })
  refreshTimer = setInterval(() => {
    refreshSilent(context)
      .then(() => buildM365Runtime(context))
      .then(() => vscode.commands.executeCommand('m365.registerMcpServer', { silent: true }))
      .catch(() => { /* sin sesion; se ignora */ })
  }, 40 * 60 * 1000)
  context.subscriptions.push({ dispose: () => { if (refreshTimer) { clearInterval(refreshTimer) } } })
}

export function deactivate() {
  if (refreshTimer) { clearInterval(refreshTimer) }
}

// ── Helpers de mcp.json ──────────────────────────────────────────
async function readTextIfExists(uri: vscode.Uri): Promise<string> {
  try { return Buffer.from(await vscode.workspace.fs.readFile(uri)).toString('utf8') }
  catch { return '' }
}

async function mcpFileHasM365(uri: vscode.Uri): Promise<boolean> {
  return hasMcpServer(await readTextIfExists(uri), 'microsoft365')
}

async function writeMcpEntry(file: vscode.Uri, entry: McpServerEntry): Promise<void> {
  const merged = mergeMcpServers(await readTextIfExists(file), 'microsoft365', entry)
  await vscode.workspace.fs.createDirectory(vscode.Uri.joinPath(file, '..'))
  await vscode.workspace.fs.writeFile(file, Buffer.from(merged, 'utf8'))
}
