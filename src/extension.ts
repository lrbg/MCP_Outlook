import * as vscode from 'vscode'
import {
  McpServerEntry, McpFileCorruptError,
  userMcpPathFromGlobalStorage, mergeMcpServers, hasMcpServer,
} from './mcpConfig'
import { buildM365Runtime, registerM365McpProvider, M365McpProvider } from './mcpProvider'
import { M365Tree } from './tree'

let mcpProvider: M365McpProvider | undefined
let tree: M365Tree | undefined

export function activate(context: vscode.ExtensionContext) {
  mcpProvider = registerM365McpProvider(context)

  tree = new M365Tree(context)
  context.subscriptions.push(vscode.window.registerTreeDataProvider('m365View', tree))

  const reg = (id: string, fn: (...a: any[]) => any) =>
    context.subscriptions.push(vscode.commands.registerCommand(id, fn))

  reg('m365.status', async () => {
    vscode.window.showInformationMessage(
      process.platform === 'win32'
        ? 'Outlook MCP usa tu Outlook de escritorio (sin login). Registra el MCP y presiona Start en Copilot Chat.'
        : 'AVISO: este plugin usa Outlook de escritorio por COM y requiere Windows con Outlook.',
    )
  })

  reg('m365.registerMcpServer', async (opts?: { silent?: boolean }) => {
    const silent = !!opts?.silent
    const c = vscode.workspace.getConfiguration('m365')

    const rt = buildM365Runtime(context)
    const entry: McpServerEntry = { command: 'node', args: [rt.serverPath], env: rt.env }

    const ws = vscode.workspace.workspaceFolders?.[0]
    const wsFile = ws ? vscode.Uri.joinPath(ws.uri, '.vscode', 'mcp.json') : undefined
    const userFile = vscode.Uri.file(userMcpPathFromGlobalStorage(context.globalStorageUri.fsPath))

    if (silent) {
      for (const f of [userFile, wsFile]) {
        if (f && await mcpFileHasOutlook(f)) {
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
          { title: 'Outlook: donde registro el servidor MCP?', ignoreFocusOut: true },
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
      tree?.refresh()
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
}

export function deactivate() { /* nada que limpiar */ }

// ── Helpers de mcp.json ──────────────────────────────────────────
async function readTextIfExists(uri: vscode.Uri): Promise<string> {
  try { return Buffer.from(await vscode.workspace.fs.readFile(uri)).toString('utf8') }
  catch { return '' }
}

async function mcpFileHasOutlook(uri: vscode.Uri): Promise<boolean> {
  return hasMcpServer(await readTextIfExists(uri), 'outlook')
}

async function writeMcpEntry(file: vscode.Uri, entry: McpServerEntry): Promise<void> {
  const merged = mergeMcpServers(await readTextIfExists(file), 'outlook', entry)
  await vscode.workspace.fs.createDirectory(vscode.Uri.joinPath(file, '..'))
  await vscode.workspace.fs.writeFile(file, Buffer.from(merged, 'utf8'))
}
