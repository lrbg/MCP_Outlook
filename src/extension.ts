import * as vscode from 'vscode'
import {
  McpServerEntry, McpFileCorruptError,
  userMcpPathFromGlobalStorage, mergeMcpServers, hasMcpServer,
} from './mcpConfig'
import { buildServers, ServerDef, recipesDir, registerM365McpProvider, M365McpProvider } from './mcpProvider'
import { M365Tree } from './tree'
import { openDashboard, refresh as refreshDashboard } from './dashboard'
import { runDailyReview, maybeRunScheduled } from './dailyReview'

let mcpProvider: M365McpProvider | undefined
let tree: M365Tree | undefined
let reviewTimer: ReturnType<typeof setInterval> | undefined

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

  reg('m365.openDashboard', () => openDashboard(context))

  reg('m365.runDailyReview', async () => {
    try {
      const e = await runDailyReview(context)
      await refreshDashboard(context)
      tree?.refresh()
      vscode.window.showInformationMessage(`Revision de bandeja lista: ${e.unreadCount} no leidos.`)
    } catch (err: any) {
      vscode.window.showErrorMessage(`No se pudo correr la revision: ${err?.message || err}`)
    }
  })

  // Programador: intenta al arrancar y cada 30 min (corre mientras VS Code este
  // abierto; si a la hora configurada estaba cerrado, se pone al dia al abrir).
  const tick = () => maybeRunScheduled(context, () => { refreshDashboard(context); tree?.refresh() })
  tick()
  reviewTimer = setInterval(tick, 30 * 60 * 1000)
  context.subscriptions.push({ dispose: () => { if (reviewTimer) { clearInterval(reviewTimer) } } })

  reg('m365.openRecipes', async () => {
    const dir = vscode.Uri.file(recipesDir(context))
    await vscode.workspace.fs.createDirectory(dir)
    await vscode.env.openExternal(dir)
  })

  reg('m365.registerMcpServer', async (opts?: { silent?: boolean }) => {
    const silent = !!opts?.silent
    const c = vscode.workspace.getConfiguration('m365')

    await ensureRecipesSeed(context)
    const servers = buildServers(context)

    const ws = vscode.workspace.workspaceFolders?.[0]
    const wsFile = ws ? vscode.Uri.joinPath(ws.uri, '.vscode', 'mcp.json') : undefined
    const userFile = vscode.Uri.file(userMcpPathFromGlobalStorage(context.globalStorageUri.fsPath))

    if (silent) {
      for (const f of [userFile, wsFile]) {
        if (f && await mcpFileHasOurs(f)) {
          try { await writeAllServers(f, servers) } catch { /* se avisara al registrar a mano */ }
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
          { title: 'Outlook: donde registro los servidores MCP?', ignoreFocusOut: true },
        )
        if (!pick) { return }
        scope = pick.scope
        await c.update('mcpScope', scope, vscode.ConfigurationTarget.Global)
      }
    }

    const target = (scope === 'workspace' && wsFile) ? wsFile : userFile
    try {
      await writeAllServers(target, servers)
      const donde = target === userFile
        ? 'MCP registrado para todo VS Code.'
        : 'MCP registrado en .vscode/mcp.json de este proyecto.'
      vscode.window.showInformationMessage(
        `${donde} Servidores: ${servers.map(s => s.name).join(', ')}. Presiona "Start" (o "Restart") en Copilot Chat.`,
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

async function mcpFileHasOurs(uri: vscode.Uri): Promise<boolean> {
  return hasMcpServer(await readTextIfExists(uri), 'outlook')
}

/** Escribe (o actualiza) todos nuestros servidores en un mcp.json conservando el resto. */
async function writeAllServers(file: vscode.Uri, servers: ServerDef[]): Promise<void> {
  let text = await readTextIfExists(file)
  for (const s of servers) {
    const entry: McpServerEntry = { command: s.command, args: s.args, env: s.env }
    text = mergeMcpServers(text, s.name, entry)
  }
  await vscode.workspace.fs.createDirectory(vscode.Uri.joinPath(file, '..'))
  await vscode.workspace.fs.writeFile(file, Buffer.from(text, 'utf8'))
}

/** Crea la carpeta de recetas y siembra un ejemplo la primera vez. */
async function ensureRecipesSeed(context: vscode.ExtensionContext): Promise<void> {
  const dir = vscode.Uri.file(recipesDir(context))
  await vscode.workspace.fs.createDirectory(dir)
  const sample = vscode.Uri.joinPath(dir, 'ejemplo-kpi.md')
  try { await vscode.workspace.fs.stat(sample); return } catch { /* no existe: sembrar */ }
  const md = [
    '# KPI en algo.com (ejemplo)',
    '',
    'Receta de ejemplo. Editala o crea las tuyas (una por sitio).',
    '',
    '## Objetivo',
    'Registrar los KPI de un recurso en el portal.',
    '',
    '## Pasos',
    '1. Abre https://www.algo.com e inicia sesion TU (el agente no teclea contrasenas).',
    '2. Ve a Reportes > KPIs > Nuevo.',
    '3. Selecciona el recurso indicado en el correo.',
    '4. Llena los campos: Meta, Avance, Comentario.',
    '5. Revisa el resumen y presiona Guardar (confirma antes de guardar).',
    '',
    '## Notas',
    '- Si el sitio pide 2FA, complétalo tu.',
    '- Los datos exactos (recurso, valores) vienen en el correo que dispara la tarea.',
  ].join('\n')
  await vscode.workspace.fs.writeFile(sample, Buffer.from(md, 'utf8'))
}
