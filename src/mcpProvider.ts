import * as vscode from 'vscode'
import * as path from 'path'

export interface ServerDef {
  /** Nombre estable para mcp.json. */
  name: string
  /** Etiqueta visible. */
  label: string
  command: string
  args: string[]
  env: Record<string, string>
  /** cwd sugerido (para los servidores locales node). */
  cwd?: string
}

/** Carpeta de recetas: ajuste del usuario o, por defecto, globalStorage/recipes. */
export function recipesDir(context: vscode.ExtensionContext): string {
  const cfg = (vscode.workspace.getConfiguration('m365').get<string>('recipesDir', '') || '').trim()
  return cfg || vscode.Uri.joinPath(context.globalStorageUri, 'recipes').fsPath
}

/**
 * Servidores MCP que publica el plugin:
 *  - outlook: tu Outlook de escritorio por COM (mcp/index.mjs).
 *  - outlook-recipes: recetas de sitios (mcp/recipes.mjs) para dar contexto al agente.
 *  - playwright: manos de navegador (Playwright MCP oficial via npx), opcional.
 */
export function buildServers(context: vscode.ExtensionContext): ServerDef[] {
  const c = vscode.workspace.getConfiguration('m365')
  const outlookPath = context.asAbsolutePath(path.join('mcp', 'index.mjs'))
  const recipesPath = context.asAbsolutePath(path.join('mcp', 'recipes.mjs'))
  const mcpDir = path.dirname(outlookPath)

  const servers: ServerDef[] = [
    { name: 'outlook', label: 'Outlook', command: 'node', args: [outlookPath], env: {}, cwd: mcpDir },
    { name: 'outlook-recipes', label: 'Recetas', command: 'node', args: [recipesPath], env: { RECIPES_DIR: recipesDir(context) }, cwd: mcpDir },
  ]

  if (c.get<boolean>('playwright.enabled', true)) {
    const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx'
    servers.push({ name: 'playwright', label: 'Playwright', command: npx, args: ['-y', '@playwright/mcp@latest'], env: {} })
  }
  return servers
}

/**
 * Provider nativo de MCP (API estable desde 1.101). Publica todos los servidores
 * para que Copilot Chat los descubra solos. Feature-detect via `any` para
 * mantener engines.vscode en ^1.96 (fallback por comando + mcp.json).
 */
export class M365McpProvider {
  private readonly _onDidChange = new vscode.EventEmitter<void>()
  readonly onDidChangeMcpServerDefinitions = this._onDidChange.event

  constructor(private readonly context: vscode.ExtensionContext) {}

  refresh(): void { this._onDidChange.fire() }

  async provideMcpServerDefinitions(): Promise<any[]> {
    const Stdio: any = (vscode as any).McpStdioServerDefinition
    const version = (this.context.extension?.packageJSON?.version as string) || '0.0.0'
    return buildServers(this.context).map((s) => {
      const def = new Stdio(s.label, s.command, s.args, s.env, version)
      if (s.cwd) { def.cwd = vscode.Uri.file(s.cwd) }
      return def
    })
  }

  async resolveMcpServerDefinition(server: any): Promise<any> { return server }
}

export function registerM365McpProvider(context: vscode.ExtensionContext): M365McpProvider | undefined {
  const lm: any = (vscode as any).lm
  if (!lm || typeof lm.registerMcpServerDefinitionProvider !== 'function') { return undefined }
  const provider = new M365McpProvider(context)
  context.subscriptions.push(lm.registerMcpServerDefinitionProvider('m365.mcpServers', provider))
  return provider
}
