import * as vscode from 'vscode'
import * as path from 'path'

export interface M365Runtime {
  serverPath: string
  env: Record<string, string>
  version: string
}

/**
 * Devuelve lo necesario para lanzar el servidor MCP por stdio. El motor es
 * Outlook de escritorio por COM: no necesita token, login ni configuracion, usa
 * el Outlook que ya esta abierto y firmado.
 */
export function buildM365Runtime(context: vscode.ExtensionContext): M365Runtime {
  const serverPath = context.asAbsolutePath(path.join('mcp', 'index.mjs'))
  const version = (context.extension?.packageJSON?.version as string) || '0.0.0'
  return { serverPath, env: {}, version }
}

/**
 * Provider nativo de MCP para VS Code (API estable desde 1.101). Publica el
 * servidor para que Copilot Chat lo descubra solo, con feature-detect via `any`
 * para mantener engines.vscode en ^1.96 (fallback por comando + mcp.json).
 */
export class M365McpProvider {
  private readonly _onDidChange = new vscode.EventEmitter<void>()
  readonly onDidChangeMcpServerDefinitions = this._onDidChange.event

  constructor(private readonly context: vscode.ExtensionContext) {}

  refresh(): void { this._onDidChange.fire() }

  async provideMcpServerDefinitions(): Promise<any[]> {
    const rt = buildM365Runtime(this.context)
    const Stdio: any = (vscode as any).McpStdioServerDefinition
    const def = new Stdio('Outlook', 'node', [rt.serverPath], rt.env, rt.version)
    def.cwd = vscode.Uri.file(path.dirname(rt.serverPath))
    return [def]
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
