import * as vscode from 'vscode'
import * as path from 'path'

/** Clave de SecretStorage para el token de lectura de minutas de Polibio. */
export const POLIBIO_TOKEN_KEY = 'm365.polibio.token'

export interface M365Runtime {
  serverPath: string
  env: Record<string, string>
  cfgFile: string
  version: string
  polibioOn: boolean
}

/**
 * Escribe `globalStorage/m365-config.json` (solo con el conector opcional de
 * Polibio) y devuelve lo necesario para lanzar el servidor MCP por stdio.
 *
 * El motor de correo/agenda es Outlook de escritorio por COM: NO necesita token
 * ni login, usa el Outlook que ya esta abierto y firmado. Por eso aqui no hay
 * nada de Graph/Entra: solo, si el usuario lo configuro, los datos para leer las
 * minutas de PolibioDesk (URL + anon key de ajustes, token de SecretStorage).
 */
export async function buildM365Runtime(context: vscode.ExtensionContext): Promise<M365Runtime> {
  await vscode.workspace.fs.createDirectory(context.globalStorageUri)

  const c = vscode.workspace.getConfiguration('m365')
  const polUrl = (c.get<string>('polibio.supabaseUrl', '') || '').trim().replace(/\/+$/, '')
  const polAnon = (c.get<string>('polibio.anonKey', '') || '').trim()
  const polToken = (await context.secrets.get(POLIBIO_TOKEN_KEY)) || ''
  const polibio = (polUrl && polAnon && polToken)
    ? { functionUrl: `${polUrl}/functions/v1/minutas-read`, anonKey: polAnon, token: polToken }
    : null

  const cfg = { polibio }
  const cfgFileUri = vscode.Uri.joinPath(context.globalStorageUri, 'm365-config.json')
  await vscode.workspace.fs.writeFile(cfgFileUri, Buffer.from(JSON.stringify(cfg, null, 2), 'utf8'))

  const serverPath = context.asAbsolutePath(path.join('mcp', 'index.mjs'))
  const env: Record<string, string> = { M365_CONFIG_FILE: cfgFileUri.fsPath }
  const version = (context.extension?.packageJSON?.version as string) || '0.0.0'

  return { serverPath, env, cfgFile: cfgFileUri.fsPath, version, polibioOn: !!polibio }
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
    const rt = await buildM365Runtime(this.context)
    const Stdio: any = (vscode as any).McpStdioServerDefinition
    const def = new Stdio('Microsoft 365', 'node', [rt.serverPath], rt.env, rt.version)
    def.cwd = vscode.Uri.file(path.dirname(rt.serverPath))
    return [def]
  }

  async resolveMcpServerDefinition(server: any): Promise<any> {
    await buildM365Runtime(this.context)
    return server
  }
}

export function registerM365McpProvider(context: vscode.ExtensionContext): M365McpProvider | undefined {
  const lm: any = (vscode as any).lm
  if (!lm || typeof lm.registerMcpServerDefinitionProvider !== 'function') { return undefined }
  const provider = new M365McpProvider(context)
  context.subscriptions.push(lm.registerMcpServerDefinitionProvider('m365.mcpServers', provider))
  return provider
}
