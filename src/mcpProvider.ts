import * as vscode from 'vscode'
import * as path from 'path'
import { refreshSilent, M365Session } from './auth'

export interface M365Runtime {
  /** Ruta absoluta al entrypoint del servidor MCP (mcp/index.mjs). */
  serverPath: string
  /** Variables de entorno para lanzar el server (solo apunta al config). */
  env: Record<string, string>
  /** Ruta del archivo de config local escrito en globalStorage. */
  cfgFile: string
  /** true si hay una sesion de Microsoft con token. */
  hasSession: boolean
  /** Permisos concedidos (nombres cortos de Graph). */
  scopes: string[]
  /** Version de la extension. */
  version: string
}

/**
 * Escribe `globalStorage/m365-config.json` con el token de Graph de la sesion de
 * Microsoft de VS Code y los permisos concedidos, y devuelve lo necesario para
 * lanzar el servidor MCP por stdio. El token vive SOLO en globalStorage, nunca
 * en el repo. Refresca en silencio (sin abrir dialogos): si no hay sesion, el
 * config queda sin token y las herramientas responderan con un aviso claro.
 */
export async function buildM365Runtime(context: vscode.ExtensionContext): Promise<M365Runtime> {
  await vscode.workspace.fs.createDirectory(context.globalStorageUri)

  let session: M365Session | undefined
  try { session = await refreshSilent(context) } catch { /* sin sesion todavia */ }

  const cfg = {
    accessToken: session?.accessToken || '',
    scopes: session?.scopes || [],
    account: session?.account || '',
    expiresOn: session?.expiresOn || 0,
    graphBase: 'https://graph.microsoft.com/v1.0',
  }
  const cfgFileUri = vscode.Uri.joinPath(context.globalStorageUri, 'm365-config.json')
  await vscode.workspace.fs.writeFile(cfgFileUri, Buffer.from(JSON.stringify(cfg, null, 2), 'utf8'))

  const serverPath = context.asAbsolutePath(path.join('mcp', 'index.mjs'))
  const env: Record<string, string> = { M365_CONFIG_FILE: cfgFileUri.fsPath }
  const version = (context.extension?.packageJSON?.version as string) || '0.0.0'

  return {
    serverPath,
    env,
    cfgFile: cfgFileUri.fsPath,
    hasSession: !!session?.accessToken,
    scopes: session?.scopes || [],
    version,
  }
}

/**
 * Provider nativo de MCP para VS Code (API `vscode.lm.registerMcpServerDefinitionProvider`,
 * estable desde 1.101). Publica el servidor para que Copilot Chat lo descubra solo.
 * En cada descubrimiento reescribe el config, asi el server arranca con token fresco.
 *
 * Se accede a la API via `any` + feature-detect para mantener engines.vscode en ^1.96
 * (en editores viejos se usa el fallback por comando + mcp.json).
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
    const rt = await buildM365Runtime(this.context)
    if (!rt.hasSession) {
      vscode.window.showWarningMessage(
        'Microsoft 365 MCP: no hay sesion. Ejecuta "M365: Iniciar sesion (Microsoft)".',
      )
    }
    return server
  }
}

/**
 * Registra el provider nativo si el editor soporta la API (VS Code 1.101+).
 * Devuelve el provider (para refrescarlo al iniciar/refrescar sesion) o undefined
 * si el editor es viejo (se usa el fallback por comando + mcp.json).
 */
export function registerM365McpProvider(context: vscode.ExtensionContext): M365McpProvider | undefined {
  const lm: any = (vscode as any).lm
  if (!lm || typeof lm.registerMcpServerDefinitionProvider !== 'function') { return undefined }
  const provider = new M365McpProvider(context)
  context.subscriptions.push(lm.registerMcpServerDefinitionProvider('m365.mcpServers', provider))
  return provider
}
