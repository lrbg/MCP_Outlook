import * as vscode from 'vscode'
import { refreshSilent } from './auth'

/**
 * Panel del icono "Microsoft 365" en la barra de actividad. Muestra el estado de
 * la sesion y una lista de acciones que ejecutan los comandos del plugin al
 * hacer clic (mismo espiritu que el explorador de GenRocket).
 */
export class M365Tree implements vscode.TreeDataProvider<vscode.TreeItem> {
  private _onDidChange = new vscode.EventEmitter<vscode.TreeItem | undefined | void>()
  readonly onDidChangeTreeData = this._onDidChange.event

  constructor(private readonly context: vscode.ExtensionContext) {}

  refresh(): void { this._onDidChange.fire() }

  getTreeItem(el: vscode.TreeItem): vscode.TreeItem { return el }

  async getChildren(el?: vscode.TreeItem): Promise<vscode.TreeItem[]> {
    if (el) { return [] }

    const items: vscode.TreeItem[] = []

    // Estado de la sesion (sin escribir config; solo consulta silenciosa).
    let session
    try { session = await refreshSilent(this.context) } catch { /* sin sesion */ }
    if (session) {
      const s = new vscode.TreeItem(`Conectado: ${session.account}`)
      s.iconPath = new vscode.ThemeIcon('pass-filled')
      s.tooltip = `Permisos concedidos: ${session.scopes.join(', ') || '(ninguno)'}`
      items.push(s)
    } else {
      const s = new vscode.TreeItem('Sin sesion de Microsoft')
      s.iconPath = new vscode.ThemeIcon('warning')
      s.tooltip = 'Usa "Iniciar sesion" para conectar tu cuenta de organizacion.'
      items.push(s)
    }

    const action = (label: string, icon: string, command: string, tooltip?: string) => {
      const t = new vscode.TreeItem(label)
      t.iconPath = new vscode.ThemeIcon(icon)
      t.command = { command, title: label }
      if (tooltip) { t.tooltip = tooltip }
      items.push(t)
    }

    action('Iniciar sesion (navegador)', 'sign-in', 'm365.signIn',
      'Login por navegador (recomendado): pasa Conditional Access en equipos unidos a Entra.')
    action('Iniciar sesion (codigo de dispositivo)', 'device-camera', 'm365.signInDeviceCode',
      'Alterno: ingresa un codigo en microsoft.com/devicelogin.')
    action('Registrar servidor MCP', 'server-process', 'm365.registerMcpServer',
      'Publica el servidor para Copilot Chat.')
    action('Refrescar sesion', 'refresh', 'm365.refresh')
    action('Estado y permisos', 'info', 'm365.status')
    action('Guardar token de Polibio (minutas)', 'key', 'm365.setPolibioToken',
      'Activa el conector con el Anotador de PolibioDesk.')

    return items
  }
}
