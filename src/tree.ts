import * as vscode from 'vscode'
import { buildM365Runtime } from './mcpProvider'

/**
 * Panel del icono "Microsoft 365" en la barra de actividad. El motor es Outlook
 * de escritorio (COM), asi que no hay login: solo estado y acciones.
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

    const engine = new vscode.TreeItem('Motor: Outlook de escritorio (sin login)')
    engine.iconPath = new vscode.ThemeIcon(process.platform === 'win32' ? 'pass-filled' : 'warning')
    engine.tooltip = process.platform === 'win32'
      ? 'Usa el Outlook que ya tienes abierto y firmado. No necesita login ni permisos de Entra.'
      : 'El motor COM requiere Windows con Outlook de escritorio.'
    items.push(engine)

    let polOn = false
    try { polOn = (await buildM365Runtime(this.context)).polibioOn } catch { /* ignore */ }
    const pol = new vscode.TreeItem(polOn ? 'Conector Polibio: activo' : 'Conector Polibio: no configurado')
    pol.iconPath = new vscode.ThemeIcon(polOn ? 'pass-filled' : 'circle-outline')
    items.push(pol)

    const action = (label: string, icon: string, command: string, tooltip?: string) => {
      const t = new vscode.TreeItem(label)
      t.iconPath = new vscode.ThemeIcon(icon)
      t.command = { command, title: label }
      if (tooltip) { t.tooltip = tooltip }
      items.push(t)
    }

    action('Registrar servidor MCP', 'server-process', 'm365.registerMcpServer',
      'Publica el servidor para Copilot Chat.')
    action('Estado', 'info', 'm365.status')
    action('Guardar token de Polibio (minutas)', 'key', 'm365.setPolibioToken',
      'Activa el conector con el Anotador de PolibioDesk.')

    return items
  }
}
