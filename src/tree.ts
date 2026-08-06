import * as vscode from 'vscode'

/**
 * Panel del icono "Outlook" en la barra de actividad. El motor es Outlook de
 * escritorio (COM): no hay login, solo estado y el registro del MCP.
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

    const action = (label: string, icon: string, command: string, tooltip?: string) => {
      const t = new vscode.TreeItem(label)
      t.iconPath = new vscode.ThemeIcon(icon)
      t.command = { command, title: label }
      if (tooltip) { t.tooltip = tooltip }
      items.push(t)
    }

    action('Abrir panel de control', 'dashboard', 'm365.openDashboard',
      'Revision diaria, correos prioritarios, bitacora y recetas.')
    action('Correr revision ahora', 'play', 'm365.runDailyReview')
    action('Recetas de sitios (Playwright)', 'book', 'm365.openRecipes',
      'Procedimientos para que el agente opere sitios web con manos de navegador.')
    action('Registrar servidores MCP', 'server-process', 'm365.registerMcpServer',
      'Publica Outlook + Recetas + Playwright para Copilot Chat.')
    action('Estado', 'info', 'm365.status')

    return items
  }
}
