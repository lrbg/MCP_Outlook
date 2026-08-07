import * as vscode from 'vscode'

let ch: vscode.OutputChannel | undefined

function chan(): vscode.OutputChannel {
  if (!ch) { ch = vscode.window.createOutputChannel('Outlook MCP') }
  return ch
}

/** Escribe una linea en el canal "Outlook MCP" (View > Output). */
export function log(msg: string): void {
  const t = new Date().toISOString().slice(11, 19)
  chan().appendLine(`${t}  ${msg}`)
}

/** Muestra el canal de registro. */
export function showLog(): void {
  chan().show(true)
}
