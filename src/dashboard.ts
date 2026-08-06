import * as vscode from 'vscode'
import { DayEntry, toMarkdown } from './bitacoraCore'
import { loadEntries, runDailyReview } from './dailyReview'

let panel: vscode.WebviewPanel | undefined

export async function openDashboard(context: vscode.ExtensionContext): Promise<void> {
  if (panel) { panel.reveal(); await refresh(context); return }
  panel = vscode.window.createWebviewPanel(
    'm365Bitacora', 'Bitacora de bandeja', vscode.ViewColumn.Active,
    { enableScripts: true, retainContextWhenHidden: true },
  )
  panel.onDidDispose(() => { panel = undefined })
  panel.webview.onDidReceiveMessage(async (msg) => {
    if (msg?.type === 'refresh') {
      try {
        await runDailyReview(context)
        vscode.window.showInformationMessage('Revision de bandeja actualizada.')
      } catch (e: any) {
        vscode.window.showErrorMessage(`No se pudo actualizar: ${e?.message || e}`)
      }
      await refresh(context)
    } else if (msg?.type === 'exportMd') {
      const entries = await loadEntries(context)
      const uri = await vscode.window.showSaveDialog({
        filters: { Markdown: ['md'] },
        defaultUri: vscode.Uri.file('bitacora-bandeja.md'),
      })
      if (uri) {
        await vscode.workspace.fs.writeFile(uri, Buffer.from(toMarkdown(entries), 'utf8'))
        vscode.window.showInformationMessage(`Bitacora guardada: ${uri.fsPath}`)
      }
    }
  })
  await refresh(context)
}

export async function refresh(context: vscode.ExtensionContext): Promise<void> {
  if (!panel) { return }
  const entries = await loadEntries(context)
  panel.webview.html = render(entries)
}

// ── Render ────────────────────────────────────────────────────────
function esc(s: string): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Markdown-lite -> HTML: encabezados, negritas, vinetas y saltos. */
function mdLite(md: string): string {
  const lines = esc(md || '').split('\n')
  const out: string[] = []
  let inList = false
  for (let raw of lines) {
    let line = raw.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>').replace(/`(.+?)`/g, '<code>$1</code>')
    const li = line.match(/^\s*[-*]\s+(.*)$/)
    const h = line.match(/^\s*(#{1,4})\s+(.*)$/)
    if (li) {
      if (!inList) { out.push('<ul>'); inList = true }
      out.push(`<li>${li[1]}</li>`)
      continue
    }
    if (inList) { out.push('</ul>'); inList = false }
    if (h) { out.push(`<h4>${h[2]}</h4>`) }
    else if (line.trim() === '') { out.push('') }
    else { out.push(`<p>${line}</p>`) }
  }
  if (inList) { out.push('</ul>') }
  return out.join('\n')
}

function render(entries: DayEntry[]): string {
  const rows = entries.map(e => `
    <tr>
      <td class="date">${esc(e.date)}</td>
      <td class="num">${e.unreadCount}</td>
      <td>${e.keySenders.map(s => `${esc(s.name)} <span class="dim">(${s.count})</span>`).join(', ') || '-'}</td>
    </tr>`).join('')

  const notes = entries.map(e => `
    <section class="day">
      <h3>${esc(e.date)} <span class="dim">· ${e.unreadCount} no leidos · corrida ${esc(e.ranAt)}</span></h3>
      <div class="notes">${mdLite(e.notesMarkdown)}</div>
    </section>`).join('')

  const empty = entries.length === 0
    ? '<p class="dim">Aun no hay revisiones. Presiona "Actualizar ahora" o espera a la corrida automatica del dia.</p>'
    : ''

  return `<!doctype html><html><head><meta charset="utf-8">
<style>
  body { font-family: var(--vscode-font-family); color: var(--vscode-foreground); padding: 16px 20px; }
  h1 { font-size: 1.3rem; margin: 0 0 4px; }
  .sub { color: var(--vscode-descriptionForeground); margin: 0 0 16px; }
  .toolbar { margin: 0 0 16px; display: flex; gap: 8px; }
  button { background: var(--vscode-button-background); color: var(--vscode-button-foreground);
    border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; }
  button.secondary { background: var(--vscode-button-secondaryBackground); color: var(--vscode-button-secondaryForeground); }
  table { border-collapse: collapse; width: 100%; margin: 0 0 24px; }
  th, td { text-align: left; padding: 6px 10px; border-bottom: 1px solid var(--vscode-panel-border); vertical-align: top; }
  th { color: var(--vscode-descriptionForeground); font-weight: 600; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  td.date { white-space: nowrap; }
  .dim { color: var(--vscode-descriptionForeground); }
  section.day { border-top: 1px solid var(--vscode-panel-border); padding: 12px 0; }
  section.day h3 { font-size: 1rem; margin: 0 0 6px; }
  .notes h4 { margin: 10px 0 4px; font-size: 0.95rem; }
  .notes p { margin: 4px 0; } .notes ul { margin: 4px 0 4px 18px; } .notes li { margin: 2px 0; }
  code { background: var(--vscode-textCodeBlock-background); padding: 0 4px; border-radius: 3px; }
</style></head><body>
  <h1>Bitacora de bandeja</h1>
  <p class="sub">Revision diaria de tus correos no leidos, resumida por Copilot.</p>
  <div class="toolbar">
    <button onclick="post('refresh')">Actualizar ahora</button>
    <button class="secondary" onclick="post('exportMd')">Guardar .md</button>
  </div>
  ${empty}
  ${entries.length ? `<table><thead><tr><th>Fecha</th><th>No leidos</th><th>Remitentes clave</th></tr></thead><tbody>${rows}</tbody></table>` : ''}
  ${notes}
  <script>
    const vscode = acquireVsCodeApi();
    function post(type){ vscode.postMessage({ type }); }
  </script>
</body></html>`
}
