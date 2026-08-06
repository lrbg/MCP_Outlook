import * as vscode from 'vscode'
import { DayEntry, toMarkdown } from './bitacoraCore'
import { loadEntries, runDailyReview } from './dailyReview'
import { getPriorityEmails, isWindows } from './priorityInbox'
import { readEmailBody, sendReply } from './outlookActions'
import { draftReply } from './copilot'
import { recipesDir } from './mcpProvider'

let panel: vscode.WebviewPanel | undefined

export async function openDashboard(context: vscode.ExtensionContext): Promise<void> {
  if (panel) { panel.reveal(); return }
  panel = vscode.window.createWebviewPanel(
    'm365Panel', 'Outlook MCP', vscode.ViewColumn.Active,
    { enableScripts: true, retainContextWhenHidden: true },
  )
  panel.onDidDispose(() => { panel = undefined })
  panel.webview.onDidReceiveMessage((m) => handle(context, m))
  await refresh(context)
}

export async function refresh(context: vscode.ExtensionContext): Promise<void> {
  if (!panel) { return }
  panel.webview.html = render(await gather(context))
}

// ── Estado ────────────────────────────────────────────────────────
interface State {
  entries: DayEntry[]
  enabled: boolean
  hour: number
  maxEmails: number
  senders: string[]
  recipes: string[]
  windows: boolean
}

async function gather(context: vscode.ExtensionContext): Promise<State> {
  const c = vscode.workspace.getConfiguration('m365')
  return {
    entries: await loadEntries(context),
    enabled: c.get('dailyReview.enabled', true),
    hour: c.get('dailyReview.hour', 8),
    maxEmails: c.get('dailyReview.maxEmails', 30),
    senders: c.get<string[]>('prioritySenders', []),
    recipes: await listRecipes(context),
    windows: isWindows,
  }
}

async function listRecipes(context: vscode.ExtensionContext): Promise<string[]> {
  try {
    const dir = vscode.Uri.file(recipesDir(context))
    const items = await vscode.workspace.fs.readDirectory(dir)
    return items.filter(([n]) => n.endsWith('.md')).map(([n]) => n.replace(/\.md$/, '')).sort()
  } catch { return [] }
}

// ── Mensajes ──────────────────────────────────────────────────────
async function handle(context: vscode.ExtensionContext, m: any): Promise<void> {
  const c = vscode.workspace.getConfiguration('m365')
  const G = vscode.ConfigurationTarget.Global
  switch (m?.type) {
    case 'loadPriority': {
      if (!panel) { return }
      try {
        const emails = getPriorityEmails(c.get<string[]>('prioritySenders', []), 14, 20)
        panel.webview.postMessage({ type: 'priority', emails })
      } catch (e: any) {
        panel.webview.postMessage({ type: 'priority', error: e?.message || String(e) })
      }
      return
    }
    case 'setSetting': {
      await c.update(m.key, m.value, G)
      return
    }
    case 'addSender': {
      const email = String(m.email || '').trim().toLowerCase()
      if (!email) { return }
      const cur = c.get<string[]>('prioritySenders', [])
      if (!cur.includes(email)) { await c.update('prioritySenders', [...cur, email], G) }
      await refresh(context)
      return
    }
    case 'removeSender': {
      const cur = c.get<string[]>('prioritySenders', [])
      await c.update('prioritySenders', cur.filter(s => s !== m.email), G)
      await refresh(context)
      return
    }
    case 'runReview': {
      try { await runDailyReview(context); vscode.window.showInformationMessage('Revisión actualizada.') }
      catch (e: any) { vscode.window.showErrorMessage(`No se pudo actualizar: ${e?.message || e}`) }
      await refresh(context)
      return
    }
    case 'draftReply': {
      if (!panel) { return }
      try {
        const email = readEmailBody(String(m.id))
        const body = await vscode.window.withProgress(
          { location: vscode.ProgressLocation.Notification, title: 'Copilot está redactando la respuesta…' },
          (_p, token) => draftReply(email, token, String(m.instruction || '')),
        )
        panel.webview.postMessage({ type: 'replyDraft', id: m.id, sender: email.sender, subject: email.subject, body })
      } catch (e: any) {
        panel.webview.postMessage({ type: 'replyDraft', id: m.id, error: e?.message || String(e) })
      }
      return
    }
    case 'sendReply': {
      const ok = await vscode.window.showWarningMessage(
        `¿Enviar esta respuesta a ${m.sender || 'el remitente'}?`, { modal: true }, 'Enviar',
      )
      if (ok !== 'Enviar') { panel?.webview.postMessage({ type: 'replyCancelled', id: m.id }); return }
      try {
        sendReply(String(m.id), String(m.body || ''), !!m.replyAll)
        vscode.window.showInformationMessage('Respuesta enviada.')
        panel?.webview.postMessage({ type: 'replySent', id: m.id })
      } catch (e: any) {
        vscode.window.showErrorMessage(`No se pudo enviar: ${e?.message || e}`)
        panel?.webview.postMessage({ type: 'replyCancelled', id: m.id })
      }
      return
    }
    case 'openRecipe': {
      const uri = vscode.Uri.joinPath(vscode.Uri.file(recipesDir(context)), `${m.name}.md`)
      vscode.window.showTextDocument(uri)
      return
    }
    case 'openRecipes': vscode.commands.executeCommand('m365.openRecipes'); return
    case 'registerServers': vscode.commands.executeCommand('m365.registerMcpServer'); return
    case 'exportMd': {
      const uri = await vscode.window.showSaveDialog({ filters: { Markdown: ['md'] }, defaultUri: vscode.Uri.file('bitacora-bandeja.md') })
      if (uri) {
        await vscode.workspace.fs.writeFile(uri, Buffer.from(toMarkdown(await loadEntries(context)), 'utf8'))
        vscode.window.showInformationMessage(`Bitácora guardada: ${uri.fsPath}`)
      }
      return
    }
  }
}

// ── Render ────────────────────────────────────────────────────────
function esc(s: string): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function mdLite(md: string): string {
  const out: string[] = []
  let inList = false
  for (const raw of esc(md || '').split('\n')) {
    const line = raw.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    const li = line.match(/^\s*[-*]\s+(.*)$/)
    const h = line.match(/^\s*#{1,4}\s+(.*)$/)
    if (li) { if (!inList) { out.push('<ul>'); inList = true } out.push(`<li>${li[1]}</li>`); continue }
    if (inList) { out.push('</ul>'); inList = false }
    if (h) { out.push(`<h4>${h[1]}</h4>`) }
    else if (line.trim() === '') { out.push('') }
    else { out.push(`<p>${line}</p>`) }
  }
  if (inList) { out.push('</ul>') }
  return out.join('\n')
}

function render(s: State): string {
  const senderChips = s.senders.length
    ? s.senders.map(e => `<span class="chip">${esc(e)}<button class="x" title="Quitar" onclick="post('removeSender',{email:'${esc(e)}'})">&times;</button></span>`).join('')
    : '<span class="muted">Sin remitentes. Agrega uno abajo.</span>'

  const bitacoraRows = s.entries.slice(0, 8).map(e =>
    `<div class="row"><span>${esc(e.date)}</span><span class="muted">${e.unreadCount} no leídos</span></div>`).join('')
    || '<p class="muted">Aún no hay revisiones.</p>'

  const notes = s.entries.slice(0, 3).map(e =>
    `<details class="day"><summary>${esc(e.date)} · ${e.unreadCount} no leídos</summary><div class="notes">${mdLite(e.notesMarkdown)}</div></details>`).join('')

  const recipes = s.recipes.length
    ? s.recipes.map(n => `<div class="row link" onclick="post('openRecipe',{name:'${esc(n)}'})"><span><i>&#9776;</i> ${esc(n)}</span><span class="muted">&rsaquo;</span></div>`).join('')
    : '<p class="muted">Sin recetas todavía.</p>'

  const winWarn = s.windows ? '' : '<div class="warn">Este panel usa Outlook de escritorio (COM) y requiere Windows.</div>'

  return `<!doctype html><html><head><meta charset="utf-8"><style>
  :root { --gap: 16px; }
  body { font-family: var(--vscode-font-family); color: var(--vscode-foreground); padding: 18px 22px; font-size: 13px; }
  h1 { font-size: 20px; margin: 0; font-weight: 600; }
  h2 { font-size: 14px; margin: 0; font-weight: 600; }
  .sub { color: var(--vscode-descriptionForeground); margin: 3px 0 0; }
  .muted { color: var(--vscode-descriptionForeground); }
  .head { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; flex-wrap: wrap; margin-bottom: var(--gap); }
  .chips { display: flex; gap: 6px; flex-wrap: wrap; }
  .pill { display: inline-flex; align-items: center; gap: 6px; font-size: 11px; padding: 4px 10px; border-radius: 999px;
    background: var(--vscode-badge-background); color: var(--vscode-badge-foreground); }
  .dot { width: 7px; height: 7px; border-radius: 50%; background: #3fb950; display: inline-block; }
  .card { background: var(--vscode-editorWidget-background, var(--vscode-editor-background)); border: 1px solid var(--vscode-panel-border);
    border-radius: 10px; padding: 14px 16px; margin-bottom: var(--gap); }
  .card h2 { display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }
  .grid2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: var(--gap); }
  .formgrid { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; align-items: end; }
  label { display: block; font-size: 11px; color: var(--vscode-descriptionForeground); margin-bottom: 4px; }
  input[type=number], input[type=text] { width: 100%; box-sizing: border-box; padding: 6px 8px; border-radius: 5px;
    background: var(--vscode-input-background); color: var(--vscode-input-foreground); border: 1px solid var(--vscode-input-border, transparent); }
  button { background: var(--vscode-button-background); color: var(--vscode-button-foreground); border: none;
    padding: 6px 12px; border-radius: 5px; cursor: pointer; font-size: 12px; }
  button.sec { background: var(--vscode-button-secondaryBackground); color: var(--vscode-button-secondaryForeground); }
  button:hover { opacity: .9; }
  .switch { display: inline-flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px; }
  .switch input { width: 16px; height: 16px; }
  .row { display: flex; justify-content: space-between; align-items: center; padding: 7px 0; border-bottom: 1px solid var(--vscode-panel-border); }
  .row:last-child { border-bottom: none; }
  .link { cursor: pointer; } .link:hover { color: var(--vscode-textLink-foreground); }
  .chip { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; padding: 3px 6px 3px 10px; border-radius: 999px;
    background: var(--vscode-badge-background); color: var(--vscode-badge-foreground); }
  .chip .x { background: none; color: inherit; padding: 0 2px; font-size: 14px; line-height: 1; }
  .senderbox { display: flex; gap: 8px; margin-top: 10px; }
  .mailrow { display: flex; flex-direction: column; padding: 8px 0; border-bottom: 1px solid var(--vscode-panel-border); }
  .mailrow:last-child { border-bottom: none; }
  .mailrow .top { display: flex; justify-content: space-between; gap: 10px; }
  .mailrow .from { font-weight: 600; } .mailrow .subj { color: var(--vscode-descriptionForeground); margin-top: 2px; }
  .mailrow.unread .from::before { content: "\\2022  "; color: var(--vscode-textLink-foreground); }
  details.day { border-top: 1px solid var(--vscode-panel-border); padding: 8px 0; }
  details.day summary { cursor: pointer; font-weight: 600; }
  .notes h4 { margin: 8px 0 3px; font-size: 12px; } .notes p { margin: 3px 0; } .notes ul { margin: 3px 0 3px 16px; }
  .warn { background: var(--vscode-inputValidation-warningBackground, #5a4d00); border: 1px solid var(--vscode-inputValidation-warningBorder, #806f00);
    padding: 8px 12px; border-radius: 6px; margin-bottom: var(--gap); }
  .toolbar { display: flex; gap: 8px; }
  .mailrow .actions { margin-top: 6px; }
  .mailrow .actions button { padding: 3px 10px; font-size: 11px; }
  #composer .composer { margin-top: 14px; border: 1px solid var(--vscode-focusBorder, var(--vscode-panel-border)); border-radius: 8px; padding: 12px; }
  #composer textarea { width: 100%; box-sizing: border-box; min-height: 160px; resize: vertical; font-family: var(--vscode-font-family);
    padding: 8px; border-radius: 5px; background: var(--vscode-input-background); color: var(--vscode-input-foreground); border: 1px solid var(--vscode-input-border, transparent); }
  #composer .crow { display: flex; align-items: center; gap: 10px; margin-top: 10px; flex-wrap: wrap; }
</style></head><body>
  <div class="head">
    <div><h1>Outlook MCP</h1><p class="sub">Correo y agenda para tu agente, sin login. Motor: Outlook de escritorio.</p></div>
    <div class="chips">
      <span class="pill"><span class="dot"></span>Outlook</span>
      <span class="pill"><span class="dot"></span>Recetas</span>
      <span class="pill"><span class="dot"></span>Playwright</span>
    </div>
  </div>
  ${winWarn}

  <div class="card">
    <h2>Revisión diaria de bandeja</h2>
    <div class="formgrid">
      <label class="switch">Activada<input type="checkbox" ${s.enabled ? 'checked' : ''} onchange="post('setSetting',{key:'dailyReview.enabled',value:this.checked})"></label>
      <div><label>Hora de corrida</label><input type="number" min="0" max="23" value="${s.hour}" onchange="post('setSetting',{key:'dailyReview.hour',value:+this.value})"></div>
      <div><label>Correos a revisar</label><input type="number" min="1" max="100" value="${s.maxEmails}" onchange="post('setSetting',{key:'dailyReview.maxEmails',value:+this.value})"></div>
      <button onclick="post('runReview')">Correr ahora</button>
    </div>
    <p class="sub" style="margin-top:10px">La resume tu Copilot. Corre sola cada día mientras VS Code esté abierto.</p>
  </div>

  <div class="card">
    <h2>Correos prioritarios <span class="muted" style="font-weight:400">· últimas 2 semanas</span></h2>
    <div class="chips">${senderChips}</div>
    <div class="senderbox">
      <input type="text" id="newSender" placeholder="nombre@empresa.com" onkeydown="if(event.key==='Enter')addSender()">
      <button class="sec" onclick="addSender()">Agregar</button>
      <button class="sec" onclick="loadPriority()">Actualizar</button>
    </div>
    <div id="priority" style="margin-top:12px"><p class="muted">Cargando correos prioritarios…</p></div>
    <div id="composer"></div>
  </div>

  <div class="grid2">
    <div class="card">
      <h2>Bitácora <span style="flex:1"></span></h2>
      ${bitacoraRows}
      ${notes}
      <div class="toolbar" style="margin-top:12px"><button class="sec" onclick="post('exportMd')">Guardar .md</button></div>
    </div>
    <div class="card">
      <h2>Recetas de sitios <span style="flex:1"></span><button class="sec" style="padding:3px 10px" onclick="post('openRecipes')">+ Carpeta</button></h2>
      ${recipes}
    </div>
  </div>

  <div class="card" style="display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap">
    <span class="muted">3 servidores MCP listos para Copilot Chat</span>
    <button onclick="post('registerServers')">Registrar servidores</button>
  </div>

  <script>
    const vscode = acquireVsCodeApi();
    const store = {};
    function post(type, extra){ vscode.postMessage(Object.assign({type}, extra||{})); }
    function addSender(){ const el=document.getElementById('newSender'); const v=(el.value||'').trim(); if(v){ post('addSender',{email:v}); } }
    function loadPriority(){ document.getElementById('priority').innerHTML='<p class="muted">Cargando…</p>'; document.getElementById('composer').innerHTML=''; post('loadPriority'); }
    function esc(s){ return String(s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

    function reply(id){ document.getElementById('composer').innerHTML='<div class="composer"><p class="muted">Copilot está redactando…</p></div>'; post('draftReply',{id}); }
    function renderComposer(m){
      const c = document.getElementById('composer');
      if (m.error) { c.innerHTML='<div class="composer"><p class="muted">No se pudo redactar: '+esc(m.error)+'</p></div>'; return; }
      store[m.id] = { sender: m.sender };
      c.innerHTML =
        '<div class="composer"><div style="font-weight:600;margin-bottom:2px">Responder a '+esc(m.sender)+'</div>'+
        '<div class="muted" style="margin-bottom:8px">Re: '+esc(m.subject||'')+' · edítalo antes de enviar</div>'+
        '<textarea id="rbody">'+esc(m.body)+'</textarea>'+
        '<div class="crow"><label class="switch"><input type="checkbox" id="rall"> Responder a todos</label>'+
        '<span style="flex:1"></span>'+
        '<button class="sec" onclick="regen(\\''+m.id+'\\')">Regenerar</button>'+
        '<button class="sec" onclick="document.getElementById(\\'composer\\').innerHTML=\\'\\'">Cancelar</button>'+
        '<button onclick="send(\\''+m.id+'\\')">Enviar</button></div></div>';
      c.scrollIntoView({behavior:'smooth', block:'nearest'});
    }
    function regen(id){ const t=document.getElementById('rbody'); const instr=(t&&t.dataset.instr)||''; document.getElementById('composer').innerHTML='<div class="composer"><p class="muted">Regenerando…</p></div>'; post('draftReply',{id}); }
    function send(id){ const body=(document.getElementById('rbody')||{}).value||''; const all=(document.getElementById('rall')||{}).checked||false; post('sendReply',{id,body,replyAll:all,sender:(store[id]||{}).sender||''}); }

    window.addEventListener('message', e => {
      const m = e.data;
      if (m.type === 'priority') {
        const box = document.getElementById('priority');
        if (m.error) { box.innerHTML = '<p class="muted">No se pudieron leer los correos: ' + esc(m.error) + '</p>'; return; }
        if (!m.emails || !m.emails.length) { box.innerHTML = '<p class="muted">Sin correos de esos remitentes en las últimas 2 semanas.</p>'; return; }
        box.innerHTML = m.emails.map(x =>
          '<div class="mailrow ' + (x.unread?'unread':'') + '"><div class="top"><span class="from">' + esc(x.sender||x.senderEmail) +
          '</span><span class="muted">' + esc(x.received) + '</span></div><div class="subj">' + esc(x.subject||'(sin asunto)') + '</div>' +
          '<div class="actions"><button class="sec" onclick="reply(\\''+esc(x.id)+'\\')">Responder con IA</button></div></div>'
        ).join('');
      } else if (m.type === 'replyDraft') { renderComposer(m); }
      else if (m.type === 'replySent') { document.getElementById('composer').innerHTML='<div class="composer"><p class="muted">Respuesta enviada.</p></div>'; }
    });
    post('loadPriority');
  </script>
</body></html>`
}
