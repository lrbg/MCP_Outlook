import * as vscode from 'vscode'
import { DayEntry, toMarkdown } from './bitacoraCore'
import { loadEntries, runDailyReview } from './dailyReview'
import { getPriorityEmails, isWindows } from './priorityInbox'
import { classifyPriority, labelText } from './priorityClassify'
import { loadStatus, setStatus, clearStatus } from './priorityState'
import { readEmailBody, sendReply, getMe } from './outlookActions'
import { draftReply, assistAgenda } from './copilot'
import { getMeetings } from './calendarRead'
import { groupByDay, findConflicts, freeSlotsByDay } from './agendaCore'
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
        const days = Math.min(Math.max(Number(m.days) || 14, 1), 60)
        const max = days <= 1 ? 25 : days <= 7 ? 35 : 50
        const emails = getPriorityEmails(c.get<string[]>('prioritySenders', []), days, max)
        const meRaw = getMe()
        const me = { name: meRaw.name, email: meRaw.email, tokens: c.get<string[]>('mentionTokens', []) }
        const status = await loadStatus(context)
        const enriched = emails.map(e => {
          const cl = classifyPriority(e, me)
          return {
            id: e.id, sender: e.sender, senderEmail: e.senderEmail, subject: e.subject,
            received: e.received, unread: e.unread,
            label: cl.label, labelText: labelText(cl.label), needsAction: cl.needsAction,
            status: status[e.id] || null,
          }
        })
        panel.webview.postMessage({ type: 'priority', emails: enriched, you: me.name })
      } catch (e: any) {
        panel.webview.postMessage({ type: 'priority', error: e?.message || String(e) })
      }
      return
    }
    case 'markStatus': {
      const now = Date.now()
      if (m.status) { await setStatus(context, String(m.id), m.status, now) }
      else { await clearStatus(context, String(m.id)) }
      return
    }
    case 'openEmail': {
      if (!panel) { return }
      try {
        const b = readEmailBody(String(m.id))
        panel.webview.postMessage({
          type: 'emailBody', id: m.id, subject: b.subject, sender: b.sender,
          senderEmail: b.senderEmail, to: b.to, cc: b.cc, received: b.received, body: b.body,
        })
      } catch (e: any) {
        panel.webview.postMessage({ type: 'emailBody', id: m.id, error: e?.message || String(e) })
      }
      return
    }
    case 'loadAgenda': {
      if (!panel) { return }
      try {
        const days = c.get<number>('agenda.days', 7)
        const meetings = getMeetings(days)
        const groups = groupByDay(meetings)
        const conflicts = findConflicts(meetings).map(([a, b]) => [a.subject, b.subject])
        const free = freeSlotsByDay(meetings)
        panel.webview.postMessage({ type: 'agenda', groups, conflicts, free })
      } catch (e: any) {
        panel.webview.postMessage({ type: 'agenda', error: e?.message || String(e) })
      }
      return
    }
    case 'agendaAssist': {
      if (!panel) { return }
      try {
        const meetings = getMeetings(c.get<number>('agenda.days', 7))
        const notes = await vscode.window.withProgress(
          { location: vscode.ProgressLocation.Notification, title: 'Copilot está revisando tu agenda…' },
          (_p, token) => assistAgenda(meetings, token),
        )
        panel.webview.postMessage({ type: 'agendaNotes', notes })
      } catch (e: any) {
        panel.webview.postMessage({ type: 'agendaNotes', error: e?.message || String(e) })
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
  .mailrow.done { opacity: .5; }
  .mailrow.done .subj { text-decoration: line-through; }
  .lab { font-size: 10px; padding: 1px 8px; border-radius: 999px; border: 1px solid currentColor; margin-left: 6px; vertical-align: 1px; }
  .lab-directed { color: #3794ff; }
  .lab-mentioned { color: #d9a017; }
  .lab-informative { color: var(--vscode-descriptionForeground); }
  .controls { display: flex; align-items: center; gap: 14px; margin-top: 12px; flex-wrap: wrap; }
  .seg { display: inline-flex; border: 1px solid var(--vscode-panel-border); border-radius: 6px; overflow: hidden; }
  .seg button { background: transparent; color: var(--vscode-foreground); border: none; border-right: 1px solid var(--vscode-panel-border);
    padding: 5px 12px; font-size: 12px; cursor: pointer; }
  .seg button:last-child { border-right: none; }
  .seg button.active { background: var(--vscode-button-background); color: var(--vscode-button-foreground); }
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
  #reader .composer { margin-top: 14px; border: 1px solid var(--vscode-panel-border); border-radius: 8px; padding: 12px; }
  #reader .crow { display: flex; gap: 10px; margin-top: 10px; }
  .mailbody { white-space: pre-wrap; word-break: break-word; max-height: 340px; overflow: auto; margin-top: 10px;
    padding: 10px; border-top: 1px solid var(--vscode-panel-border); font-size: 12px; line-height: 1.5; }
  .overlay { position: fixed; inset: 0; background: rgba(0,0,0,.55); display: flex; align-items: center; justify-content: center; z-index: 1000; padding: 20px; }
  .modal { background: var(--vscode-editorWidget-background, var(--vscode-editor-background)); border: 1px solid var(--vscode-panel-border);
    border-radius: 10px; width: min(720px, 100%); max-height: 86vh; overflow: auto; padding: 16px 18px; box-shadow: 0 10px 34px rgba(0,0,0,.45); }
  .mhead { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; }
  .mtitle { font-weight: 600; font-size: 15px; }
  .mhead .x { background: none; color: var(--vscode-foreground); font-size: 20px; line-height: 1; padding: 0 4px; }
  .mmeta { margin: 6px 0; font-size: 12px; }
  .mactions { display: flex; gap: 8px; margin-top: 12px; }
  #mcompose textarea { width: 100%; box-sizing: border-box; min-height: 170px; resize: vertical; font-family: var(--vscode-font-family);
    padding: 8px; border-radius: 5px; background: var(--vscode-input-background); color: var(--vscode-input-foreground); border: 1px solid var(--vscode-input-border, transparent); margin-top: 12px; }
  #mcompose .crow { display: flex; align-items: center; gap: 10px; margin-top: 10px; flex-wrap: wrap; }
  .agday { margin-bottom: 10px; }
  .agdate { font-weight: 600; font-size: 12px; color: var(--vscode-descriptionForeground); text-transform: capitalize; margin: 8px 0 2px; }
  #agendaNotes .composer { margin-top: 12px; border: 1px solid var(--vscode-panel-border); border-radius: 8px; padding: 12px; }
  #agendaNotes h4 { margin: 8px 0 3px; font-size: 12px; }
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
    <h2>Correos prioritarios</h2>
    <div class="chips">${senderChips}</div>
    <div class="senderbox">
      <input type="text" id="newSender" placeholder="nombre@empresa.com" onkeydown="if(event.key==='Enter')addSender()">
      <button class="sec" onclick="addSender()">Agregar</button>
      <button class="sec" onclick="loadPriority()">Actualizar</button>
    </div>
    <div class="controls">
      <div class="seg" id="rangeseg">
        <button data-d="1" onclick="setRange(1)">Hoy</button>
        <button data-d="7" onclick="setRange(7)">Semana</button>
        <button data-d="14" class="active" onclick="setRange(14)">2 semanas</button>
      </div>
      <label class="switch"><input type="checkbox" id="actiononly" onchange="renderPriority()"> Solo lo que requiere mi acción</label>
    </div>
    <div id="priority" style="margin-top:12px"><p class="muted">Cargando correos prioritarios…</p></div>
  </div>

  <div class="card">
    <h2>Agenda <span class="muted" style="font-weight:400">· próximos días</span><span style="flex:1"></span><button class="sec" onclick="post('agendaAssist');document.getElementById('agendaNotes').innerHTML='<div class=\\'composer\\'><p class=\\'muted\\'>Copilot está revisando…</p></div>'">Asistente de agenda</button></h2>
    <div id="agenda"><p class="muted">Cargando agenda…</p></div>
    <div id="agendaNotes"></div>
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

  <div id="modalHost"></div>

  <script>
    const vscode = acquireVsCodeApi();
    const store = {};
    function post(type, extra){ vscode.postMessage(Object.assign({type}, extra||{})); }
    function addSender(){ const el=document.getElementById('newSender'); const v=(el.value||'').trim(); if(v){ post('addSender',{email:v}); } }
    let priorityDays = 14;
    function setRange(d){ priorityDays=d; document.querySelectorAll('#rangeseg button').forEach(b=>b.classList.toggle('active', (+b.dataset.d)===d)); loadPriority(); }
    function loadPriority(){ document.getElementById('priority').innerHTML='<p class="muted">Cargando…</p>'; closeModal(); post('loadPriority',{days:priorityDays}); }
    function esc(s){ return String(s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

    let modalId = null;
    function openEmail(id, auto){ modalId=id; document.getElementById('modalHost').innerHTML='<div class="overlay" onclick="if(event.target===this)closeModal()"><div class="modal"><p class="muted">Abriendo correo…</p></div></div>'; post('openEmail',{id, auto:!!auto}); }
    function closeModal(){ modalId=null; document.getElementById('modalHost').innerHTML=''; }
    function renderEmail(m){
      const host=document.getElementById('modalHost');
      if(m.error){ host.innerHTML='<div class="overlay" onclick="if(event.target===this)closeModal()"><div class="modal"><div class="mhead"><span class="mtitle">Correo</span><button class="x" onclick="closeModal()">&times;</button></div><p class="muted">No se pudo abrir: '+esc(m.error)+'</p></div></div>'; return; }
      store[m.id]={sender:m.sender};
      host.innerHTML='<div class="overlay" onclick="if(event.target===this)closeModal()"><div class="modal">'+
        '<div class="mhead"><div class="mtitle">'+esc(m.subject||'(sin asunto)')+'</div><button class="x" onclick="closeModal()" aria-label="Cerrar">&times;</button></div>'+
        '<div class="muted mmeta">De: '+esc(m.sender||'')+' &lt;'+esc(m.senderEmail||'')+'&gt; · '+esc(m.received||'')+'<br>Para: '+esc(m.to||'-')+(m.cc?' · CC: '+esc(m.cc):'')+'</div>'+
        '<div class="mailbody">'+esc(m.body||'')+'</div>'+
        '<div class="mactions" id="mactions">'+
          '<button onclick="draftIA(\\''+esc(m.id)+'\\')">Responder con IA</button>'+
          '<button class="sec" onclick="replyManual(\\''+esc(m.id)+'\\')">Responder manual</button>'+
        '</div><div id="mcompose"></div></div></div>';
      if(m.auto){ draftIA(m.id); }
    }
    function draftIA(id){ const a=document.getElementById('mactions'); if(a)a.style.display='none'; document.getElementById('mcompose').innerHTML='<p class="muted" style="margin-top:12px">Copilot está redactando…</p>'; post('draftReply',{id}); }
    function replyManual(id){ showCompose(id,'',false); }
    function showCompose(id, body, isIA){
      const a=document.getElementById('mactions'); if(a)a.style.display='none';
      document.getElementById('mcompose').innerHTML=
        '<textarea id="rbody">'+esc(body||'')+'</textarea>'+
        '<div class="crow"><label class="switch"><input type="checkbox" id="rall"> Responder a todos</label><span style="flex:1"></span>'+
        (isIA?'<button class="sec" onclick="draftIA(\\''+esc(id)+'\\')">Regenerar</button>':'')+
        '<button class="sec" onclick="closeModal()">Cancelar</button>'+
        '<button onclick="send(\\''+esc(id)+'\\')">Enviar</button></div>';
      const t=document.getElementById('rbody'); if(t){ t.focus(); }
    }
    function send(id){ const body=(document.getElementById('rbody')||{}).value||''; const all=(document.getElementById('rall')||{}).checked||false; post('sendReply',{id,body,replyAll:all,sender:(store[id]||{}).sender||''}); }

    let allEmails = [];
    function statusLabel(s){ return s==='handled'?'Atendido':s==='dismissed'?'No requiere respuesta':''; }
    function renderPriority(){
      const box=document.getElementById('priority');
      const only=(document.getElementById('actiononly')||{}).checked;
      if(!allEmails.length){ box.innerHTML='<p class="muted">Sin correos de esos remitentes en el rango elegido.</p>'; return; }
      let list = only ? allEmails.filter(x=>x.needsAction && !x.status) : allEmails;
      if(!list.length){ box.innerHTML='<p class="muted">Nada pendiente que requiera tu acción.</p>'; return; }
      box.innerHTML=list.map(x=>{
        const lab='<span class="lab lab-'+x.label+'">'+esc(x.labelText)+'</span>';
        let actions='<button class="sec" onclick="openEmail(\\''+esc(x.id)+'\\')">Ver</button> ';
        if(x.status){ actions+='<span class="muted">'+statusLabel(x.status)+'</span> <button class="sec" onclick="markStatus(\\''+esc(x.id)+'\\',null)">Deshacer</button>'; }
        else { actions+='<button class="sec" onclick="markStatus(\\''+esc(x.id)+'\\',\\'handled\\')">Atendido</button>'+
               '<button class="sec" onclick="markStatus(\\''+esc(x.id)+'\\',\\'dismissed\\')">No requiere respuesta</button>'+
               '<button class="sec" onclick="openEmail(\\''+esc(x.id)+'\\',true)">Responder con IA</button>'; }
        return '<div class="mailrow '+(x.unread?'unread ':'')+(x.status?'done':'')+'"><div class="top"><span class="from">'+esc(x.sender||x.senderEmail)+' '+lab+
          '</span><span class="muted">'+esc(x.received)+'</span></div><div class="subj">'+esc(x.subject||'(sin asunto)')+'</div>'+
          '<div class="actions">'+actions+'</div></div>';
      }).join('');
    }
    function markStatus(id,status){ const it=allEmails.find(x=>x.id===id); if(it){ it.status=status; } post('markStatus',{id,status}); renderPriority(); }

    window.addEventListener('message', e => {
      const m = e.data;
      if (m.type === 'priority') {
        if (m.error) { document.getElementById('priority').innerHTML = '<p class="muted">No se pudieron leer los correos: ' + esc(m.error) + '</p>'; return; }
        allEmails = m.emails || [];
        renderPriority();
      } else if (m.type === 'replyDraft') {
        if (m.error) { const c=document.getElementById('mcompose'); if(c){ c.innerHTML='<p class="muted" style="margin-top:12px">No se pudo redactar: '+esc(m.error)+'</p>'; } }
        else if (modalId===m.id) { showCompose(m.id, m.body, true); }
      }
      else if (m.type === 'emailBody') { renderEmail(m); }
      else if (m.type === 'replySent') { closeModal(); }
      else if (m.type === 'agenda') { renderAgenda(m); }
      else if (m.type === 'agendaNotes') {
        const box=document.getElementById('agendaNotes');
        if(m.error){ box.innerHTML='<div class="composer"><p class="muted">No se pudo: '+esc(m.error)+'</p></div>'; }
        else { box.innerHTML='<div class="composer">'+mdlite(m.notes||'')+'</div>'; }
      }
    });
    function mdlite(s){ return esc(s).replace(/\\*\\*(.+?)\\*\\*/g,'<strong>$1</strong>').split('\\n').map(l=>{const li=l.match(/^\\s*[-*]\\s+(.*)$/);const h=l.match(/^\\s*#{1,4}\\s+(.*)$/);if(li)return '<div>&bull; '+li[1]+'</div>';if(h)return '<h4>'+h[1]+'</h4>';return l.trim()?'<div>'+l+'</div>':'';}).join(''); }
    function renderAgenda(m){
      const box=document.getElementById('agenda');
      if(m.error){ box.innerHTML='<p class="muted">No se pudo leer la agenda: '+esc(m.error)+'</p>'; return; }
      if(!m.groups||!m.groups.length){ box.innerHTML='<p class="muted">Sin reuniones en los próximos días.</p>'; return; }
      const freeMap={}; (m.free||[]).forEach(f=>freeMap[f.date]=f.slots);
      let html='';
      if(m.conflicts&&m.conflicts.length){ html+='<div class="warn">Empalmes: '+m.conflicts.map(c=>esc(c[0])+' &harr; '+esc(c[1])).join('; ')+'</div>'; }
      html+=m.groups.map(g=>{
        const slots=(freeMap[g.date]||[]).map(s=>s.start+'–'+s.end).join(', ');
        return '<div class="agday"><div class="agdate">'+esc(g.date)+'</div>'+
          g.items.map(it=>'<div class="row"><span>'+esc((it.start||'').slice(11))+'–'+esc((it.end||'').slice(11))+'  '+esc(it.subject||'(sin título)')+'</span><span class="muted">'+(it.attendees?esc(it.attendees)+' inv.':'')+'</span></div>').join('')+
          (slots?'<div class="muted" style="padding:4px 0 2px">Libre: '+esc(slots)+'</div>':'')+'</div>';
      }).join('');
      box.innerHTML=html;
    }
    window.addEventListener('keydown', e => { if (e.key === 'Escape' && modalId) { closeModal(); } });
    post('loadPriority');
    post('loadAgenda');
  </script>
</body></html>`
}
