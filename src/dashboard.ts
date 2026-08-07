import * as vscode from 'vscode'
import { DayEntry, toMarkdown } from './bitacoraCore'
import { loadEntries, runDailyReview } from './dailyReview'
import { getPriorityEmails, getInboxClassified, isWindows } from './priorityInbox'
import { getRequests } from './dataRequests'
import { loadAssignments, assign as assignReq, setStatus as setReqStatus, unassign as unassignReq } from './dataAssign'
import { classifyPriority, labelText } from './priorityClassify'
import { loadStatus, setStatus, clearStatus } from './priorityState'
import { readEmailBody, sendReply, sendMail, getMe } from './outlookActions'
import { draftReply, assistAgenda, summarizePriority } from './copilot'
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

interface KindDef { keyword: string; file: string; teamCfg: string; reqMsg: string; teamMsg: string; label: string }
const KIND: Record<string, KindDef> = {
  data: { keyword: 'datos sint', file: 'dataAssignments.json', teamCfg: 'dataTeam', reqMsg: 'requests', teamMsg: 'team', label: 'Datos Sintéticos' },
  perf: { keyword: 'solicitud de performance', file: 'perfAssignments.json', teamCfg: 'perfTeam', reqMsg: 'perfRequests', teamMsg: 'perfTeam', label: 'Performance' },
}

async function postRequests(context: vscode.ExtensionContext, k: string): Promise<void> {
  if (!panel) { return }
  const K = KIND[k]
  const c = vscode.workspace.getConfiguration('m365')
  try {
    const reqs = getRequests(K.keyword, 45, 150)
    const map = await loadAssignments(context, K.file)
    panel.webview.postMessage({ type: K.reqMsg, items: reqs.map(r => ({ ...r, assignment: map[r.id] || null })), team: c.get<string[]>(K.teamCfg, []) })
  } catch (e: any) {
    panel.webview.postMessage({ type: K.reqMsg, error: e?.message || String(e) })
  }
}

async function doAssign(context: vscode.ExtensionContext, k: string, m: any): Promise<void> {
  const K = KIND[k]
  const member = String(m.member)
  await assignReq(context, K.file, String(m.id), member, Date.now())
  const pick = await vscode.window.showInformationMessage(
    `Solicitud asignada a ${member}. ¿Enviarle un correo de aviso?`, 'Enviar aviso',
  )
  if (pick === 'Enviar aviso') {
    const body =
      `Hola,\n\nSe te asignó una solicitud de ${K.label}.\n\n` +
      `Solicitante: ${m.solicitante || '-'}\n` +
      `Equipo: ${m.equipo || '-'}\n` +
      `Proyecto: ${m.proyecto || '-'}\n` +
      `Fecha solicitada: ${m.fecha || '-'}\n\n` +
      `Por favor dale seguimiento. Gracias.`
    try {
      sendMail(member, `Solicitud de ${K.label} asignada${m.proyecto ? ' — ' + m.proyecto : ''}`, body)
      vscode.window.showInformationMessage(`Aviso enviado a ${member}.`)
    } catch (e: any) {
      vscode.window.showErrorMessage(`No se pudo enviar el aviso: ${e?.message || e}`)
    }
  }
}

async function addTeamMember(context: vscode.ExtensionContext, k: string, email: string): Promise<void> {
  const K = KIND[k]
  const c = vscode.workspace.getConfiguration('m365')
  const clean = String(email || '').trim()
  if (clean) {
    const cur = c.get<string[]>(K.teamCfg, [])
    if (!cur.includes(clean)) { await c.update(K.teamCfg, [...cur, clean], vscode.ConfigurationTarget.Global) }
  }
  if (panel) { panel.webview.postMessage({ type: K.teamMsg, team: c.get<string[]>(K.teamCfg, []) }) }
}

async function removeTeamMember(context: vscode.ExtensionContext, k: string, email: string): Promise<void> {
  const K = KIND[k]
  const c = vscode.workspace.getConfiguration('m365')
  await c.update(K.teamCfg, c.get<string[]>(K.teamCfg, []).filter(x => x !== email), vscode.ConfigurationTarget.Global)
  if (panel) { panel.webview.postMessage({ type: K.teamMsg, team: c.get<string[]>(K.teamCfg, []) }) }
}

/** Convierte el rango elegido (today/week/2weeks, por calendario) a dias hacia atras. */
function daysFromRange(range: string): number {
  const DAY = 86400000
  const now = Date.now()
  if (range === 'week' || range === '2weeks') {
    const x = new Date()
    const diff = (x.getDay() + 6) % 7 // lunes = inicio de semana
    x.setDate(x.getDate() - diff)
    x.setHours(0, 0, 0, 0)
    if (range === '2weeks') { x.setDate(x.getDate() - 7) }
    return Math.max((now - x.getTime()) / DAY, 0.01)
  }
  const m = new Date()
  m.setHours(0, 0, 0, 0)
  return Math.max((now - m.getTime()) / DAY, 0.01)
}

// ── Estado ────────────────────────────────────────────────────────
interface State {
  entries: DayEntry[]
  enabled: boolean
  days: number[]
  startHour: number
  endHour: number
  maxEmails: number
  pollMinutes: number
  senders: string[]
  team: string[]
  perfTeam: string[]
  recipes: string[]
  windows: boolean
}

async function gather(context: vscode.ExtensionContext): Promise<State> {
  const c = vscode.workspace.getConfiguration('m365')
  return {
    entries: await loadEntries(context),
    enabled: c.get('dailyReview.enabled', true),
    days: c.get<number[]>('dailyReview.days', [1, 2, 3, 4, 5]),
    startHour: c.get('dailyReview.startHour', 7),
    endHour: c.get('dailyReview.endHour', 3),
    maxEmails: c.get('dailyReview.maxEmails', 40),
    pollMinutes: c.get('dailyReview.pollMinutes', 30),
    senders: c.get<string[]>('prioritySenders', []),
    team: c.get<string[]>('dataTeam', []),
    perfTeam: c.get<string[]>('perfTeam', []),
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
    case 'loadInbox': {
      if (!panel) { return }
      try {
        const days = daysFromRange(String(m.range || 'today'))
        const senders = c.get<string[]>('prioritySenders', [])
        const priority = getPriorityEmails(senders, days, 40)
        const other = getInboxClassified(senders, days, 150).filter(x => !x.isPriority)
        panel.webview.postMessage({ type: 'inbox', priority, other })
        // Resumen de Copilot de los prioritarios (una sola llamada), en segundo plano.
        try {
          const summaries = await summarizePriority(priority, new vscode.CancellationTokenSource().token)
          if (panel) { panel.webview.postMessage({ type: 'inboxSummaries', summaries }) }
        } catch {
          if (panel) { panel.webview.postMessage({ type: 'inboxSummaries', summaries: [] }) }
        }
      } catch (e: any) {
        panel.webview.postMessage({ type: 'inbox', error: e?.message || String(e) })
      }
      return
    }
    case 'loadRequests': { await postRequests(context, 'data'); return }
    case 'loadPerfRequests': { await postRequests(context, 'perf'); return }
    case 'assignRequest': { await doAssign(context, 'data', m); return }
    case 'assignPerf': { await doAssign(context, 'perf', m); return }
    case 'setRequestStatus': { await setReqStatus(context, KIND.data.file, String(m.id), m.status === 'finalizada' ? 'finalizada' : 'seguimiento'); return }
    case 'setPerfStatus': { await setReqStatus(context, KIND.perf.file, String(m.id), m.status === 'finalizada' ? 'finalizada' : 'seguimiento'); return }
    case 'unassignRequest': { await unassignReq(context, KIND.data.file, String(m.id)); return }
    case 'unassignPerf': { await unassignReq(context, KIND.perf.file, String(m.id)); return }
    case 'addMember': { await addTeamMember(context, 'data', m.email); return }
    case 'addPerfMember': { await addTeamMember(context, 'perf', m.email); return }
    case 'removeMember': { await removeTeamMember(context, 'data', m.email); return }
    case 'removePerfMember': { await removeTeamMember(context, 'perf', m.email); return }
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
          const replied = !!(e.repliedAt && e.received && e.repliedAt >= e.received)
          return {
            id: e.id, sender: e.sender, senderEmail: e.senderEmail, subject: e.subject,
            received: e.received, unread: e.unread,
            label: cl.label, labelText: labelText(cl.label),
            needsAction: cl.needsAction && !replied,
            replied, repliedAt: e.repliedAt || '', replyBody: e.replyBody || '',
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

  const hourOpts = Array.from({ length: 24 }, (_, h) => `<option value="${h}">${String(h).padStart(2, '0')}:00</option>`).join('')
  const filterBar = (p: string) =>
    `<div class="filters">
      <input type="text" placeholder="Buscar remitente o asunto" oninput="setFilter('${p}','search',this.value)">
      <input type="date" onchange="setFilter('${p}','date',this.value)" title="Filtrar por fecha">
      <select onchange="setFilter('${p}','hour',this.value)"><option value="">Hora</option>${hourOpts}</select>
      <select onchange="setFilter('${p}','status',this.value)"><option value="">Estatus</option><option value="unread">No leídos</option><option value="replied">Respondidos</option><option value="unreplied">Sin responder</option></select>
    </div>`

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
  input[type=number], input[type=text], select { width: 100%; box-sizing: border-box; padding: 6px 8px; border-radius: 5px;
    background: var(--vscode-input-background); color: var(--vscode-input-foreground); border: 1px solid var(--vscode-input-border, transparent); }
  select { background: var(--vscode-dropdown-background, var(--vscode-input-background)); color: var(--vscode-dropdown-foreground, var(--vscode-input-foreground)); }
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
  .lab-replied { color: #3fb950; }
  .replied-callout { border-left: 3px solid #3fb950; background: rgba(63,185,80,.08); border-radius: 4px; padding: 8px 10px; margin: 10px 0; }
  .controls { display: flex; align-items: center; gap: 14px; margin-top: 12px; flex-wrap: wrap; }
  .seg { display: inline-flex; border: 1px solid var(--vscode-panel-border); border-radius: 6px; overflow: hidden; }
  .seg button { background: transparent; color: var(--vscode-foreground); border: none; border-right: 1px solid var(--vscode-panel-border);
    padding: 5px 12px; font-size: 12px; cursor: pointer; }
  .seg button:last-child { border-right: none; }
  .seg button.active { background: var(--vscode-button-background); color: var(--vscode-button-foreground); }
  .filters { display: flex; gap: 6px; flex-wrap: wrap; margin: 8px 0 10px; }
  .filters input, .filters select { flex: 1; min-width: 92px; padding: 4px 6px; font-size: 12px; }
  .filters input[type=text] { flex: 2; min-width: 130px; }
  .pager { display: flex; align-items: center; justify-content: center; gap: 10px; margin-top: 10px; font-size: 12px; }
  .pager button[disabled] { opacity: .4; cursor: default; }
  .badge-meet { display: inline-block; margin-top: 6px; font-size: 11px; padding: 2px 8px; border-radius: 4px; background: rgba(215,160,23,.15); color: #d9a017; border: 1px solid #d9a017; }
  .mailrow.ismeet { border-left: 3px solid #d9a017; padding-left: 8px; }
  .summ { margin: 6px 0; font-size: 12px; }
  .rec { font-size: 11px; margin-top: 2px; }
  .rep { font-size: 11px; color: #3fb950; margin-top: 3px; }
  .pend { font-size: 11px; color: var(--vscode-descriptionForeground); margin-top: 3px; }
  .wrow { padding: 8px 0; border-bottom: 1px solid var(--vscode-panel-border); }
  .wrow:last-child { border-bottom: none; }
  .loadbar { height: 8px; background: var(--vscode-panel-border); border-radius: 4px; overflow: hidden; margin: 5px 0; }
  .loadbar span { display: block; height: 100%; background: var(--vscode-button-background); }
  .wreq { font-size: 11px; padding: 1px 0 1px 6px; }
  .actions select { max-width: 200px; }
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
    <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap">
      <h2>Revisión de correo</h2>
      <label class="switch">Activada<input type="checkbox" ${s.enabled ? 'checked' : ''} onchange="post('setSetting',{key:'dailyReview.enabled',value:this.checked})"></label>
    </div>
    <div style="margin-top:12px">
      <label>Días</label>
      <div class="seg" id="daysseg">
        ${[{ n: 1, l: 'L' }, { n: 2, l: 'M' }, { n: 3, l: 'M' }, { n: 4, l: 'J' }, { n: 5, l: 'V' }, { n: 6, l: 'S' }, { n: 0, l: 'D' }]
          .map(d => `<button data-n="${d.n}" class="${s.days.includes(d.n) ? 'active' : ''}" onclick="toggleDay(${d.n})">${d.l}</button>`).join('')}
      </div>
    </div>
    <div class="formgrid" style="margin-top:12px">
      <div><label>Desde (hora)</label><input type="number" min="0" max="23" value="${s.startHour}" onchange="post('setSetting',{key:'dailyReview.startHour',value:+this.value})"></div>
      <div><label>Hasta (hora)</label><input type="number" min="0" max="23" value="${s.endHour}" onchange="post('setSetting',{key:'dailyReview.endHour',value:+this.value})"></div>
      <div><label>Sondeo cada</label>
        <select onchange="post('setSetting',{key:'dailyReview.pollMinutes',value:+this.value})">
          <option value="10" ${s.pollMinutes === 10 ? 'selected' : ''}>10 min</option>
          <option value="30" ${s.pollMinutes === 30 ? 'selected' : ''}>30 min</option>
          <option value="60" ${s.pollMinutes === 60 ? 'selected' : ''}>60 min</option>
        </select>
      </div>
      <div><label>Correos por carpeta</label><input type="number" min="1" max="100" value="${s.maxEmails}" onchange="post('setSetting',{key:'dailyReview.maxEmails',value:+this.value})"></div>
      <button onclick="post('runReview')">Correr ahora</button>
    </div>
    <p class="sub" style="margin-top:10px">El agente sondea tus correos cada ${s.pollMinutes} min, en los días y la ventana horaria elegidos (mientras VS Code esté abierto). En cada sondeo toma ${s.maxEmails} de bandeja de entrada y ${s.maxEmails} de enviados para detectar correos nuevos y los que ya respondiste, y tu Copilot los resume en la Bitácora.</p>
  </div>

  <div class="card">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <h2>Correos</h2>
      <div class="seg" id="inboxseg">
        <button data-r="today" class="active" onclick="setInboxRange('today')">Hoy</button>
        <button data-r="week" onclick="setInboxRange('week')">Esta semana</button>
        <button data-r="2weeks" onclick="setInboxRange('2weeks')">2 semanas</button>
      </div>
      <span style="flex:1"></span>
      <button class="sec" onclick="loadInbox()">Actualizar</button>
    </div>
    <div class="senderbox" style="margin-top:10px">
      <input type="text" id="newSender" placeholder="correo prioritario: nombre@empresa.com" onkeydown="if(event.key==='Enter')addSender()">
      <button class="sec" onclick="addSender()">Agregar prioritario</button>
    </div>
    <div class="chips" style="margin-top:8px">${senderChips}</div>
  </div>

  <div class="grid2">
    <div class="card">
      <h2>Prioritarios <span class="muted" style="font-weight:400">· sí me importan</span></h2>
      ${filterBar('prio')}
      <div id="prio"><p class="muted">Cargando…</p></div>
      <div id="prio-pager" class="pager"></div>
    </div>
    <div class="card">
      <h2>No prioritarios <span class="muted" style="font-weight:400">· pueden esperar</span></h2>
      ${filterBar('other')}
      <div id="other"><p class="muted">Cargando…</p></div>
      <div id="other-pager" class="pager"></div>
    </div>
  </div>

  <div class="card">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <h2>Solicitudes de Datos Sintéticos</h2>
      <span style="flex:1"></span>
      <button class="sec" onclick="loadReqsK('data')">Actualizar</button>
    </div>
    <div class="senderbox" style="margin-top:10px">
      <input type="text" id="newMember" placeholder="correo de un miembro del equipo" onkeydown="if(event.key==='Enter')addMemberK('data')">
      <button class="sec" onclick="addMemberK('data')">Agregar miembro</button>
    </div>
    <div class="chips" id="teamchips" style="margin-top:8px"></div>
  </div>

  <div class="grid2">
    <div class="card">
      <h2>Bandeja de solicitudes <span class="muted" id="reqcount"></span></h2>
      <div id="requests"><p class="muted">Cargando…</p></div>
    </div>
    <div class="card">
      <h2>Equipo operador <span class="muted" style="font-weight:400">· carga por persona</span></h2>
      <div id="workload"><p class="muted">—</p></div>
    </div>
  </div>

  <div class="card">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <h2>Solicitudes de Performance</h2>
      <span style="flex:1"></span>
      <button class="sec" onclick="loadReqsK('perf')">Actualizar</button>
    </div>
    <div class="senderbox" style="margin-top:10px">
      <input type="text" id="newPerfMember" placeholder="correo de un miembro del equipo" onkeydown="if(event.key==='Enter')addMemberK('perf')">
      <button class="sec" onclick="addMemberK('perf')">Agregar miembro</button>
    </div>
    <div class="chips" id="perfTeamchips" style="margin-top:8px"></div>
  </div>

  <div class="grid2">
    <div class="card">
      <h2>Bandeja de solicitudes <span class="muted" id="perfcount"></span></h2>
      <div id="perfRequests"><p class="muted">Cargando…</p></div>
    </div>
    <div class="card">
      <h2>Equipo operador <span class="muted" style="font-weight:400">· carga por persona</span></h2>
      <div id="perfWorkload"><p class="muted">—</p></div>
    </div>
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

    let reviewDays = ${JSON.stringify(s.days)};
    function toggleDay(n){ const i=reviewDays.indexOf(n); if(i>=0){reviewDays.splice(i,1);}else{reviewDays.push(n);} document.querySelectorAll('#daysseg button').forEach(b=>b.classList.toggle('active', reviewDays.includes(+b.dataset.n))); post('setSetting',{key:'dailyReview.days',value:reviewDays}); }

    let modalId = null;
    function openEmail(id, auto){ modalId=id; document.getElementById('modalHost').innerHTML='<div class="overlay" onclick="if(event.target===this)closeModal()"><div class="modal"><p class="muted">Abriendo correo…</p></div></div>'; post('openEmail',{id, auto:!!auto}); }
    function closeModal(){ modalId=null; document.getElementById('modalHost').innerHTML=''; }
    function renderEmail(m){
      const host=document.getElementById('modalHost');
      if(m.error){ host.innerHTML='<div class="overlay" onclick="if(event.target===this)closeModal()"><div class="modal"><div class="mhead"><span class="mtitle">Correo</span><button class="x" onclick="closeModal()">&times;</button></div><p class="muted">No se pudo abrir: '+esc(m.error)+'</p></div></div>'; return; }
      store[m.id]={sender:m.sender};
      const src=allEmails.find(x=>x.id===m.id)||{};
      const repliedBlock=src.replied?('<div class="replied-callout"><div style="font-weight:600;color:#3fb950">Ya respondiste este correo · '+esc(src.repliedAt)+'</div><div class="mailbody" style="max-height:200px;border:none;padding:6px 0 0">'+esc(src.replyBody||'(sin texto)')+'</div></div>'):'';
      host.innerHTML='<div class="overlay" onclick="if(event.target===this)closeModal()"><div class="modal">'+
        '<div class="mhead"><div class="mtitle">'+esc(m.subject||'(sin asunto)')+'</div><button class="x" onclick="closeModal()" aria-label="Cerrar">&times;</button></div>'+
        '<div class="muted mmeta">De: '+esc(m.sender||'')+' &lt;'+esc(m.senderEmail||'')+'&gt; · '+esc(m.received||'')+'<br>Para: '+esc(m.to||'-')+(m.cc?' · CC: '+esc(m.cc):'')+'</div>'+
        repliedBlock+
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
        const lab='<span class="lab lab-'+x.label+'">'+esc(x.labelText)+'</span>'+(x.replied?'<span class="lab lab-replied">Respondido</span>':'');
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

    let inboxRange = 'today';
    const lists = { prio:{data:[],page:0,search:'',date:'',hour:'',status:''}, other:{data:[],page:0,search:'',date:'',hour:'',status:''} };
    const PAGE = 6;
    function setInboxRange(r){ inboxRange=r; document.querySelectorAll('#inboxseg button').forEach(b=>b.classList.toggle('active',b.dataset.r===r)); loadInbox(); }
    function loadInbox(){ ['prio','other'].forEach(k=>{document.getElementById(k).innerHTML='<p class="muted">Cargando…</p>';document.getElementById(k+'-pager').innerHTML='';}); post('loadInbox',{range:inboxRange}); }
    function setFilter(key,f,v){ lists[key][f]=v; lists[key].page=0; renderList(key); }
    function pageNav(key,d){ lists[key].page+=d; renderList(key); }
    function daysAgo(recv){ const d=Date.parse(String(recv||'').replace(' ','T')); if(!d)return ''; const n=Math.floor((Date.now()-d)/86400000); return n<=0?'hoy':(n===1?'1 día en bandeja':(n+' días en bandeja')); }
    function filtered(key){ const s=lists[key]; const q=(s.search||'').toLowerCase();
      return s.data.filter(x=>{
        if(q && !(((x.sender||'')+' '+(x.senderEmail||'')+' '+(x.subject||'')).toLowerCase().includes(q))) return false;
        if(s.date && !String(x.received||'').startsWith(s.date)) return false;
        if(s.hour!=='' && String(x.received||'').slice(11,13)!==String(s.hour).padStart(2,'0')) return false;
        if(s.status==='unread' && !x.unread) return false;
        if(s.status==='replied' && !x.repliedAt) return false;
        if(s.status==='unreplied' && x.repliedAt) return false;
        return true;
      });
    }
    function renderList(key){
      const s=lists[key]; const arr=filtered(key);
      const pages=Math.max(1,Math.ceil(arr.length/PAGE));
      if(s.page>=pages)s.page=pages-1; if(s.page<0)s.page=0;
      const slice=arr.slice(s.page*PAGE,(s.page+1)*PAGE);
      document.getElementById(key).innerHTML = slice.length? slice.map(x=> key==='prio'?prioItem(x):otherItem(x)).join('') : '<p class="muted">Sin correos con esos filtros.</p>';
      document.getElementById(key+'-pager').innerHTML = pages>1
        ? ('<button class="sec" onclick="pageNav(\\''+key+'\\',-1)"'+(s.page===0?' disabled':'')+'>‹ Anterior</button><span class="muted">'+(s.page+1)+' / '+pages+' · '+arr.length+'</span><button class="sec" onclick="pageNav(\\''+key+'\\',1)"'+(s.page>=pages-1?' disabled':'')+'>Siguiente ›</button>')
        : (arr.length? '<span class="muted">'+arr.length+' correos</span>':'');
    }
    function otherItem(x){ return '<div class="mailrow '+(x.unread?'unread ':'')+'"><div class="top"><span class="from">'+esc(x.sender||x.senderEmail)+'</span><span class="muted">'+esc(x.received)+' · '+daysAgo(x.received)+'</span></div><div class="subj">'+esc(x.subject||'(sin asunto)')+'</div><div class="actions"><button class="sec" onclick="openEmail(\\''+esc(x.id)+'\\')">Ver</button></div></div>'; }
    function prioItem(x){
      const rec=(x.to||'')+(x.cc?('; '+x.cc):'');
      const mt=x.meeting&&x.meeting.es;
      const meetBadge= mt? '<div class="badge-meet">Junta'+([x.meeting.modo,x.meeting.cuando,x.meeting.donde].filter(Boolean).length?' · '+esc([x.meeting.modo,x.meeting.cuando,x.meeting.donde].filter(Boolean).join(' · ')):'')+'</div>' : '';
      const summ = (x.resumen!==undefined)? '<div class="summ">'+esc(x.resumen||'—')+'</div>' : '<div class="summ muted">Resumiendo…</div>';
      const replied = x.repliedAt? '<div class="rep">Respondido · '+esc(x.repliedAt)+(x.respuesta?(' — '+esc(x.respuesta)):'')+'</div>' : '<div class="pend">Sin responder</div>';
      return '<div class="mailrow '+(x.unread?'unread ':'')+(mt?'ismeet ':'')+'">'+
        '<div class="top"><span class="from">'+esc(x.sender||x.senderEmail)+'</span><span class="muted">'+esc(x.received)+' · '+daysAgo(x.received)+'</span></div>'+
        '<div class="subj">'+esc(x.subject||'(sin asunto)')+'</div>'+
        meetBadge+summ+
        '<div class="muted rec">Con: '+esc(rec||'—')+'</div>'+
        replied+
        '<div class="actions"><button class="sec" onclick="openEmail(\\''+esc(x.id)+'\\')">Ver</button></div></div>';
    }
    function renderInbox(m){
      if(m.error){ ['prio','other'].forEach(k=>document.getElementById(k).innerHTML='<p class="muted">No se pudo leer la bandeja: '+esc(m.error)+'</p>'); return; }
      lists.prio.data=m.priority||[]; lists.prio.page=0;
      lists.other.data=m.other||[]; lists.other.page=0;
      renderList('prio'); renderList('other');
    }
    function mergeSummaries(arr){ const by={}; (arr||[]).forEach(s=>{by[s.id]=s;}); lists.prio.data.forEach(x=>{ const s=by[x.id]; if(s){ x.resumen=s.resumen||''; x.meeting=s.reunion||{es:false}; x.respuesta=s.respuesta||''; } else { x.resumen=x.resumen||''; } }); renderList('prio'); }

    const BRD = {
      data: { team: ${JSON.stringify(s.team)}, reqs: [], ids:{req:'requests',wl:'workload',chips:'teamchips',cnt:'reqcount',input:'newMember'}, msg:{load:'loadRequests',assign:'assignRequest',status:'setRequestStatus',unassign:'unassignRequest',add:'addMember',remove:'removeMember'} },
      perf: { team: ${JSON.stringify(s.perfTeam)}, reqs: [], ids:{req:'perfRequests',wl:'perfWorkload',chips:'perfTeamchips',cnt:'perfcount',input:'newPerfMember'}, msg:{load:'loadPerfRequests',assign:'assignPerf',status:'setPerfStatus',unassign:'unassignPerf',add:'addPerfMember',remove:'removePerfMember'} },
    };
    function memberName(m){ return String(m||'').split('@')[0]; }
    function daysSince(ts){ if(!ts)return ''; const n=Math.floor((Date.now()-ts)/86400000); return n<=0?'hoy':(n===1?'1 día':(n+' días')); }
    function chipsK(k){ const B=BRD[k]; document.getElementById(B.ids.chips).innerHTML = B.team.length? B.team.map(m=>'<span class="chip">'+esc(m)+'<button class="x" onclick="rmMemberK(\\''+k+'\\',\\''+esc(m)+'\\')">&times;</button></span>').join('') : '<span class="muted">Sin miembros. Agrega correos arriba.</span>'; }
    function addMemberK(k){ const B=BRD[k]; const el=document.getElementById(B.ids.input); const v=(el.value||'').trim(); if(v){ post(B.msg.add,{email:v}); el.value=''; } }
    function rmMemberK(k,m){ post(BRD[k].msg.remove,{email:m}); }
    function loadReqsK(k){ document.getElementById(BRD[k].ids.req).innerHTML='<p class="muted">Cargando…</p>'; post(BRD[k].msg.load); }
    function reqSelectK(k,r){ const B=BRD[k]; const a=r.assignment; return '<select onchange="assignK(\\''+k+'\\',\\''+esc(r.id)+'\\',this.value)"><option value="">Asignar a…</option>'+B.team.map(m=>'<option value="'+esc(m)+'"'+(a&&a.member===m?' selected':'')+'>'+esc(memberName(m))+'</option>').join('')+'</select>'; }
    function renderReqsK(k){
      const B=BRD[k]; const box=document.getElementById(B.ids.req);
      const rc=document.getElementById(B.ids.cnt); if(rc)rc.textContent=B.reqs.length?('· '+B.reqs.length):'';
      if(!B.reqs.length){ box.innerHTML='<p class="muted">Sin solicitudes en el rango.</p>'; renderWLK(k); return; }
      box.innerHTML=B.reqs.map(r=>{
        const a=r.assignment; let row;
        if(a){ row='<div class="rec">Asignada a <strong>'+esc(memberName(a.member))+'</strong> · hace '+daysSince(a.assignedAt)+' · '+(a.status==='finalizada'?'<span class="rep">Finalizada</span>':'<span class="pend">En seguimiento</span>')+'</div>'+
          '<div class="actions">'+reqSelectK(k,r)+
          (a.status==='finalizada'?'<button class="sec" onclick="statusK(\\''+k+'\\',\\''+esc(r.id)+'\\',\\'seguimiento\\')">Reabrir</button>':'<button class="sec" onclick="statusK(\\''+k+'\\',\\''+esc(r.id)+'\\',\\'finalizada\\')">Marcar finalizada</button>')+
          '<button class="sec" onclick="unassignK(\\''+k+'\\',\\''+esc(r.id)+'\\')">Quitar</button>'+
          '<button class="sec" onclick="openEmail(\\''+esc(r.id)+'\\')">Ver</button></div>';
        } else { row='<div class="actions">'+reqSelectK(k,r)+'<button class="sec" onclick="openEmail(\\''+esc(r.id)+'\\')">Ver</button></div>'; }
        return '<div class="mailrow"><div class="top"><span class="from">'+esc(r.solicitante||'(sin solicitante)')+'</span><span class="muted">'+esc(r.received)+'</span></div>'+
          '<div class="subj">'+esc(r.proyecto||r.subject||'')+'</div>'+
          '<div class="muted rec">Equipo: '+esc(r.equipo||'—')+' · Fecha solicitada: '+esc(r.fecha||'—')+'</div>'+
          row+'</div>';
      }).join('');
      renderWLK(k);
    }
    function renderWLK(k){
      const B=BRD[k]; const box=document.getElementById(B.ids.wl);
      if(!B.team.length){ box.innerHTML='<p class="muted">Agrega miembros del equipo para ver la carga.</p>'; return; }
      const c={}; B.team.forEach(m=>c[m]={t:0,f:0,s:0,reqs:[]});
      B.reqs.forEach(r=>{ const a=r.assignment; if(a&&c[a.member]){ c[a.member].t++; if(a.status==='finalizada')c[a.member].f++; else c[a.member].s++; c[a.member].reqs.push(r);} });
      const max=Math.max(1, ...B.team.map(m=>c[m].t));
      box.innerHTML=B.team.map(m=>{ const x=c[m];
        const reqs=x.reqs.map(r=>'<div class="wreq muted">• '+esc(r.proyecto||r.solicitante||'solicitud')+' · hace '+daysSince(r.assignment.assignedAt)+' · '+(r.assignment.status==='finalizada'?'finalizada':'seguimiento')+'</div>').join('');
        return '<div class="wrow"><div class="top"><span class="from">'+esc(memberName(m))+'</span><span class="muted">'+x.t+' ('+x.s+' seg · '+x.f+' fin)</span></div>'+
          '<div class="loadbar"><span style="width:'+Math.round(x.t/max*100)+'%"></span></div>'+reqs+'</div>';
      }).join('');
    }
    function assignK(k,id,member){ if(!member)return; const B=BRD[k]; const r=B.reqs.find(x=>x.id===id); if(r){ r.assignment={member, assignedAt:(r.assignment&&r.assignment.member===member)?r.assignment.assignedAt:Date.now(), status:r.assignment?r.assignment.status:'seguimiento'}; } post(B.msg.assign,{id,member,solicitante:(r&&r.solicitante)||'',equipo:(r&&r.equipo)||'',proyecto:(r&&r.proyecto)||'',fecha:(r&&r.fecha)||''}); renderReqsK(k); }
    function statusK(k,id,status){ const B=BRD[k]; const r=B.reqs.find(x=>x.id===id); if(r&&r.assignment){ r.assignment.status=status; } post(B.msg.status,{id,status}); renderReqsK(k); }
    function unassignK(k,id){ const B=BRD[k]; const r=B.reqs.find(x=>x.id===id); if(r){ r.assignment=null; } post(B.msg.unassign,{id}); renderReqsK(k); }

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
      else if (m.type === 'inbox') { renderInbox(m); }
      else if (m.type === 'inboxSummaries') { mergeSummaries(m.summaries); }
      else if (m.type === 'requests') { if(m.error){ document.getElementById('requests').innerHTML='<p class="muted">No se pudieron leer las solicitudes: '+esc(m.error)+'</p>'; } else { BRD.data.reqs=m.items||[]; if(m.team)BRD.data.team=m.team; chipsK('data'); renderReqsK('data'); } }
      else if (m.type === 'team') { BRD.data.team=m.team||[]; chipsK('data'); renderReqsK('data'); }
      else if (m.type === 'perfRequests') { if(m.error){ document.getElementById('perfRequests').innerHTML='<p class="muted">No se pudieron leer las solicitudes: '+esc(m.error)+'</p>'; } else { BRD.perf.reqs=m.items||[]; if(m.team)BRD.perf.team=m.team; chipsK('perf'); renderReqsK('perf'); } }
      else if (m.type === 'perfTeam') { BRD.perf.team=m.team||[]; chipsK('perf'); renderReqsK('perf'); }
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
    chipsK('data'); chipsK('perf');
    loadInbox();
    loadReqsK('data'); loadReqsK('perf');
  </script>
</body></html>`
}
