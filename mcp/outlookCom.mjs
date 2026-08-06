/**
 * Herramientas de Outlook de escritorio (COM via PowerShell).
 * Correo: outlook_list_emails, outlook_read_email, outlook_send_email (confirm),
 *         outlook_reply_email (confirm)
 * Agenda: outlook_list_calendar, outlook_shared_calendar, outlook_create_event,
 *         outlook_find_free
 * Skills: daily_briefing, inbox_triage
 * Diagnostico: outlook_status
 */
import { z } from 'zod'
import { ps, isWindows, getMeCom } from './comShared.mjs'
import { buildPreview, isConfirmed } from './lib/preview.mjs'
import { classifyEmail, sortByPriority } from './lib/triage.mjs'
import { freeSlots } from './lib/freebusy.mjs'

const ok = (text) => ({ content: [{ type: 'text', text }] })
const bad = (text) => ({ content: [{ type: 'text', text }], isError: true })
const json = (obj) => ok(JSON.stringify(obj, null, 2))

function parseJson(out) {
  if (!out || !out.trim()) { return [] }
  const j = JSON.parse(out)
  return Array.isArray(j) ? j : [j]
}

export function registerOutlookComTools(server) {
  // ── Diagnostico ─────────────────────────────────────────────────
  server.tool(
    'outlook_status',
    'Verifica que Outlook de escritorio este accesible por COM y devuelve tu identidad (nombre/correo). Util como health-check.',
    {},
    async () => {
      if (!isWindows) { return bad('Este plugin usa Outlook de escritorio (COM) y requiere Windows con Outlook instalado y abierto.') }
      try {
        const me = getMeCom()
        return json({ connected: !!me.email || !!me.name, you: me, engine: 'Outlook desktop (COM)' })
      } catch (e) { return bad(`No se pudo hablar con Outlook: ${e.message}`) }
    },
  )

  // ── Listar correos ──────────────────────────────────────────────
  server.tool(
    'outlook_list_emails',
    'Lista los correos recientes de una carpeta de Outlook (Inbox, SentItems, Drafts).',
    {
      count: z.number().int().min(1).max(50).optional().describe('Cuantos correos (default 20, max 50)'),
      folder: z.enum(['Inbox', 'SentItems', 'Drafts']).optional().describe('Carpeta (default Inbox)'),
    },
    async (args) => {
      const count = Math.min(args?.count ?? 20, 50)
      const folderMap = { Inbox: 6, SentItems: 5, Drafts: 16 }
      const folderNum = folderMap[args?.folder ?? 'Inbox'] ?? 6
      const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$folder = $ns.GetDefaultFolder(${folderNum})
$items = $folder.Items
$items.Sort("[ReceivedTime]", $true)
$total = [Math]::Min($items.Count, ${count})
$result = @()
for ($i = 1; $i -le $total; $i++) {
  try {
    $item = $items.Item($i)
    $result += [PSCustomObject]@{
      id = $item.EntryID; subject = $item.Subject; sender = $item.SenderName
      senderEmail = $item.SenderEmailAddress
      received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm"); unread = $item.UnRead
      to = $item.To; cc = $item.CC
      preview = $item.Body.Substring(0, [Math]::Min($item.Body.Length, 200))
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
      try { const m = parseJson(ps(script)); return json({ count: m.length, messages: m }) }
      catch (e) { return bad(e.message) }
    },
  )

  // ── Leer correo ─────────────────────────────────────────────────
  server.tool(
    'outlook_read_email',
    'Lee el contenido completo de un correo por su EntryID (de outlook_list_emails).',
    { entry_id: z.string().describe('EntryID del correo') },
    async ({ entry_id }) => {
      const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$item = $ns.GetItemFromID($env:ENTRY_ID)
[PSCustomObject]@{
  subject = $item.Subject; sender = $item.SenderName; senderEmail = $item.SenderEmailAddress
  received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm"); to = $item.To; cc = $item.CC
  body = $item.Body.Substring(0, [Math]::Min($item.Body.Length, 8000))
} | ConvertTo-Json -Depth 2
`
      try { return ok(ps(script, { ENTRY_ID: entry_id })) }
      catch (e) { return bad(e.message) }
    },
  )

  // ── Enviar correo (confirm) ─────────────────────────────────────
  server.tool(
    'outlook_send_email',
    'Envia un correo desde Outlook. Requiere confirm:true; sin el, devuelve un preview y no envia.',
    {
      to: z.string().describe('Destinatario(s), separados por punto y coma'),
      subject: z.string().describe('Asunto'),
      body: z.string().describe('Cuerpo'),
      cc: z.string().optional().describe('CC (opcional), separados por punto y coma'),
      confirm: z.boolean().optional().describe('Debe ser true para enviar'),
    },
    async (args) => {
      if (!isConfirmed(args)) {
        return ok(buildPreview('Enviar correo', {
          'Para': args.to, 'CC': args.cc, 'Asunto': args.subject, 'Cuerpo': args.body,
        }))
      }
      const script = `
$ol = New-Object -ComObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.To = $env:MAIL_TO
$mail.Subject = $env:MAIL_SUBJECT
$mail.Body = $env:MAIL_BODY
if ($env:MAIL_CC -and $env:MAIL_CC -ne "") { $mail.CC = $env:MAIL_CC }
$mail.Send()
Write-Output "Correo enviado a $env:MAIL_TO"
`
      try {
        return ok(ps(script, { MAIL_TO: args.to, MAIL_SUBJECT: args.subject, MAIL_BODY: args.body, MAIL_CC: args.cc ?? '' }))
      } catch (e) { return bad(e.message) }
    },
  )

  // ── Responder correo (confirm) ──────────────────────────────────
  server.tool(
    'outlook_reply_email',
    'Responde a un correo por su EntryID. Requiere confirm:true; sin el, devuelve un preview.',
    {
      entry_id: z.string().describe('EntryID del correo a responder'),
      body: z.string().describe('Texto de la respuesta'),
      replyAll: z.boolean().optional().describe('Responder a todos (default: solo al remitente)'),
      confirm: z.boolean().optional(),
    },
    async (args) => {
      if (!isConfirmed(args)) {
        return ok(buildPreview(args.replyAll ? 'Responder a todos' : 'Responder al remitente', {
          'Correo id': args.entry_id, 'Respuesta': args.body,
        }))
      }
      const method = args.replyAll ? 'ReplyAll' : 'Reply'
      const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$item = $ns.GetItemFromID($env:ENTRY_ID)
$reply = $item.${method}()
$reply.Body = $env:REPLY_BODY + "\`n\`n" + $reply.Body
$reply.Send()
Write-Output "Respuesta enviada."
`
      try { return ok(ps(script, { ENTRY_ID: args.entry_id, REPLY_BODY: args.body })) }
      catch (e) { return bad(e.message) }
    },
  )

  // ── Listar calendario ───────────────────────────────────────────
  server.tool(
    'outlook_list_calendar',
    'Lista los proximos eventos de tu calendario de Outlook.',
    { days: z.number().int().min(1).max(60).optional().describe('Dias hacia adelante (default 7)') },
    async ({ days }) => {
      const d = days ?? 7
      const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$calendar = $ns.GetDefaultFolder(9)
$items = $calendar.Items
$items.IncludeRecurrences = $true
$items.Sort("[Start]")
$startDate = (Get-Date).ToString("MM/dd/yyyy HH:mm")
$endDate = (Get-Date).AddDays(${d}).ToString("MM/dd/yyyy HH:mm")
$filter = "[Start] >= '$startDate' AND [Start] <= '$endDate'"
$filtered = $items.Restrict($filter)
$result = @()
foreach ($item in $filtered) {
  try {
    $result += [PSCustomObject]@{
      subject = $item.Subject
      start = $item.Start.ToString("yyyy-MM-dd HH:mm"); end = $item.End.ToString("yyyy-MM-dd HH:mm")
      location = $item.Location; organizer = $item.Organizer
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
      try { const ev = parseJson(ps(script)); return json({ count: ev.length, events: ev }) }
      catch (e) { return bad(e.message) }
    },
  )

  // ── Calendario compartido ───────────────────────────────────────
  server.tool(
    'outlook_shared_calendar',
    'Consulta el calendario de otro usuario de Outlook (debe tener permisos o el calendario ya anadido).',
    {
      user: z.string().describe('Nombre o email del usuario'),
      start_date: z.string().optional().describe('Inicio (yyyy-MM-dd), default: primer dia del mes'),
      end_date: z.string().optional().describe('Fin (yyyy-MM-dd), default: hoy'),
    },
    async (args) => {
      const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$recipient = $ns.CreateRecipient($env:CAL_USER)
$recipient.Resolve() | Out-Null
if (-not $recipient.Resolved) {
  Write-Output "ERROR: Usuario '$env:CAL_USER' no encontrado. Agrega su calendario en Outlook: Calendario > Abrir calendario > De la libreta de direcciones."
  exit 0
}
try { $calendar = $ns.GetSharedDefaultFolder($recipient, 9) }
catch { Write-Output "ERROR: Sin acceso al calendario de '$env:CAL_USER': $_"; exit 0 }
$startDT = [DateTime]::Parse($env:CAL_START)
$endDT = [DateTime]::Parse($env:CAL_END).Date.AddDays(1).AddSeconds(-1)
$filter = "[Start] >= '$($startDT.ToString("MM/dd/yyyy HH:mm"))' AND [Start] <= '$($endDT.ToString("MM/dd/yyyy HH:mm"))'"
$items1 = $calendar.Items; $items1.Sort("[Start]"); $r1 = $items1.Restrict($filter)
$items2 = $calendar.Items; $items2.Sort("[Start]"); $items2.IncludeRecurrences = $true; $r2 = $items2.Restrict($filter)
$seen = @{}; $result = @()
foreach ($item in (@($r1) + @($r2))) {
  try {
    $key = "$($item.Subject)|$($item.Start)"
    if ($seen[$key]) { continue }
    $seen[$key] = $true
    if ($item.Start -lt $startDT -or $item.Start -gt $endDT) { continue }
    $result += [PSCustomObject]@{
      subject = $item.Subject; start = $item.Start.ToString("yyyy-MM-dd HH:mm"); end = $item.End.ToString("yyyy-MM-dd HH:mm")
      location = $item.Location; organizer = $item.Organizer
      busyStatus = switch ($item.BusyStatus) { 0 {"Libre"} 1 {"Provisional"} 2 {"Ocupado"} 3 {"Fuera de oficina"} default {"Desconocido"} }
    }
  } catch {}
}
$sorted = $result | Sort-Object start
if ($sorted.Count -eq 0) { Write-Output "[]" } else { $sorted | ConvertTo-Json -Depth 2 }
`
      const today = new Date()
      const startDefault = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-01`
      const endDefault = today.toISOString().slice(0, 10)
      try {
        const out = ps(script, { CAL_USER: args.user, CAL_START: args.start_date ?? startDefault, CAL_END: args.end_date ?? endDefault })
        if (out.startsWith('ERROR:')) { return bad(out) }
        return ok(out)
      } catch (e) { return bad(e.message) }
    },
  )

  // ── Crear evento ────────────────────────────────────────────────
  server.tool(
    'outlook_create_event',
    'Crea un evento en tu calendario de Outlook (solo tuyo, no envia invitaciones).',
    {
      subject: z.string().describe('Titulo'),
      start: z.string().describe('Inicio (yyyy-MM-dd HH:mm)'),
      end: z.string().describe('Fin (yyyy-MM-dd HH:mm)'),
      location: z.string().optional().describe('Ubicacion'),
      body: z.string().optional().describe('Descripcion'),
    },
    async (args) => {
      const script = `
$ol = New-Object -ComObject Outlook.Application
$appt = $ol.CreateItem(1)
$appt.Subject = $env:EVT_SUBJECT
$appt.Start = [DateTime]::Parse($env:EVT_START)
$appt.End = [DateTime]::Parse($env:EVT_END)
$appt.Location = $env:EVT_LOCATION
$appt.Body = $env:EVT_BODY
$appt.Save()
Write-Output "Evento '$($appt.Subject)' creado: $($appt.Start.ToString('yyyy-MM-dd HH:mm')) - $($appt.End.ToString('HH:mm'))"
`
      try {
        return ok(ps(script, {
          EVT_SUBJECT: args.subject, EVT_START: args.start, EVT_END: args.end,
          EVT_LOCATION: args.location ?? '', EVT_BODY: args.body ?? '',
        }))
      } catch (e) { return bad(e.message) }
    },
  )

  // ── Buscar huecos libres ────────────────────────────────────────
  server.tool(
    'outlook_find_free',
    'Busca huecos libres en tu calendario (horario laboral 9-18) para los proximos dias.',
    {
      days: z.number().int().min(1).max(14).optional().describe('Dias a revisar (default 5)'),
      minMinutes: z.number().int().min(15).optional().describe('Duracion minima del hueco (default 30)'),
    },
    async ({ days, minMinutes }) => {
      try {
        const events = listCalendarRaw(days ?? 5)
        const min = minMinutes ?? 30
        const all = []
        for (let off = 0; off < (days ?? 5); off++) {
          const day = new Date(); day.setHours(0, 0, 0, 0); day.setDate(day.getDate() + off)
          const ws = new Date(day); ws.setHours(9, 0, 0, 0)
          const we = new Date(day); we.setHours(18, 0, 0, 0)
          const busy = events
            .map(e => ({ start: Date.parse(e.start), end: Date.parse(e.end) }))
            .filter(b => Number.isFinite(b.start) && Number.isFinite(b.end))
          const slots = freeSlots(ws.getTime(), we.getTime(), busy, min)
          for (const s of slots) { all.push({ start: new Date(s.start).toISOString(), end: new Date(s.end).toISOString() }) }
        }
        return json({ count: all.length, freeSlots: all })
      } catch (e) { return bad(e.message) }
    },
  )

  // ── Skill: briefing del dia ─────────────────────────────────────
  server.tool(
    'daily_briefing',
    'Resumen del dia: correos no leidos importantes + reuniones de hoy. Solo lectura.',
    {},
    async () => {
      try {
        const me = getMeCom()
        const emails = listEmailsRaw(25, 'Inbox').filter(m => m.unread)
        const classified = sortByPriority(emails.map(m => ({
          ...m,
          ...classifyEmail({
            subject: m.subject, bodyPreview: m.preview, isRead: !m.unread,
            toRecipients: splitAddrs(m.to), ccRecipients: splitAddrs(m.cc),
          }, me.email),
        })))
        const important = classified.filter(c => c.priority !== 'baja' || c.needsMyAction).slice(0, 10)
        const events = listCalendarRaw(1)
        return json({
          you: me.email || undefined,
          importantUnread: important.map(m => ({ from: m.senderEmail || m.sender, subject: m.subject, priority: m.priority, needsMyAction: m.needsMyAction, received: m.received })),
          totalUnread: emails.length,
          meetingsToday: events.map(e => ({ subject: e.subject, start: e.start, end: e.end, location: e.location })),
        })
      } catch (e) { return bad(e.message) }
    },
  )

  // ── Skill: triage de bandeja ────────────────────────────────────
  server.tool(
    'inbox_triage',
    'Clasifica tus correos no leidos por prioridad y marca los que piden accion tuya.',
    { count: z.number().int().min(1).max(50).optional().describe('Cuantos revisar (default 25)') },
    async ({ count }) => {
      try {
        const me = getMeCom()
        const emails = listEmailsRaw(count ?? 25, 'Inbox').filter(m => m.unread)
        const classified = sortByPriority(emails.map(m => ({
          ...m,
          ...classifyEmail({
            subject: m.subject, bodyPreview: m.preview, isRead: !m.unread,
            toRecipients: splitAddrs(m.to), ccRecipients: splitAddrs(m.cc),
          }, me.email),
        })))
        return json({
          count: classified.length,
          items: classified.map(m => ({ id: m.id, from: m.senderEmail || m.sender, subject: m.subject, priority: m.priority, needsMyAction: m.needsMyAction, reasons: m.reasons })),
        })
      } catch (e) { return bad(e.message) }
    },
  )
}

// ── Helpers internos reutilizados por las skills ──────────────────
function splitAddrs(s) {
  return String(s || '').split(';').map(x => x.trim()).filter(Boolean).map(a => ({ address: a.toLowerCase() }))
}

function listEmailsRaw(count, folder) {
  const folderMap = { Inbox: 6, SentItems: 5, Drafts: 16 }
  const folderNum = folderMap[folder ?? 'Inbox'] ?? 6
  const c = Math.min(count ?? 20, 50)
  const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$folder = $ns.GetDefaultFolder(${folderNum})
$items = $folder.Items
$items.Sort("[ReceivedTime]", $true)
$total = [Math]::Min($items.Count, ${c})
$result = @()
for ($i = 1; $i -le $total; $i++) {
  try {
    $item = $items.Item($i)
    $result += [PSCustomObject]@{
      id = $item.EntryID; subject = $item.Subject; sender = $item.SenderName; senderEmail = $item.SenderEmailAddress
      received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm"); unread = $item.UnRead; to = $item.To; cc = $item.CC
      preview = $item.Body.Substring(0, [Math]::Min($item.Body.Length, 200))
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
  const out = ps(script)
  if (!out || !out.trim()) { return [] }
  const j = JSON.parse(out)
  return Array.isArray(j) ? j : [j]
}

function listCalendarRaw(days) {
  const d = days ?? 7
  const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$calendar = $ns.GetDefaultFolder(9)
$items = $calendar.Items
$items.IncludeRecurrences = $true
$items.Sort("[Start]")
$startDate = (Get-Date).ToString("MM/dd/yyyy HH:mm")
$endDate = (Get-Date).AddDays(${d}).ToString("MM/dd/yyyy HH:mm")
$filter = "[Start] >= '$startDate' AND [Start] <= '$endDate'"
$filtered = $items.Restrict($filter)
$result = @()
foreach ($item in $filtered) {
  try {
    $result += [PSCustomObject]@{
      subject = $item.Subject; start = $item.Start.ToString("yyyy-MM-dd HH:mm"); end = $item.End.ToString("yyyy-MM-dd HH:mm")
      location = $item.Location; organizer = $item.Organizer
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
  const out = ps(script)
  if (!out || !out.trim()) { return [] }
  const j = JSON.parse(out)
  return Array.isArray(j) ? j : [j]
}
