#!/usr/bin/env node
/**
 * Servidor MCP de Outlook de ESCRITORIO (COM via PowerShell).
 * No usa Microsoft Graph ni Entra ID: maneja el Outlook ya abierto y firmado,
 * asi que evita el AADSTS65002 / consentimiento de admin. Windows-only.
 */
import { Server } from '@modelcontextprotocol/sdk/server/index.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ErrorCode,
  McpError,
} from '@modelcontextprotocol/sdk/types.js'
import { execSync } from 'child_process'

// ─── Helper: run PowerShell con datos de usuario por env vars (sin inyeccion) ──
function ps(script, env = {}) {
  const fullScript = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(fullScript, 'utf16le').toString('base64')
  try {
    return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, {
      encoding: 'utf8', timeout: 30000, env: { ...process.env, ...env }, stdio: ['pipe', 'pipe', 'pipe'],
    }).trim()
  } catch (err) {
    const raw = err.stderr ?? err.stdout ?? err.message ?? ''
    const xmlMatch = raw.match(/<S S="Error">([\s\S]*?)<\/S>/)
    const msg = xmlMatch
      ? xmlMatch[1].replace(/&#x([0-9A-F]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16)))
      : (raw.replace(/#< CLIXML[\s\S]*/, '').trim() || err.message)
    throw new Error(msg)
  }
}

// ─── Servidor MCP ──────────────────────────────────────────────────────────────
const server = new Server(
  { name: 'outlook-mcp', version: '1.0.0' },
  { capabilities: { tools: {} } },
)

// ─── Lista de herramientas disponibles ────────────────────────────────────────
server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: 'outlook_list_emails',
      description: 'Lista los correos recientes de una carpeta de Outlook',
      inputSchema: {
        type: 'object',
        properties: {
          count: { type: 'number', description: 'Cantidad de correos a obtener (default: 20, max: 50)' },
          folder: { type: 'string', enum: ['Inbox', 'SentItems', 'Drafts'], description: 'Carpeta (default: Inbox)' },
        },
      },
    },
    {
      name: 'outlook_read_email',
      description: 'Lee el contenido completo de un correo por su EntryID',
      inputSchema: {
        type: 'object',
        properties: {
          entry_id: { type: 'string', description: 'El EntryID del correo (obtenido de outlook_list_emails)' },
        },
        required: ['entry_id'],
      },
    },
    {
      name: 'outlook_send_email',
      description: 'Envia un correo desde Outlook',
      inputSchema: {
        type: 'object',
        properties: {
          to: { type: 'string', description: 'Destinatario(s), separados por punto y coma' },
          subject: { type: 'string', description: 'Asunto del correo' },
          body: { type: 'string', description: 'Cuerpo del correo' },
          cc: { type: 'string', description: 'Copia a (opcional), separados por punto y coma' },
        },
        required: ['to', 'subject', 'body'],
      },
    },
    {
      name: 'outlook_list_calendar',
      description: 'Lista los proximos eventos del calendario de Outlook',
      inputSchema: {
        type: 'object',
        properties: {
          days: { type: 'number', description: 'Dias hacia adelante a consultar (default: 7)' },
        },
      },
    },
    {
      name: 'outlook_shared_calendar',
      description: 'Consulta el calendario publico de otro usuario de Outlook (debe tener permisos o el calendario ya anadido)',
      inputSchema: {
        type: 'object',
        properties: {
          user: { type: 'string', description: 'Nombre o email del usuario cuyo calendario consultar' },
          start_date: { type: 'string', description: 'Fecha de inicio (yyyy-MM-dd), default: primer dia del mes actual' },
          end_date: { type: 'string', description: 'Fecha de fin (yyyy-MM-dd), default: hoy' },
        },
        required: ['user'],
      },
    },
    {
      name: 'outlook_create_event',
      description: 'Crea un nuevo evento en el calendario de Outlook',
      inputSchema: {
        type: 'object',
        properties: {
          subject: { type: 'string', description: 'Titulo del evento' },
          start: { type: 'string', description: 'Fecha y hora de inicio (yyyy-MM-dd HH:mm)' },
          end: { type: 'string', description: 'Fecha y hora de fin (yyyy-MM-dd HH:mm)' },
          location: { type: 'string', description: 'Ubicacion (opcional)' },
          body: { type: 'string', description: 'Descripcion (opcional)' },
        },
        required: ['subject', 'start', 'end'],
      },
    },
  ],
}))

// ─── Ejecucion de herramientas ─────────────────────────────────────────────────
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params

  try {
    switch (name) {
      // ── Listar correos ───────────────────────────────────────────────────────
      case 'outlook_list_emails': {
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
      id       = $item.EntryID
      subject  = $item.Subject
      sender   = $item.SenderName
      received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
      unread   = $item.UnRead
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
        return { content: [{ type: 'text', text: ps(script) }] }
      }

      // ── Leer correo ─────────────────────────────────────────────────────────
      case 'outlook_read_email': {
        const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$item = $ns.GetItemFromID($env:ENTRY_ID)
[PSCustomObject]@{
  subject     = $item.Subject
  sender      = $item.SenderName
  senderEmail = $item.SenderEmailAddress
  received    = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
  to          = $item.To
  cc          = $item.CC
  body        = $item.Body.Substring(0, [Math]::Min($item.Body.Length, 5000))
} | ConvertTo-Json -Depth 2
`
        return { content: [{ type: 'text', text: ps(script, { ENTRY_ID: args.entry_id }) }] }
      }

      // ── Enviar correo ────────────────────────────────────────────────────────
      case 'outlook_send_email': {
        const script = `
$ol = New-Object -ComObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.To      = $env:MAIL_TO
$mail.Subject = $env:MAIL_SUBJECT
$mail.Body    = $env:MAIL_BODY
if ($env:MAIL_CC -and $env:MAIL_CC -ne "") { $mail.CC = $env:MAIL_CC }
$mail.Send()
Write-Output "Correo enviado correctamente a $env:MAIL_TO"
`
        return {
          content: [{
            type: 'text',
            text: ps(script, {
              MAIL_TO: args.to, MAIL_SUBJECT: args.subject, MAIL_BODY: args.body, MAIL_CC: args.cc ?? '',
            }),
          }],
        }
      }

      // ── Listar calendario ────────────────────────────────────────────────────
      case 'outlook_list_calendar': {
        const days = args?.days ?? 7
        const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$calendar = $ns.GetDefaultFolder(9)
$items = $calendar.Items
$items.IncludeRecurrences = $true
$items.Sort("[Start]")
$startDate = (Get-Date).ToString("MM/dd/yyyy HH:mm")
$endDate   = (Get-Date).AddDays(${days}).ToString("MM/dd/yyyy HH:mm")
$filter    = "[Start] >= '$startDate' AND [Start] <= '$endDate'"
$filtered  = $items.Restrict($filter)
$result = @()
foreach ($item in $filtered) {
  try {
    $result += [PSCustomObject]@{
      subject   = $item.Subject
      start     = $item.Start.ToString("yyyy-MM-dd HH:mm")
      end       = $item.End.ToString("yyyy-MM-dd HH:mm")
      location  = $item.Location
      organizer = $item.Organizer
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
        return { content: [{ type: 'text', text: ps(script) }] }
      }

      // ── Calendario compartido de otro usuario ────────────────────────────────
      case 'outlook_shared_calendar': {
        const script = `
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$recipient = $ns.CreateRecipient($env:CAL_USER)
$recipient.Resolve() | Out-Null
if (-not $recipient.Resolved) {
  Write-Output "ERROR: Usuario '$env:CAL_USER' no encontrado. Agrega su calendario en Outlook: Calendario > Abrir calendario > De la libreta de direcciones."
  exit 0
}
try {
  $calendar = $ns.GetSharedDefaultFolder($recipient, 9)
} catch {
  Write-Output "ERROR: Sin acceso al calendario de '$env:CAL_USER': $_"
  exit 0
}
$startDT  = [DateTime]::Parse($env:CAL_START)
$endDT    = [DateTime]::Parse($env:CAL_END).Date.AddDays(1).AddSeconds(-1)
$startFmt = $startDT.ToString("MM/dd/yyyy HH:mm")
$endFmt   = $endDT.ToString("MM/dd/yyyy HH:mm")
$filter   = "[Start] >= '$startFmt' AND [Start] <= '$endFmt'"

# Paso 1: eventos no recurrentes en rango
$items1 = $calendar.Items
$items1.Sort("[Start]")
$restricted1 = $items1.Restrict($filter)

# Paso 2: instancias recurrentes expandidas en rango
$items2 = $calendar.Items
$items2.Sort("[Start]")
$items2.IncludeRecurrences = $true
$restricted2 = $items2.Restrict($filter)

$seen   = @{}
$result = @()
foreach ($item in (@($restricted1) + @($restricted2))) {
  try {
    $key = "$($item.Subject)|$($item.Start)"
    if ($seen[$key]) { continue }
    $seen[$key] = $true
    if ($item.Start -lt $startDT -or $item.Start -gt $endDT) { continue }
    $result += [PSCustomObject]@{
      subject    = $item.Subject
      start      = $item.Start.ToString("yyyy-MM-dd HH:mm")
      end        = $item.End.ToString("yyyy-MM-dd HH:mm")
      location   = $item.Location
      organizer  = $item.Organizer
      busyStatus = switch ($item.BusyStatus) {
        0 { "Libre" } 1 { "Provisional" } 2 { "Ocupado" } 3 { "Fuera de oficina" } default { "Desconocido" }
      }
    }
  } catch {}
}
$sorted = $result | Sort-Object start
if ($sorted.Count -eq 0) { Write-Output "[]" } else { $sorted | ConvertTo-Json -Depth 2 }
`
        const today = new Date()
        const startDefault = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-01`
        const endDefault = today.toISOString().slice(0, 10)
        return {
          content: [{
            type: 'text',
            text: ps(script, {
              CAL_USER: args.user, CAL_START: args.start_date ?? startDefault, CAL_END: args.end_date ?? endDefault,
            }),
          }],
        }
      }

      // ── Crear evento ─────────────────────────────────────────────────────────
      case 'outlook_create_event': {
        const script = `
$ol   = New-Object -ComObject Outlook.Application
$appt = $ol.CreateItem(1)
$appt.Subject  = $env:EVT_SUBJECT
$appt.Start    = [DateTime]::Parse($env:EVT_START)
$appt.End      = [DateTime]::Parse($env:EVT_END)
$appt.Location = $env:EVT_LOCATION
$appt.Body     = $env:EVT_BODY
$appt.Save()
Write-Output "Evento '$($appt.Subject)' creado: $($appt.Start.ToString('yyyy-MM-dd HH:mm')) - $($appt.End.ToString('HH:mm'))"
`
        return {
          content: [{
            type: 'text',
            text: ps(script, {
              EVT_SUBJECT: args.subject, EVT_START: args.start, EVT_END: args.end,
              EVT_LOCATION: args.location ?? '', EVT_BODY: args.body ?? '',
            }),
          }],
        }
      }

      default:
        throw new McpError(ErrorCode.MethodNotFound, `Herramienta desconocida: ${name}`)
    }
  } catch (err) {
    if (err instanceof McpError) throw err
    throw new McpError(ErrorCode.InternalError, `Error de Outlook: ${err.stderr ?? err.message}`)
  }
})

// ─── Arrancar servidor ─────────────────────────────────────────────────────────
async function main() {
  const transport = new StdioServerTransport()
  await server.connect(transport)
}

main()
