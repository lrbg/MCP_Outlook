/**
 * Acciones sobre un correo concreto desde la extension (Outlook COM): leer el
 * cuerpo completo y responder. Datos por variables de entorno (sin inyeccion).
 */
import { execSync } from 'node:child_process'

export const isWindows = process.platform === 'win32'

function ps(script: string, env: Record<string, string> = {}): string {
  const full = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(full, 'utf16le').toString('base64')
  return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, {
    encoding: 'utf8', timeout: 60000, env: { ...process.env, ...env },
  }).trim()
}

export interface EmailBody {
  subject: string
  sender: string
  senderEmail: string
  to: string
  cc: string
  received: string
  body: string
}

/** Lee el correo por EntryID (cuerpo recortado a 8000 caracteres). */
export function readEmailBody(entryId: string): EmailBody {
  if (!isWindows) { throw new Error('Requiere Windows con Outlook de escritorio.') }
  const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$item = $ns.GetItemFromID($env:ENTRY_ID)
[PSCustomObject]@{
  subject = $item.Subject; sender = $item.SenderName; senderEmail = $item.SenderEmailAddress
  to = $item.To; cc = $item.CC; received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
  body = $item.Body.Substring(0, [Math]::Min($item.Body.Length, 8000))
} | ConvertTo-Json -Depth 2
`
  return JSON.parse(ps(script, { ENTRY_ID: entryId }))
}

/** Responde el correo por EntryID con el texto dado. */
export function sendReply(entryId: string, body: string, replyAll = false): string {
  if (!isWindows) { throw new Error('Requiere Windows con Outlook de escritorio.') }
  const method = replyAll ? 'ReplyAll' : 'Reply'
  const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$item = $ns.GetItemFromID($env:ENTRY_ID)
$reply = $item.${method}()
$reply.Body = $env:REPLY_BODY + "\`n\`n" + $reply.Body
$reply.Send()
Write-Output "enviado"
`
  ps(script, { ENTRY_ID: entryId, REPLY_BODY: body })
  return 'enviado'
}
