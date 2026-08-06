/**
 * Acciones sobre un correo concreto desde la extension (Outlook COM): leer el
 * cuerpo completo y responder. Datos por variables de entorno (sin inyeccion).
 */
import { execSync } from 'node:child_process'
import { cleanPsJson, dec, PS_B64_FN } from './psJson'

export const isWindows = process.platform === 'win32'

function ps(script: string, env: Record<string, string> = {}): string {
  const full = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(full, 'utf16le').toString('base64')
  return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, {
    encoding: 'utf8', timeout: 60000, env: { ...process.env, ...env },
  }).trim()
}

/** Identidad del usuario segun el Outlook local (nombre + SMTP). */
let cachedMe: { name: string; email: string } | undefined
export function getMe(): { name: string; email: string } {
  if (cachedMe) { return cachedMe }
  if (!isWindows) { return { name: '', email: '' } }
  const script = `
${PS_B64_FN}
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$u = $ns.CurrentUser
$email = ""
try { $email = $u.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch {}
if (-not $email) { try { $email = $u.Address } catch {} }
[PSCustomObject]@{ n = B64($u.Name); e = B64($email) } | ConvertTo-Json
`
  try {
    const j = JSON.parse(cleanPsJson(ps(script)))
    cachedMe = { name: dec(j.n), email: dec(j.e).toLowerCase() }
  } catch { cachedMe = { name: '', email: '' } }
  return cachedMe
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
${PS_B64_FN}
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$item = $ns.GetItemFromID($env:ENTRY_ID)
$body = $item.Body; if ($body.Length -gt 8000) { $body = $body.Substring(0, 8000) }
[PSCustomObject]@{
  s = B64($item.Subject); n = B64($item.SenderName); e = B64($item.SenderEmailAddress)
  t = B64($item.To); c = B64($item.CC); received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
  b = B64($body)
} | ConvertTo-Json -Depth 2
`
  const r = JSON.parse(cleanPsJson(ps(script, { ENTRY_ID: entryId })))
  return { subject: dec(r.s), sender: dec(r.n), senderEmail: dec(r.e), to: dec(r.t), cc: dec(r.c), received: r.received, body: dec(r.b) }
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
