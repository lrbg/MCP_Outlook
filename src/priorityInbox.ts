/**
 * Correos de remitentes prioritarios en las ultimas N semanas, desde el Outlook
 * de escritorio (COM). Resuelve la direccion SMTP real de remitentes internos de
 * Exchange (que llegan como DN X.500) para poder comparar por email.
 */
import { execSync } from 'node:child_process'
import { RawEmail } from './bitacoraCore'
import { parsePsArray, dec, PS_B64_FN } from './psJson'

export const isWindows = process.platform === 'win32'

export interface PriorityEmail extends RawEmail {
  id: string
  to?: string
  cc?: string
  bodySnippet?: string
  /** Fecha de tu respuesta en Enviados (misma conversacion), vacia si no respondiste. */
  repliedAt?: string
  /** Texto de tu respuesta (recortado). */
  replyBody?: string
}

function ps(script: string, env: Record<string, string>): string {
  const full = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(full, 'utf16le').toString('base64')
  return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, {
    encoding: 'utf8', timeout: 90000, env: { ...process.env, ...env },
  }).trim()
}

/**
 * Devuelve hasta `max` correos recibidos en los ultimos `days` dias cuyo
 * remitente (email) este en `senders`. Ordenados del mas reciente al mas antiguo.
 */
export function getPriorityEmails(senders: string[], days = 14, max = 20): PriorityEmail[] {
  if (!isWindows) {
    throw new Error('Requiere Windows con Outlook de escritorio.')
  }
  const list = senders.map(s => s.trim().toLowerCase()).filter(Boolean)
  if (list.length === 0) { return [] }

  const script = `
${PS_B64_FN}
$set = @{}
foreach ($e in ($env:PRIORITY_SENDERS -split ',')) { $t = $e.Trim().ToLower(); if ($t) { $set[$t] = $true } }
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")

# Mapa de respuestas enviadas por conversacion (para saber si ya respondi).
$sent = @{}
try {
  $sf = $ns.GetDefaultFolder(5)
  $si = $sf.Items
  $si.Sort("[SentOn]", $true)
  $scut = (Get-Date).AddDays(-(${days} + 3)).ToString("MM/dd/yyyy HH:mm")
  $sr = $si.Restrict("[SentOn] >= '$scut'")
  $sc = 0
  foreach ($it in $sr) {
    if ($sc -ge 800) { break }
    $sc++
    try {
      $cid = [string]$it.ConversationID
      if ($cid -and -not $sent.ContainsKey($cid)) {
        $sent[$cid] = [PSCustomObject]@{ at = $it.SentOn.ToString("yyyy-MM-dd HH:mm"); body = [string]$it.Body }
      }
    } catch {}
  }
} catch {}

$folder = $ns.GetDefaultFolder(6)
$items = $folder.Items
$items.Sort("[ReceivedTime]", $true)
$cut = (Get-Date).AddDays(-${days}).ToString("MM/dd/yyyy HH:mm")
$recent = $items.Restrict("[ReceivedTime] >= '$cut'")
$result = @()
$seen = 0
foreach ($item in $recent) {
  if ($seen -ge 800) { break }
  $seen++
  try {
    $smtp = $item.SenderEmailAddress
    if ($item.SenderEmailType -eq 'EX') {
      try { $ex = $item.Sender.GetExchangeUser(); if ($ex) { $smtp = $ex.PrimarySmtpAddress } } catch {}
    }
    $key = ([string]$smtp).ToLower()
    if ($set.ContainsKey($key)) {
      $body = $item.Body; if ($body.Length -gt 1500) { $body = $body.Substring(0, 1500) }
      $ra = ""; $rb = ""
      try {
        $cid = [string]$item.ConversationID
        if ($cid -and $sent.ContainsKey($cid)) {
          $s = $sent[$cid]
          $ra = $s.at
          $rb = $s.body; if ($rb.Length -gt 1200) { $rb = $rb.Substring(0, 1200) }
        }
      } catch {}
      $result += [PSCustomObject]@{
        id = $item.EntryID
        senderEmail = $smtp
        received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        unread = $item.UnRead
        s = B64($item.Subject)
        n = B64($item.SenderName)
        t = B64($item.To)
        c = B64($item.CC)
        b = B64($body)
        ra = B64($ra)
        rb = B64($rb)
      }
      if ($result.Count -ge ${max}) { break }
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
  const raw = parsePsArray(ps(script, { PRIORITY_SENDERS: list.join(',') }))
  return raw.map(r => ({
    id: r.id, senderEmail: r.senderEmail, received: r.received, unread: r.unread,
    subject: dec(r.s), sender: dec(r.n), to: dec(r.t), cc: dec(r.c), bodySnippet: dec(r.b),
    repliedAt: dec(r.ra), replyBody: dec(r.rb),
  })) as PriorityEmail[]
}
