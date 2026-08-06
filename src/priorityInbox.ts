/**
 * Correos de remitentes prioritarios en las ultimas N semanas, desde el Outlook
 * de escritorio (COM). Resuelve la direccion SMTP real de remitentes internos de
 * Exchange (que llegan como DN X.500) para poder comparar por email.
 */
import { execSync } from 'node:child_process'
import { RawEmail } from './bitacoraCore'

export const isWindows = process.platform === 'win32'

export interface PriorityEmail extends RawEmail {
  id: string
  to?: string
  cc?: string
  bodySnippet?: string
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
$set = @{}
foreach ($e in ($env:PRIORITY_SENDERS -split ',')) { $t = $e.Trim().ToLower(); if ($t) { $set[$t] = $true } }
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
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
      $result += [PSCustomObject]@{
        id = $item.EntryID
        subject = $item.Subject
        sender = $item.SenderName
        senderEmail = $smtp
        received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        unread = $item.UnRead
        to = $item.To
        cc = $item.CC
        bodySnippet = $item.Body.Substring(0, [Math]::Min($item.Body.Length, 1500))
      }
      if ($result.Count -ge ${max}) { break }
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { ($result | ConvertTo-Json -Depth 2) -replace '[\x00-\x1F]', ' ' }
`
  const out = ps(script, { PRIORITY_SENDERS: list.join(',') })
  if (!out || !out.trim()) { return [] }
  const j = JSON.parse(out)
  return Array.isArray(j) ? j : [j]
}
