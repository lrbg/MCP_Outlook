/**
 * Lectura de la bandeja desde la extension (Windows, Outlook por COM via
 * PowerShell). Se usa para la revision diaria automatica. Misma tecnica que el
 * servidor MCP: script codificado y sin datos interpolados.
 */
import { execSync } from 'node:child_process'
import { RawEmail } from './bitacoraCore'
import { parsePsArray, dec, PS_B64_FN } from './psJson'

export const isWindows = process.platform === 'win32'

function ps(script: string): string {
  const full = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(full, 'utf16le').toString('base64')
  return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, {
    encoding: 'utf8', timeout: 60000,
  }).trim()
}

/** Devuelve los correos NO LEIDOS de la bandeja (hasta `max`). */
export function getUnread(max = 30): RawEmail[] {
  if (!isWindows) {
    throw new Error('La revision usa Outlook de escritorio (COM) y requiere Windows.')
  }
  const n = Math.min(Math.max(max, 1), 100)
  const script = `
${PS_B64_FN}
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$folder = $ns.GetDefaultFolder(6)
$items = $folder.Items
$items.Sort("[ReceivedTime]", $true)
$unread = $items.Restrict("[UnRead] = true")
$total = [Math]::Min($unread.Count, ${n})
$result = @()
for ($i = 1; $i -le $total; $i++) {
  try {
    $item = $unread.Item($i)
    $prev = $item.Body; if ($prev.Length -gt 300) { $prev = $prev.Substring(0, 300) }
    $result += [PSCustomObject]@{
      received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
      unread = $true
      s = B64($item.Subject); n = B64($item.SenderName); e = B64($item.SenderEmailAddress); p = B64($prev)
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
  return parsePsArray(ps(script)).map(r => ({
    received: r.received, unread: r.unread,
    subject: dec(r.s), sender: dec(r.n), senderEmail: dec(r.e), preview: dec(r.p),
  })) as RawEmail[]
}
