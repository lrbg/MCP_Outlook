/**
 * Lectura de carpetas (bandeja de entrada / enviados) desde la extension
 * (Windows, Outlook por COM). Se usa para la revision diaria. Textos por Base64
 * para no romper JSON.
 */
import { execSync } from 'node:child_process'
import { RawEmail } from './bitacoraCore'
import { parsePsArray, dec, PS_B64_FN } from './psJson'

export const isWindows = process.platform === 'win32'

function ps(script: string): string {
  const full = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(full, 'utf16le').toString('base64')
  return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, {
    encoding: 'utf8', timeout: 90000,
  }).trim()
}

/**
 * Correos recientes de una carpeta. folderNum: 6 = Bandeja de entrada, 5 = Enviados.
 * dateProp: propiedad de fecha (ReceivedTime para entrada, SentOn para enviados).
 */
export function getRecent(folderNum: number, max = 40, dateProp = 'ReceivedTime'): RawEmail[] {
  if (!isWindows) { throw new Error('Requiere Windows con Outlook de escritorio.') }
  const n = Math.min(Math.max(max, 1), 100)
  const script = `
${PS_B64_FN}
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$folder = $ns.GetDefaultFolder(${folderNum})
$items = $folder.Items
$items.Sort("[${dateProp}]", $true)
$total = [Math]::Min($items.Count, ${n})
$result = @()
for ($i = 1; $i -le $total; $i++) {
  try {
    $item = $items.Item($i)
    $prev = $item.Body; if ($prev.Length -gt 300) { $prev = $prev.Substring(0, 300) }
    $when = ""
    try { $when = $item.${dateProp}.ToString("yyyy-MM-dd HH:mm") } catch {}
    $result += [PSCustomObject]@{
      received = $when
      unread = $item.UnRead
      s = B64($item.Subject); n = B64($item.SenderName); e = B64($item.SenderEmailAddress)
      t = B64($item.To); p = B64($prev)
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
  return parsePsArray(ps(script)).map(r => ({
    received: r.received, unread: r.unread,
    subject: dec(r.s), sender: dec(r.n), senderEmail: dec(r.e), to: dec(r.t), preview: dec(r.p),
  })) as RawEmail[]
}
