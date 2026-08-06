/**
 * Lectura de la agenda desde la extension (Outlook COM). Proximas reuniones en
 * los siguientes N dias.
 */
import { execSync } from 'node:child_process'
import { Meeting } from './agendaCore'
import { parsePsArray, dec, PS_B64_FN } from './psJson'

export const isWindows = process.platform === 'win32'

function ps(script: string): string {
  const full = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(full, 'utf16le').toString('base64')
  return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, { encoding: 'utf8', timeout: 150000 }).trim()
}

/** Reuniones desde ahora hasta dentro de `days` dias. */
export function getMeetings(days = 7): Meeting[] {
  if (!isWindows) { throw new Error('Requiere Windows con Outlook de escritorio.') }
  const d = Math.min(Math.max(days, 1), 60)
  const script = `
${PS_B64_FN}
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
      start = $item.Start.ToString("yyyy-MM-dd HH:mm")
      end = $item.End.ToString("yyyy-MM-dd HH:mm")
      attendees = @($item.Recipients).Count
      s = B64($item.Subject); l = B64($item.Location); o = B64($item.Organizer)
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
  return parsePsArray(ps(script)).map(r => ({
    start: r.start, end: r.end, attendees: r.attendees,
    subject: dec(r.s), location: dec(r.l), organizer: dec(r.o),
  })) as Meeting[]
}
