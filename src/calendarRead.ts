/**
 * Lectura de la agenda desde la extension (Outlook COM). Proximas reuniones en
 * los siguientes N dias.
 */
import { execSync } from 'node:child_process'
import { Meeting } from './agendaCore'

export const isWindows = process.platform === 'win32'

function ps(script: string): string {
  const full = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(full, 'utf16le').toString('base64')
  return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, { encoding: 'utf8', timeout: 60000 }).trim()
}

/** Reuniones desde ahora hasta dentro de `days` dias. */
export function getMeetings(days = 7): Meeting[] {
  if (!isWindows) { throw new Error('Requiere Windows con Outlook de escritorio.') }
  const d = Math.min(Math.max(days, 1), 60)
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
      start = $item.Start.ToString("yyyy-MM-dd HH:mm")
      end = $item.End.ToString("yyyy-MM-dd HH:mm")
      location = $item.Location
      organizer = $item.Organizer
      attendees = @($item.Recipients).Count
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
