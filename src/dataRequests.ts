/**
 * Solicitudes de Datos Sinteticos: detecta los correos-plantilla en la bandeja
 * (asunto "Nueva solicitud de Datos Sinteticos - Web / Automatizacion E2E", de
 * Ilver Penaloza) y extrae Solicitante / Equipo / Proyecto / Fecha solicitada.
 */
import { execSync } from 'node:child_process'
import { parsePsArray, dec, PS_B64_FN } from './psJson'

export const isWindows = process.platform === 'win32'

export interface DataRequest {
  id: string
  received: string
  subject: string
  solicitante: string
  equipo: string
  proyecto: string
  fecha: string
}

function ps(script: string): string {
  const full = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(full, 'utf16le').toString('base64')
  return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, { encoding: 'utf8', timeout: 120000 }).trim()
}

/** Quita acentos y baja a minusculas para comparar. */
function norm(s: string): string {
  return (s || '').normalize('NFD').replace(new RegExp('[\\u0300-\\u036f]', 'g'), '').toLowerCase()
}

/** Extrae un campo "Etiqueta: valor" hasta la siguiente etiqueta, coma o salto. */
function field(body: string, label: string): string {
  const stop = '(?=,|\\n|\\r|Solicitante:|Equipo:|Proyecto:|Fecha|$)'
  const m = body.match(new RegExp(label + '\\s*:\\s*(.+?)' + stop, 'i'))
  return m ? m[1].trim().replace(/\s+/g, ' ') : ''
}

/** Lee las solicitudes de datos sinteticos de los ultimos `days` dias. */
export function getDataRequests(days = 30, max = 100): DataRequest[] {
  if (!isWindows) { throw new Error('Requiere Windows con Outlook de escritorio.') }
  const d = Math.min(Math.max(Math.ceil(days), 1), 120)
  const script = `
${PS_B64_FN}
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$folder = $ns.GetDefaultFolder(6)
$items = $folder.Items
$items.Sort("[ReceivedTime]", $true)
$cut = (Get-Date).AddDays(-${d}).ToString("MM/dd/yyyy HH:mm")
$recent = $items.Restrict("[ReceivedTime] >= '$cut'")
$result = @()
$seen = 0
foreach ($item in $recent) {
  if ($seen -ge 1000) { break }
  $seen++
  try {
    $subj = [string]$item.Subject
    if ($subj -match 'Datos Sint') {
      $body = $item.Body; if ($body.Length -gt 2500) { $body = $body.Substring(0, 2500) }
      $result += [PSCustomObject]@{
        id = $item.EntryID
        received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        s = B64($subj); b = B64($body)
      }
      if ($result.Count -ge ${max}) { break }
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
  return parsePsArray(ps(script)).map(r => {
    const subject = dec(r.s)
    const body = dec(r.b)
    return {
      id: r.id,
      received: r.received,
      subject,
      solicitante: field(body, 'Solicitante'),
      equipo: field(body, 'Equipo'),
      proyecto: field(body, 'Proyecto'),
      fecha: field(body, 'Fecha\\s+solic\\w*'),
    }
  }).filter(x => norm(x.subject).includes('datos sint')) as DataRequest[]
}
