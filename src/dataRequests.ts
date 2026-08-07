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
  conv: string
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

/**
 * Para cada EntryID dado, arma un "digest" del hilo de conversacion (mensajes de
 * bandeja de entrada + enviados de la misma conversacion), para que la IA infiera
 * el estado real de la solicitud a partir del historial.
 */
export function getThreadDigests(entryIds: string[], daysBack = 60): { id: string; digest: string }[] {
  if (!isWindows) { throw new Error('Requiere Windows con Outlook de escritorio.') }
  const ids = entryIds.filter(Boolean)
  if (!ids.length) { return [] }
  const d = Math.min(Math.max(daysBack, 7), 120)
  const script = `
${PS_B64_FN}
$ids = $env:IDS -split ','
$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI")
$want = @{}
$idToCid = @{}
foreach ($id in $ids) {
  try { $it = $ns.GetItemFromID($id); $cid = [string]$it.ConversationID; if ($cid) { $want[$cid] = $true; $idToCid[$id] = $cid } } catch {}
}
$msgs = @{}
$cut = (Get-Date).AddDays(-${d}).ToString("MM/dd/yyyy HH:mm")
foreach ($fn in @(6,5)) {
  try {
    $folder = $ns.GetDefaultFolder($fn)
    $items = $folder.Items
    if ($fn -eq 5) { $items.Sort("[SentOn]", $true); $sub = $items.Restrict("[SentOn] >= '$cut'") }
    else { $items.Sort("[ReceivedTime]", $true); $sub = $items.Restrict("[ReceivedTime] >= '$cut'") }
    $n = 0
    foreach ($it in $sub) {
      if ($n -ge 1500) { break }
      $n++
      try {
        $cid = [string]$it.ConversationID
        if ($cid -and $want.ContainsKey($cid)) {
          $when = ""; $who = ""
          if ($fn -eq 5) { try { $when = $it.SentOn.ToString("yyyy-MM-dd HH:mm") } catch {}; $who = "YO (enviado)" }
          else { try { $when = $it.ReceivedTime.ToString("yyyy-MM-dd HH:mm") } catch {}; $who = [string]$it.SenderName }
          $bd = [string]$it.Body; if ($bd.Length -gt 400) { $bd = $bd.Substring(0,400) }
          if (-not $msgs.ContainsKey($cid)) { $msgs[$cid] = @() }
          $msgs[$cid] += ("[" + $when + "] " + $who + ": " + $bd)
        }
      } catch {}
    }
  } catch {}
}
$result = @()
foreach ($id in $ids) {
  $cid = $idToCid[$id]
  $lines = ""
  if ($cid -and $msgs.ContainsKey($cid)) { $lines = (($msgs[$cid] | Select-Object -First 12) -join "  ||  ") }
  $result += [PSCustomObject]@{ id = $id; d = B64($lines) }
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
  return parsePsArrayEnv(script, { IDS: ids.join(',') }).map((r: any) => ({ id: r.id, digest: dec(r.d) })) as { id: string; digest: string }[]
}

/** Extrae un campo "Etiqueta: valor" hasta la siguiente etiqueta, coma o salto. */
function field(body: string, label: string): string {
  const stop = '(?=,|\\n|\\r|Solicitante:|Equipo:|Proyecto:|Fecha|$)'
  const m = body.match(new RegExp(label + '\\s*:\\s*(.+?)' + stop, 'i'))
  return m ? m[1].trim().replace(/\s+/g, ' ') : ''
}

/**
 * Lee las solicitudes cuyo asunto contiene `keyword` (ej. "datos sint" o
 * "performance") de los ultimos `days` dias, y parsea sus campos.
 */
export function getRequests(keyword: string, days = 45, max = 150): DataRequest[] {
  if (!isWindows) { throw new Error('Requiere Windows con Outlook de escritorio.') }
  const d = Math.min(Math.max(Math.ceil(days), 1), 120)
  const kw = norm(keyword)
  const script = `
${PS_B64_FN}
$kw = $env:KW.ToLower()
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
  if ($seen -ge 1500) { break }
  $seen++
  try {
    $subj = [string]$item.Subject
    if ($subj.ToLower().Contains($kw)) {
      $body = $item.Body; if ($body.Length -gt 2500) { $body = $body.Substring(0, 2500) }
      $result += [PSCustomObject]@{
        id = $item.EntryID
        conv = [string]$item.ConversationID
        received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        s = B64($subj); b = B64($body)
      }
      if ($result.Count -ge ${max}) { break }
    }
  } catch {}
}
if ($result.Count -eq 0) { Write-Output "[]" } else { $result | ConvertTo-Json -Depth 2 }
`
  return parsePsArrayEnv(script, { KW: keyword }).map((r: any) => {
    const subject = dec(r.s)
    const body = dec(r.b)
    return {
      id: r.id, conv: r.conv || '', received: r.received, subject,
      solicitante: field(body, 'Solicitante'),
      equipo: field(body, 'Equipo'),
      proyecto: field(body, 'Proyecto'),
      fecha: field(body, 'Fecha\\s+solic\\w*'),
    }
  }).filter((x: DataRequest) => norm(x.subject).includes(kw)) as DataRequest[]
}

function parsePsArrayEnv(script: string, env: Record<string, string>): any[] {
  const full = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(full, 'utf16le').toString('base64')
  const out = execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, {
    encoding: 'utf8', timeout: 120000, env: { ...process.env, ...env },
  }).trim()
  return parsePsArray(out)
}
