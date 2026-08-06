/**
 * Puente a Outlook de ESCRITORIO por COM, via PowerShell. NO usa Microsoft Graph
 * ni Entra ID: maneja el Outlook que el usuario ya tiene abierto y firmado, asi
 * que evita el AADSTS65002 / consentimiento de admin del tenant. Windows-only.
 *
 * Los datos del usuario se pasan por VARIABLES DE ENTORNO (no interpolados en el
 * script) para evitar inyeccion de PowerShell.
 */
import { execSync } from 'node:child_process'

export const isWindows = process.platform === 'win32'

/** Ejecuta un script de PowerShell con datos de usuario por env vars. */
export function ps(script, env = {}, timeout = 60000) {
  if (!isWindows) {
    throw new Error('Este modo usa Outlook de escritorio por COM y requiere Windows con Outlook instalado y abierto.')
  }
  const fullScript = `$ProgressPreference = 'SilentlyContinue'\n$ErrorActionPreference = 'Stop'\n${script}`
  const encoded = Buffer.from(fullScript, 'utf16le').toString('base64')
  try {
    return execSync(`powershell -NonInteractive -EncodedCommand ${encoded}`, {
      encoding: 'utf8', timeout, env: { ...process.env, ...env }, stdio: ['pipe', 'pipe', 'pipe'],
    }).trim()
  } catch (err) {
    const raw = err.stderr ?? err.stdout ?? err.message ?? ''
    const xmlMatch = raw.match(/<S S="Error">([\s\S]*?)<\/S>/)
    const msg = xmlMatch
      ? xmlMatch[1].replace(/&#x([0-9A-F]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16)))
      : (raw.replace(/#< CLIXML[\s\S]*/, '').trim() || err.message)
    throw new Error(msg)
  }
}

/** Identidad del usuario segun el Outlook local (para distinguir "lo tuyo"). */
let cachedMe = null
export function getMeCom() {
  if (cachedMe) { return cachedMe }
  const script = `
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$u = $ns.CurrentUser
$name = $u.Name
$email = ""
try { $email = $u.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch {}
if (-not $email) { try { $email = $u.Address } catch {} }
[PSCustomObject]@{ name = $name; email = $email } | ConvertTo-Json
`
  try {
    const out = ps(script)
    const j = JSON.parse(out)
    cachedMe = { name: j.name || '', email: (j.email || '').toLowerCase() }
  } catch {
    cachedMe = { name: '', email: '' }
  }
  return cachedMe
}
