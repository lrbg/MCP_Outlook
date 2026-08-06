/**
 * Autenticacion a Microsoft Graph por DEVICE CODE FLOW y por AUTH-CODE + PKCE
 * con el navegador del sistema.
 *
 * El proveedor de Microsoft integrado de VS Code no puede pedir scopes de
 * correo/agenda/Teams (Mail/Calendars/Chat): da AADSTS65002 (su app no esta
 * preautorizada contra Graph). Aqui se usa un client_id publico preconsentido
 * (por defecto el de Microsoft Graph PowerShell) para obtener un token de Graph
 * con esos permisos. Solo HTTP; el usuario autentica en microsoft.com/devicelogin
 * o en el navegador. No se guarda ninguna contrasena.
 */
import * as http from 'node:http'
import * as crypto from 'node:crypto'

const authority = (tenant: string) => `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0`

const b64url = (buf: Buffer) => buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')

async function form(url: string, body: Record<string, string>): Promise<{ status: number; json: any }> {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(body).toString(),
  })
  let json: any = {}
  try { json = await res.json() } catch { /* respuesta no JSON */ }
  return { status: res.status, json }
}

export interface DeviceCode {
  user_code: string
  device_code: string
  verification_uri: string
  interval: number
  expires_in: number
  message: string
}

export interface TokenSet {
  access_token: string
  refresh_token?: string
  expires_in: number
}

/** Solicita un codigo de dispositivo. El usuario lo ingresa en verification_uri. */
export async function requestDeviceCode(clientId: string, tenant: string, scopes: string[]): Promise<DeviceCode> {
  const { status, json } = await form(`${authority(tenant)}/devicecode`, { client_id: clientId, scope: scopes.join(' ') })
  if (status !== 200 || !json.device_code) {
    throw new Error(json.error_description || json.error || `devicecode HTTP ${status}`)
  }
  return json as DeviceCode
}

/** Hace polling hasta que el usuario completa el login o expira el codigo. */
export async function pollForToken(clientId: string, tenant: string, deviceCode: string, intervalSec: number, expiresInSec: number): Promise<TokenSet> {
  const deadline = Date.now() + Math.max(expiresInSec, 60) * 1000
  let interval = Math.max(intervalSec || 5, 1) * 1000
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, interval))
    const { status, json } = await form(`${authority(tenant)}/token`, {
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: clientId,
      device_code: deviceCode,
    })
    if (status === 200 && json.access_token) { return json as TokenSet }
    const err = json.error
    if (err === 'authorization_pending') { continue }
    if (err === 'slow_down') { interval += 5000; continue }
    throw new Error(json.error_description || err || `token HTTP ${status}`)
  }
  throw new Error('El codigo de dispositivo expiro antes de completar el inicio de sesion.')
}

/** Renueva el access_token con el refresh_token guardado. */
export async function refreshAccessToken(clientId: string, tenant: string, scopes: string[], refreshToken: string): Promise<TokenSet> {
  const { status, json } = await form(`${authority(tenant)}/token`, {
    grant_type: 'refresh_token',
    client_id: clientId,
    scope: scopes.join(' '),
    refresh_token: refreshToken,
  })
  if (status !== 200 || !json.access_token) {
    throw new Error(json.error_description || json.error || `refresh HTTP ${status}`)
  }
  return json as TokenSet
}

/**
 * Auth-code + PKCE con el NAVEGADOR DEL SISTEMA y redirect a loopback.
 * A diferencia del device code, el navegador (Edge en un equipo unido a Entra)
 * presenta el PRT/claim de dispositivo administrado, asi que satisface las
 * politicas de Conditional Access que exigen dispositivo compliant.
 * Levanta un servidor local efimero para recibir el redirect.
 */
export async function authCodeFlow(clientId: string, tenant: string, scopes: string[], openBrowser: (url: string) => void): Promise<TokenSet> {
  const verifier = b64url(crypto.randomBytes(32))
  const challenge = b64url(crypto.createHash('sha256').update(verifier).digest())
  const state = b64url(crypto.randomBytes(16))

  return await new Promise<TokenSet>((resolve, reject) => {
    let redirectUri = ''
    let timer: NodeJS.Timeout

    const finish = (fn: () => void) => { clearTimeout(timer); try { server.close() } catch { /* ya cerrado */ } fn() }

    const server = http.createServer(async (req, res) => {
      const url = new URL(req.url || '/', 'http://localhost')
      if (!url.searchParams.has('code') && !url.searchParams.has('error')) { res.writeHead(404); res.end(); return }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' })
      res.end('<html><body style="font-family:sans-serif;padding:2rem"><h3>Microsoft 365 · MCP</h3><p>Listo. Ya puedes cerrar esta pestana y volver a VS Code.</p></body></html>')

      const err = url.searchParams.get('error')
      if (err) { finish(() => reject(new Error(url.searchParams.get('error_description') || err))); return }
      if (url.searchParams.get('state') !== state) { finish(() => reject(new Error('state no coincide (posible CSRF)'))); return }
      const code = url.searchParams.get('code')
      if (!code) { finish(() => reject(new Error('sin authorization code'))); return }
      try {
        const { status, json } = await form(`${authority(tenant)}/token`, {
          grant_type: 'authorization_code', client_id: clientId, code, redirect_uri: redirectUri, code_verifier: verifier, scope: scopes.join(' '),
        })
        if (status !== 200 || !json.access_token) { finish(() => reject(new Error(json.error_description || json.error || `token HTTP ${status}`))); return }
        finish(() => resolve(json as TokenSet))
      } catch (e) { finish(() => reject(e as Error)) }
    })

    server.on('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const port = (server.address() as any).port
      redirectUri = `http://localhost:${port}`
      const p = new URLSearchParams({
        client_id: clientId,
        response_type: 'code',
        redirect_uri: redirectUri,
        response_mode: 'query',
        scope: scopes.join(' '),
        state,
        code_challenge: challenge,
        code_challenge_method: 'S256',
        prompt: 'select_account',
      })
      openBrowser(`${authority(tenant)}/authorize?${p.toString()}`)
    })
    timer = setTimeout(() => finish(() => reject(new Error('Tiempo agotado esperando el login en el navegador.'))), 5 * 60 * 1000)
  })
}
