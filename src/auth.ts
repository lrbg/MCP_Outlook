import * as vscode from 'vscode'
import {
  requestDeviceCode, pollForToken, refreshAccessToken, authCodeFlow, TokenSet,
} from './graphAuth'

/**
 * Autenticacion a Microsoft 365 (Graph). NO usa el proveedor 'microsoft' de VS
 * Code porque su app (aebc6443-...) no esta preautorizada para pedir scopes de
 * correo/agenda/Teams y da AADSTS65002. En su lugar usamos un client_id publico
 * preconsentido (Microsoft Graph PowerShell por defecto) con:
 *   - login por NAVEGADOR (auth-code + PKCE): primario, pasa Conditional Access
 *     en equipos unidos a Entra (Edge presenta el claim de dispositivo).
 *   - device code: alterno, para cuando el navegador no es viable.
 * El refresh_token se guarda cifrado en SecretStorage; nada vive en el repo.
 * El usuario puede apuntar a una App Registration propia con m365.graph.clientId.
 */

/** Microsoft Graph PowerShell — client publico preconsentido en muchos tenants. */
export const GRAPH_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
export const GRAPH_RT_KEY = 'm365.graph.refreshToken'

const DEFAULT_SCOPES = [
  'User.Read', 'Mail.ReadWrite', 'Mail.Send', 'Calendars.ReadWrite', 'Chat.ReadWrite', 'offline_access',
]

export interface M365Session {
  accessToken: string
  /** Permisos realmente concedidos (claim scp del token). */
  scopes: string[]
  account: string
  expiresOn: number
}

function cfg() { return vscode.workspace.getConfiguration('m365') }
function clientId() { return (cfg().get<string>('graph.clientId') || '').trim() || GRAPH_CLIENT_ID }
function tenant() { return (cfg().get<string>('graph.tenantId') || '').trim() || 'organizations' }
function scopes(): string[] {
  const s = cfg().get<string[]>('graph.scopes')
  const list = (Array.isArray(s) && s.length) ? s.slice() : DEFAULT_SCOPES.slice()
  if (!list.includes('offline_access')) { list.push('offline_access') }
  return list
}

/** Decodifica el payload de un JWT (sin verificar; solo para leer claims). */
function decodeJwt(token: string): any {
  try {
    const part = token.split('.')[1] || ''
    const buf = Buffer.from(part.replace(/-/g, '+').replace(/_/g, '/'), 'base64')
    return JSON.parse(buf.toString('utf8'))
  } catch { return {} }
}

function toSession(tok: TokenSet): M365Session {
  const c = decodeJwt(tok.access_token)
  const scp = typeof c.scp === 'string' ? c.scp.split(' ').filter(Boolean) : []
  return {
    accessToken: tok.access_token,
    scopes: scp,
    account: c.upn || c.preferred_username || c.unique_name || c.email || '',
    expiresOn: Date.now() + (tok.expires_in ? tok.expires_in * 1000 : 3600 * 1000),
  }
}

/**
 * Login por NAVEGADOR (recomendado). Abre el navegador del sistema; al volver,
 * guarda el refresh_token y devuelve la sesion.
 */
export async function signIn(context: vscode.ExtensionContext): Promise<M365Session> {
  const tok = await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: 'Microsoft 365: completa el inicio de sesion en el navegador…' },
    () => authCodeFlow(clientId(), tenant(), scopes(), (url) => { vscode.env.openExternal(vscode.Uri.parse(url)) }),
  )
  await context.secrets.store(GRAPH_RT_KEY, tok.refresh_token || '')
  return toSession(tok)
}

/**
 * Login por DEVICE CODE (alterno). Muestra el codigo y abre microsoft.com/devicelogin.
 */
export async function signInDeviceCode(context: vscode.ExtensionContext): Promise<M365Session> {
  const dc = await requestDeviceCode(clientId(), tenant(), scopes())
  const abrir = 'Abrir y copiar codigo'
  const pick = await vscode.window.showInformationMessage(
    `Ve a ${dc.verification_uri} e ingresa el codigo: ${dc.user_code}`,
    { modal: false }, abrir,
  )
  if (pick === abrir) {
    await vscode.env.clipboard.writeText(dc.user_code)
    await vscode.env.openExternal(vscode.Uri.parse(dc.verification_uri))
  }
  const tok = await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: `Microsoft 365: esperando el login (codigo ${dc.user_code})…` },
    () => pollForToken(clientId(), tenant(), dc.device_code, dc.interval, dc.expires_in),
  )
  await context.secrets.store(GRAPH_RT_KEY, tok.refresh_token || '')
  return toSession(tok)
}

/**
 * Refresco silencioso: usa el refresh_token guardado. Devuelve undefined si no
 * hay sesion previa. Si el refresh falla (revocado/expirado), borra el token.
 */
export async function refreshSilent(context: vscode.ExtensionContext): Promise<M365Session | undefined> {
  const rt = (await context.secrets.get(GRAPH_RT_KEY)) || ''
  if (!rt) { return undefined }
  try {
    const tok = await refreshAccessToken(clientId(), tenant(), scopes(), rt)
    if (tok.refresh_token) { await context.secrets.store(GRAPH_RT_KEY, tok.refresh_token) }
    return toSession(tok)
  } catch {
    // Token revocado o invalido: limpiar para forzar un nuevo login.
    await context.secrets.delete(GRAPH_RT_KEY)
    return undefined
  }
}
