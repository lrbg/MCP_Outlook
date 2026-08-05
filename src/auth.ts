import * as vscode from 'vscode'

/**
 * Autenticacion contra Microsoft 365 reutilizando el proveedor 'microsoft'
 * integrado de VS Code (la misma cuenta de organizacion con la que el usuario
 * ya esta firmado, la que habilita su Copilot de equipo). No se registra una
 * app propia en Azure: el proveedor de VS Code aporta su client_id.
 *
 * Riesgo conocido: algunos permisos (Mail.Send, Chat.ReadWrite) pueden requerir
 * consentimiento del administrador del tenant. Por eso pedimos los permisos en
 * una "escalera": si el conjunto completo falla, probamos conjuntos mas chicos
 * hasta obtener al menos correo (lectura) y agenda. Cada herramienta del MCP se
 * habilita segun los permisos realmente concedidos (ver availableScopes).
 */

const GRAPH = 'https://graph.microsoft.com/'
const PROVIDER = 'microsoft'
const STATE_KEY = 'm365.grantedScopes'

/**
 * Escalera de conjuntos de permisos (nombres cortos de Graph), del mas completo
 * al minimo. Se prueba en orden hasta que uno se conceda.
 */
const SCOPE_LADDER: string[][] = [
  ['User.Read', 'Mail.ReadWrite', 'Mail.Send', 'Calendars.ReadWrite', 'Chat.ReadWrite'],
  ['User.Read', 'Mail.ReadWrite', 'Calendars.ReadWrite', 'Chat.ReadWrite'],
  ['User.Read', 'Mail.ReadWrite', 'Calendars.ReadWrite'],
  ['User.Read', 'Mail.Read', 'Calendars.Read'],
  ['User.Read'],
]

export interface M365Session {
  accessToken: string
  /** Nombres cortos de los permisos concedidos, ej. ['Mail.ReadWrite','Calendars.ReadWrite']. */
  scopes: string[]
  account: string
  /** Epoch ms de expiracion estimada del token (~55 min si el proveedor no la da). */
  expiresOn: number
}

/** Lee m365.clientId / m365.tenantId de la configuracion (opcionales, viven fuera del repo). */
function readOverrides() {
  const c = vscode.workspace.getConfiguration('m365')
  return {
    clientId: (c.get<string>('clientId', '') || '').trim(),
    tenantId: (c.get<string>('tenantId', '') || '').trim(),
  }
}

/**
 * Convierte permisos cortos en las cadenas que espera el proveedor 'microsoft':
 * URIs completos de Graph + offline_access, y opcionalmente los overrides
 * VSCODE_CLIENT_ID / VSCODE_TENANT para apuntar a una app registrada propia.
 */
function buildRequestScopes(shortScopes: string[]): string[] {
  const { clientId, tenantId } = readOverrides()
  const req: string[] = []
  if (clientId) { req.push(`VSCODE_CLIENT_ID:${clientId}`) }
  if (tenantId) { req.push(`VSCODE_TENANT:${tenantId}`) }
  for (const s of shortScopes) { req.push(GRAPH + s) }
  req.push('offline_access')
  return req
}

/** Reduce las cadenas de permisos concedidos a nombres cortos de Graph. */
function normalizeGranted(scopes: readonly string[] | undefined): string[] {
  const out: string[] = []
  for (const raw of scopes || []) {
    if (!raw) { continue }
    if (raw.startsWith('VSCODE_')) { continue }
    let s = raw
    if (s.startsWith(GRAPH)) { s = s.slice(GRAPH.length) }
    if (s === 'offline_access') { continue }
    if (!out.includes(s)) { out.push(s) }
  }
  return out
}

function toSession(session: vscode.AuthenticationSession): M365Session {
  return {
    accessToken: session.accessToken,
    scopes: normalizeGranted(session.scopes),
    account: session.account.label,
    // El proveedor no expone la expiracion; asumimos ~55 min (tokens de Graph duran ~60-75).
    expiresOn: Date.now() + 55 * 60 * 1000,
  }
}

/**
 * Inicio de sesion interactivo: recorre la escalera con createIfNone hasta que
 * un conjunto de permisos se conceda. Guarda el conjunto concedido para futuros
 * refrescos silenciosos.
 */
export async function signIn(context: vscode.ExtensionContext): Promise<M365Session> {
  let lastError: unknown
  for (const shortScopes of SCOPE_LADDER) {
    try {
      const session = await vscode.authentication.getSession(
        PROVIDER,
        buildRequestScopes(shortScopes),
        { createIfNone: true },
      )
      if (session) {
        await context.globalState.update(STATE_KEY, shortScopes)
        return toSession(session)
      }
    } catch (e) {
      lastError = e
      // Este conjunto no se pudo conceder (probablemente falta consentimiento de
      // admin para algun permiso). Probamos el siguiente, mas acotado.
    }
  }
  throw new Error(
    'No se pudo iniciar sesion con Microsoft. ' +
    (lastError instanceof Error ? lastError.message : String(lastError ?? '')),
  )
}

/**
 * Refresco silencioso: reusa el conjunto de permisos ya concedido, sin abrir
 * dialogos. Devuelve undefined si no hay sesion previa (nunca fuerza login).
 */
export async function refreshSilent(context: vscode.ExtensionContext): Promise<M365Session | undefined> {
  const saved = context.globalState.get<string[]>(STATE_KEY)
  const ladder = saved && saved.length ? [saved, ...SCOPE_LADDER] : SCOPE_LADDER
  for (const shortScopes of ladder) {
    try {
      const session = await vscode.authentication.getSession(
        PROVIDER,
        buildRequestScopes(shortScopes),
        { createIfNone: false, silent: true },
      )
      if (session) {
        await context.globalState.update(STATE_KEY, normalizeGranted(session.scopes))
        return toSession(session)
      }
    } catch {
      // Sin sesion para este conjunto; probamos otro.
    }
  }
  return undefined
}
