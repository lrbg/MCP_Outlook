/**
 * Cliente minimo de Microsoft Graph para el servidor MCP.
 * - Lee el token del config local en cada llamada (siempre el mas fresco).
 * - Maneja 401 (token expirado/insuficiente) con un mensaje accionable.
 * - Respeta 429 (throttling) con Retry-After y backoff.
 * - Sigue la paginacion @odata.nextLink hasta un tope.
 */
import { loadConfig } from './config.mjs'

/** Error con mensaje ya legible para el usuario final. */
export class GraphError extends Error {}

function authHeader() {
  const { accessToken, expiresOn } = loadConfig()
  if (!accessToken) {
    throw new GraphError(
      'No hay sesion de Microsoft. En VS Code ejecuta "M365: Iniciar sesion (Microsoft)" y luego "M365: Registrar servidor MCP".',
    )
  }
  if (expiresOn && Date.now() > expiresOn) {
    throw new GraphError(
      'El token de Microsoft expiro. En VS Code ejecuta "M365: Refrescar sesion" y reintenta.',
    )
  }
  return accessToken
}

function base() {
  return loadConfig().graphBase || 'https://graph.microsoft.com/v1.0'
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

async function request(method, path, body, { retry = 2 } = {}) {
  const url = path.startsWith('http') ? path : base() + path
  let attempt = 0
  while (true) {
    const token = authHeader()
    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(body ? { 'Content-Type': 'application/json' } : {}),
      },
      body: body ? JSON.stringify(body) : undefined,
    })

    if (res.status === 429 && attempt < retry) {
      const wait = Number(res.headers.get('Retry-After') || 2) * 1000
      await sleep(wait)
      attempt++
      continue
    }
    if (res.status === 401) {
      throw new GraphError(
        'Microsoft Graph rechazo el token (401). Puede haber expirado o faltar el permiso para esta accion. ' +
        'En VS Code ejecuta "M365: Refrescar sesion"; si persiste, el permiso quiza necesite consentimiento del administrador.',
      )
    }
    if (res.status === 403) {
      throw new GraphError(
        'Microsoft Graph nego el permiso (403). Esta accion requiere un permiso que tu organizacion no ha concedido.',
      )
    }
    if (res.status === 204) { return null }
    const text = await res.text()
    let json = null
    try { json = text ? JSON.parse(text) : null } catch { /* respuesta no-JSON */ }
    if (!res.ok) {
      const msg = json?.error?.message || text || `HTTP ${res.status}`
      throw new GraphError(`Graph ${method} ${path} fallo (${res.status}): ${msg}`)
    }
    return json
  }
}

export const graphGet = (path) => request('GET', path)
export const graphPost = (path, body) => request('POST', path, body)
export const graphPatch = (path, body) => request('PATCH', path, body)
export const graphDelete = (path) => request('DELETE', path)

/** GET con paginacion: junta value[] siguiendo @odata.nextLink hasta maxItems. */
export async function graphGetAll(path, maxItems = 100) {
  const out = []
  let next = path
  while (next && out.length < maxItems) {
    const page = await request('GET', next)
    const items = page?.value || []
    out.push(...items)
    next = page?.['@odata.nextLink'] || null
  }
  return out.slice(0, maxItems)
}
