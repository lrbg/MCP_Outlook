/**
 * Identidad del usuario (/me). Se usa para distinguir TUS pendientes/acuerdos
 * de los de otros en las skills. Se cachea por proceso.
 */
import { graphGet } from './graph.mjs'

let cached = null

export async function getMe() {
  if (cached) { return cached }
  const me = await graphGet('/me')
  cached = {
    id: me.id,
    name: me.displayName || '',
    email: (me.mail || me.userPrincipalName || '').toLowerCase(),
  }
  return cached
}

/** Solo para pruebas: limpia el cache. */
export function _resetMeCache() { cached = null }
