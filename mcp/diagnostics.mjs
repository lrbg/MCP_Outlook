/**
 * Herramientas de diagnostico, siempre disponibles (no dependen de permisos).
 * Sirven para que el agente verifique la conexion y sepa que puede hacer.
 */
import { loadConfig, capabilities } from './config.mjs'
import { getMe } from './me.mjs'

const ok = (text) => ({ content: [{ type: 'text', text }] })
const json = (obj) => ok(JSON.stringify(obj, null, 2))

export function registerDiagnosticsTools(server, caps) {
  server.tool(
    'm365_status',
    'Estado de la conexion a Microsoft 365: si hay token, que permisos se concedieron y que herramientas quedan habilitadas.',
    {},
    async () => {
      const cfg = loadConfig()
      const hasToken = !!cfg.accessToken
      const expired = cfg.expiresOn && Date.now() > cfg.expiresOn
      let identity = null
      if (hasToken && !expired) {
        try { identity = await getMe() } catch (e) { identity = { error: e.message } }
      }
      return json({
        connected: hasToken && !expired,
        account: cfg.account || null,
        tokenExpired: !!expired,
        grantedScopes: cfg.scopes,
        capabilities: capabilities(cfg.scopes),
        identity,
        hint: hasToken
          ? undefined
          : 'Sin token. En VS Code ejecuta "M365: Iniciar sesion (Microsoft)" y luego "M365: Registrar servidor MCP".',
      })
    },
  )
}
