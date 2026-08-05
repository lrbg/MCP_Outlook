/**
 * Lee la config local que escribe la extension de VS Code
 * (globalStorage/m365-config.json), apuntada por la env M365_CONFIG_FILE.
 * Contiene el token de Graph y los permisos concedidos. Se relee en cada
 * llamada para tomar siempre el token mas fresco (opcion B: token en archivo).
 */
import { readFileSync, existsSync } from 'node:fs'

export function loadConfig() {
  const f = process.env.M365_CONFIG_FILE
  if (!f || !existsSync(f)) {
    return { accessToken: '', scopes: [], account: '', expiresOn: 0, graphBase: 'https://graph.microsoft.com/v1.0' }
  }
  try {
    const cfg = JSON.parse(readFileSync(f, 'utf8'))
    return {
      accessToken: cfg.accessToken || '',
      scopes: Array.isArray(cfg.scopes) ? cfg.scopes : [],
      account: cfg.account || '',
      expiresOn: cfg.expiresOn || 0,
      graphBase: cfg.graphBase || 'https://graph.microsoft.com/v1.0',
    }
  } catch {
    return { accessToken: '', scopes: [], account: '', expiresOn: 0, graphBase: 'https://graph.microsoft.com/v1.0' }
  }
}

/** Capacidades derivadas de los permisos concedidos. */
export function capabilities(scopes = []) {
  const has = (s) => scopes.includes(s)
  return {
    readMail: has('Mail.Read') || has('Mail.ReadWrite'),
    draftMail: has('Mail.ReadWrite'),
    sendMail: has('Mail.Send'),
    readCalendar: has('Calendars.Read') || has('Calendars.ReadWrite'),
    writeCalendar: has('Calendars.ReadWrite'),
    readChat: has('Chat.Read') || has('Chat.ReadWrite'),
    sendChat: has('Chat.ReadWrite'),
  }
}
