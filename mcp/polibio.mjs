/**
 * Conector con el Anotador de Reuniones de PolibioDesk (proyecto RAG
 * "Minutas de Reuniones"). Lee el texto crudo de minutas/transcripciones a
 * traves de una edge function de LECTURA en Polibio (verify_jwt=false,
 * autenticada con un token propio, mismo patron que ingest-minuta).
 *
 * La idea: entregarle al agente la minuta + transcripcion para que razone,
 * cruzando con correo/agenda/Teams (herramientas M365 del mismo plugin) y, en
 * planeacion, liste TUS pendientes (extraccion a cargo del agente).
 *
 * Config (escrita por la extension en m365-config.json, seccion `polibio`):
 *   functionUrl  https://<ref>.supabase.co/functions/v1/minutas-read
 *   anonKey      anon key de Polibio (header apikey del gateway)
 *   token        token de lectura (valida la edge function)
 */
import { z } from 'zod'
import { getMeCom } from './comShared.mjs'

const ok = (text) => ({ content: [{ type: 'text', text }] })
const bad = (text) => ({ content: [{ type: 'text', text }], isError: true })
const json = (obj) => ok(JSON.stringify(obj, null, 2))

/** Llama la edge function de lectura de minutas. */
async function callMinutas(polibio, action, extra = {}) {
  const res = await fetch(polibio.functionUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      apikey: polibio.anonKey,
      Authorization: `Bearer ${polibio.anonKey}`,
    },
    body: JSON.stringify({ token: polibio.token, action, ...extra }),
  })
  const text = await res.text()
  let data = null
  try { data = text ? JSON.parse(text) : null } catch { /* respuesta no-JSON */ }
  if (!res.ok) {
    const msg = data?.error || text || `HTTP ${res.status}`
    throw new Error(`Polibio minutas-read fallo (${res.status}): ${msg}`)
  }
  return data
}

export function registerPolibioTools(server, polibio) {
  if (!polibio || !polibio.functionUrl || !polibio.token || !polibio.anonKey) {
    // Sin configurar: exponemos una sola tool que explica como activarlo.
    server.tool(
      'polibio_status',
      'Estado del conector con el Anotador de PolibioDesk (minutas).',
      {},
      async () => ok(
        'Conector de Polibio no configurado. En VS Code abre la configuracion y define ' +
        'm365.polibio.supabaseUrl y m365.polibio.anonKey, y guarda el token con ' +
        '"M365: Guardar token de Polibio (minutas)".',
      ),
    )
    return
  }

  server.tool(
    'polibio_minutas_list',
    'Lista las minutas/actas recientes del Anotador de PolibioDesk (proyecto Minutas de Reuniones).',
    {
      limit: z.number().int().min(1).max(100).optional().describe('Cuantas traer (default 25)'),
      search: z.string().optional().describe('Filtra por texto en el titulo/nombre'),
    },
    async ({ limit, search }) => {
      try {
        const data = await callMinutas(polibio, 'list', { limit: limit || 25, search: search || '' })
        return json({ count: (data?.minutas || []).length, minutas: data?.minutas || [] })
      } catch (e) { return bad(e.message) }
    },
  )

  server.tool(
    'polibio_minuta_get',
    'Devuelve el texto crudo (minuta + transcripcion) de una minuta por su id, para que el agente lo analice.',
    { id: z.string().describe('id de la minuta (de polibio_minutas_list)') },
    async ({ id }) => {
      try {
        const data = await callMinutas(polibio, 'get', { id })
        return json({ id: data?.id, name: data?.name, created_at: data?.created_at, text: data?.text || '' })
      } catch (e) { return bad(e.message) }
    },
  )

  // ── Skill: contexto de reunion (minuta cruda para el agente) ──
  server.tool(
    'meeting_context',
    'Devuelve el texto de la minuta para que el agente lo analice y lo cruce con tu correo/agenda (usando outlook_list_calendar y outlook_list_emails).',
    { minutaId: z.string().describe('id de la minuta') },
    async ({ minutaId }) => {
      try {
        const minuta = await callMinutas(polibio, 'get', { id: minutaId })
        return json({
          minuta: { id: minuta?.id, name: minuta?.name, created_at: minuta?.created_at, text: minuta?.text || '' },
          hint: 'Para dar contexto, cruza esta minuta con outlook_list_calendar (reuniones de ese dia) y ' +
            'outlook_list_emails (correos relacionados por asunto/remitente).',
        })
      } catch (e) { return bad(e.message) }
    },
  )

  // ── Skill: mis pendientes (el agente extrae; la tool arma el material) ──
  server.tool(
    'my_action_items',
    'Reune el texto de una o varias minutas junto con tu identidad (de Outlook), para que el agente extraiga y liste TUS acuerdos/pendientes (los tuyos, no los de otros) y proponga seguimiento con correo.',
    { minutaIds: z.array(z.string()).min(1).describe('ids de las minutas a revisar') },
    async ({ minutaIds }) => {
      try {
        const me = getMeCom()
        const minutas = []
        for (const id of minutaIds.slice(0, 10)) {
          const m = await callMinutas(polibio, 'get', { id }).catch(() => null)
          if (m) { minutas.push({ id: m.id, name: m.name, created_at: m.created_at, text: m.text || '' }) }
        }
        return json({
          you: { name: me.name, email: me.email },
          minutas,
          instruction:
            `Extrae de estas minutas SOLO los acuerdos y pendientes asignados a ${me.name || 'el usuario'} ` +
            `(${me.email}). Ignora los de otras personas. Para cada pendiente indica: descripcion, fecha limite si ` +
            `aparece, y de que reunion viene. Si algun pendiente se relaciona con un correo, sugiere el seguimiento ` +
            `con outlook_send_email o outlook_reply_email.`,
        })
      } catch (e) { return bad(e.message) }
    },
  )
}
