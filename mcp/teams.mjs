/**
 * Herramientas de chats de Teams (Microsoft Graph).
 *   Lectura: teams_chats_list, teams_messages_read
 *   Escritura: teams_send_message (confirm)
 */
import { z } from 'zod'
import { graphGetAll, graphPost } from './graph.mjs'
import { buildPreview, isConfirmed } from './lib/preview.mjs'

const ok = (text) => ({ content: [{ type: 'text', text }] })
const bad = (text) => ({ content: [{ type: 'text', text }], isError: true })
const json = (obj) => ok(JSON.stringify(obj, null, 2))

function slimChat(c) {
  return {
    id: c.id,
    topic: c.topic || '',
    chatType: c.chatType || '',
    lastUpdated: c.lastUpdatedDateTime || '',
    members: (c.members || []).map(m => m.displayName || m.email).filter(Boolean),
  }
}

function slimMessage(m) {
  return {
    id: m.id,
    from: m.from?.user?.displayName || '',
    created: m.createdDateTime || '',
    text: (m.body?.content || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(),
    mentions: (m.mentions || []).map(x => x.mentionText).filter(Boolean),
  }
}

export function registerTeamsTools(server, caps) {
  if (caps.readChat) {
    server.tool(
      'teams_chats_list',
      'Lista tus chats de Teams mas recientes (1:1 y grupales), con sus miembros.',
      { top: z.number().int().min(1).max(50).optional().describe('Cuantos traer (default 20)') },
      async ({ top }) => {
        try {
          const max = top || 20
          const path = `/me/chats?$expand=members&$top=${max}&$orderby=lastUpdatedDateTime desc`
          const items = await graphGetAll(path, max)
          return json({ count: items.length, chats: items.map(slimChat) })
        } catch (e) { return bad(e.message) }
      },
    )

    server.tool(
      'teams_messages_read',
      'Lee los mensajes recientes de un chat de Teams por su id (obtenido con teams_chats_list).',
      {
        chatId: z.string().describe('id del chat'),
        top: z.number().int().min(1).max(50).optional().describe('Cuantos mensajes (default 20)'),
      },
      async ({ chatId, top }) => {
        try {
          const max = top || 20
          const path = `/me/chats/${encodeURIComponent(chatId)}/messages?$top=${max}`
          const items = await graphGetAll(path, max)
          return json({ count: items.length, messages: items.map(slimMessage) })
        } catch (e) { return bad(e.message) }
      },
    )
  }

  if (caps.sendChat) {
    server.tool(
      'teams_send_message',
      'Envia un mensaje a un chat de Teams. Requiere confirm:true; sin el, devuelve un preview.',
      {
        chatId: z.string().describe('id del chat destino'),
        text: z.string().describe('Texto del mensaje'),
        confirm: z.boolean().optional().describe('Debe ser true para enviar de verdad'),
      },
      async (args) => {
        if (!isConfirmed(args)) {
          return ok(buildPreview('Enviar mensaje de Teams', {
            'Chat id': args.chatId,
            'Mensaje': args.text,
          }))
        }
        try {
          await graphPost(`/me/chats/${encodeURIComponent(args.chatId)}/messages`, {
            body: { contentType: 'text', content: args.text },
          })
          return ok('Mensaje enviado a Teams.')
        } catch (e) { return bad(e.message) }
      },
    )
  }
}
