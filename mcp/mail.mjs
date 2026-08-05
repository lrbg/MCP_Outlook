/**
 * Herramientas de correo (Outlook / Microsoft Graph).
 *   Lectura: mail_list, mail_read
 *   Escritura: mail_draft (borrador), mail_send, mail_reply (con confirm)
 */
import { z } from 'zod'
import { graphGet, graphGetAll, graphPost } from './graph.mjs'
import { capabilities } from './config.mjs'
import { buildMailFilter } from './lib/odata.mjs'
import { buildPreview, isConfirmed } from './lib/preview.mjs'

const ok = (text) => ({ content: [{ type: 'text', text }] })
const bad = (text) => ({ content: [{ type: 'text', text }], isError: true })
const json = (obj) => ok(JSON.stringify(obj, null, 2))

/** Normaliza un mensaje de Graph a una forma compacta. */
function slim(m) {
  return {
    id: m.id,
    from: m.from?.emailAddress?.address || '',
    fromName: m.from?.emailAddress?.name || '',
    to: (m.toRecipients || []).map(r => r.emailAddress?.address).filter(Boolean),
    cc: (m.ccRecipients || []).map(r => r.emailAddress?.address).filter(Boolean),
    subject: m.subject || '',
    receivedDateTime: m.receivedDateTime || '',
    isRead: m.isRead,
    importance: m.importance,
    bodyPreview: m.bodyPreview || '',
    webLink: m.webLink || '',
  }
}

function recipients(list) {
  return (list || []).map(addr => ({ emailAddress: { address: addr } }))
}

export function registerMailTools(server, caps) {
  if (caps.readMail) {
    server.tool(
      'mail_list',
      'Lista/busca correos de tu bandeja. Filtra por remitente, asunto, no-leidos y rango de fechas (ISO 8601). Solo lectura.',
      {
        from: z.string().optional().describe('Filtra por direccion del remitente (contiene)'),
        subject: z.string().optional().describe('Filtra por asunto (contiene)'),
        unreadOnly: z.boolean().optional().describe('Solo correos no leidos'),
        since: z.string().optional().describe('Recibidos desde esta fecha/hora ISO 8601'),
        until: z.string().optional().describe('Recibidos hasta esta fecha/hora ISO 8601'),
        top: z.number().int().min(1).max(50).optional().describe('Cuantos traer (default 15)'),
      },
      async (args) => {
        try {
          const filter = buildMailFilter(args)
          const top = args.top || 15
          const select = 'id,from,toRecipients,ccRecipients,subject,receivedDateTime,isRead,importance,bodyPreview,webLink'
          let path = `/me/messages?$select=${select}&$top=${top}&$orderby=receivedDateTime desc`
          if (filter) { path += `&$filter=${encodeURIComponent(filter)}` }
          const items = await graphGetAll(path, top)
          return json({ count: items.length, messages: items.map(slim) })
        } catch (e) { return bad(e.message) }
      },
    )

    server.tool(
      'mail_read',
      'Lee el cuerpo completo de un correo por su id (obtenido con mail_list).',
      { id: z.string().describe('id del mensaje') },
      async ({ id }) => {
        try {
          const m = await graphGet(`/me/messages/${encodeURIComponent(id)}`)
          return json({
            ...slim(m),
            body: m.body?.content || '',
            bodyType: m.body?.contentType || 'text',
          })
        } catch (e) { return bad(e.message) }
      },
    )
  }

  if (caps.draftMail) {
    server.tool(
      'mail_draft',
      'Crea un BORRADOR de correo (no lo envia). Devuelve el id del borrador.',
      {
        to: z.array(z.string()).describe('Destinatarios (direcciones)'),
        subject: z.string().describe('Asunto'),
        body: z.string().describe('Cuerpo del correo (texto)'),
        cc: z.array(z.string()).optional().describe('Copias (CC)'),
      },
      async ({ to, subject, body, cc }) => {
        try {
          const draft = await graphPost('/me/messages', {
            subject,
            body: { contentType: 'Text', content: body },
            toRecipients: recipients(to),
            ccRecipients: recipients(cc),
          })
          return json({ created: true, id: draft.id, webLink: draft.webLink || '' })
        } catch (e) { return bad(e.message) }
      },
    )
  }

  if (caps.sendMail) {
    server.tool(
      'mail_send',
      'Envia un correo nuevo. Requiere confirm:true; sin el, devuelve un preview y no envia nada.',
      {
        to: z.array(z.string()).describe('Destinatarios (direcciones)'),
        subject: z.string().describe('Asunto'),
        body: z.string().describe('Cuerpo del correo (texto)'),
        cc: z.array(z.string()).optional().describe('Copias (CC)'),
        confirm: z.boolean().optional().describe('Debe ser true para enviar de verdad'),
      },
      async (args) => {
        if (!isConfirmed(args)) {
          return ok(buildPreview('Enviar correo', {
            'Para': (args.to || []).join(', '),
            'CC': (args.cc || []).join(', '),
            'Asunto': args.subject,
            'Cuerpo': args.body,
          }))
        }
        try {
          await graphPost('/me/sendMail', {
            message: {
              subject: args.subject,
              body: { contentType: 'Text', content: args.body },
              toRecipients: recipients(args.to),
              ccRecipients: recipients(args.cc),
            },
            saveToSentItems: true,
          })
          return ok(`Correo enviado a ${(args.to || []).join(', ')}.`)
        } catch (e) { return bad(e.message) }
      },
    )

    server.tool(
      'mail_reply',
      'Responde a un correo existente. Requiere confirm:true; sin el, devuelve un preview.',
      {
        id: z.string().describe('id del mensaje a responder'),
        body: z.string().describe('Texto de la respuesta'),
        replyAll: z.boolean().optional().describe('Responder a todos (default: solo al remitente)'),
        confirm: z.boolean().optional().describe('Debe ser true para enviar de verdad'),
      },
      async (args) => {
        if (!isConfirmed(args)) {
          return ok(buildPreview(args.replyAll ? 'Responder a todos' : 'Responder al remitente', {
            'Mensaje id': args.id,
            'Respuesta': args.body,
          }))
        }
        try {
          const action = args.replyAll ? 'replyAll' : 'reply'
          await graphPost(`/me/messages/${encodeURIComponent(args.id)}/${action}`, {
            comment: args.body,
          })
          return ok('Respuesta enviada.')
        } catch (e) { return bad(e.message) }
      },
    )
  }
}
