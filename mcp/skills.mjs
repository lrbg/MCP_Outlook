/**
 * Skills = orquestaciones de alto nivel sobre las herramientas base.
 * Todas usan /me para separar TUS pendientes de los de otros.
 *   daily_briefing      (lectura)
 *   inbox_triage        (lectura + borradores opcionales)
 *   schedule_from_email (escritura con confirm)
 *   teams_pending       (lectura)
 */
import { z } from 'zod'
import { graphGet, graphGetAll, graphPost } from './graph.mjs'
import { getMe } from './me.mjs'
import { dayRange } from './lib/odata.mjs'
import { classifyEmail, sortByPriority } from './lib/triage.mjs'
import { freeSlots, eventToBusy, findConflicts } from './lib/freebusy.mjs'
import { buildPreview, isConfirmed } from './lib/preview.mjs'

const ok = (text) => ({ content: [{ type: 'text', text }] })
const bad = (text) => ({ content: [{ type: 'text', text }], isError: true })
const json = (obj) => ok(JSON.stringify(obj, null, 2))

function slimMail(m) {
  return {
    id: m.id,
    from: m.from?.emailAddress?.address || '',
    subject: m.subject || '',
    receivedDateTime: m.receivedDateTime || '',
    isRead: m.isRead,
    importance: m.importance,
    bodyPreview: m.bodyPreview || '',
    toRecipients: (m.toRecipients || []).map(r => ({ address: r.emailAddress?.address })),
    ccRecipients: (m.ccRecipients || []).map(r => ({ address: r.emailAddress?.address })),
  }
}

export function registerSkillTools(server, caps) {
  // ── Briefing del dia ────────────────────────────────────────────
  if (caps.readMail || caps.readCalendar) {
    server.tool(
      'daily_briefing',
      'Resumen del dia: correos importantes sin responder + reuniones de hoy + conflictos de horario + huecos libres. Solo lectura.',
      { date: z.string().optional().describe('Fecha base ISO (default hoy)') },
      async ({ date }) => {
        try {
          const me = await getMe().catch(() => ({ email: '' }))
          const baseDate = date || new Date().toISOString()
          const { start, end } = dayRange(baseDate, 0)
          const result = { date: start.slice(0, 10), you: me.email || undefined }

          if (caps.readMail) {
            const select = 'id,from,toRecipients,ccRecipients,subject,receivedDateTime,isRead,importance,bodyPreview'
            const path = `/me/messages?$select=${select}&$top=25&$orderby=receivedDateTime desc&$filter=isRead eq false`
            const unread = (await graphGetAll(path, 25)).map(slimMail)
            const classified = unread.map(m => ({ ...m, ...classifyEmail(m, me.email) }))
            const important = sortByPriority(classified.filter(c => c.priority !== 'baja' || c.needsMyAction))
            result.importantUnread = important.slice(0, 10).map(m => ({
              from: m.from, subject: m.subject, priority: m.priority,
              needsMyAction: m.needsMyAction, receivedDateTime: m.receivedDateTime,
            }))
            result.totalUnread = unread.length
          }

          if (caps.readCalendar) {
            const evPath = `/me/calendarView?startDateTime=${encodeURIComponent(start)}` +
              `&endDateTime=${encodeURIComponent(end)}&$select=id,subject,start,end,attendees,isOnlineMeeting&$top=50&$orderby=start/dateTime`
            const events = await graphGetAll(evPath, 50)
            result.meetings = events.map(e => ({
              subject: e.subject, start: e.start?.dateTime, end: e.end?.dateTime,
              attendees: (e.attendees || []).length, isOnline: !!e.isOnlineMeeting,
            }))
            result.conflicts = findConflicts(events).map(([a, b]) => [a.subject, b.subject])
            const busy = events.map(eventToBusy)
            const workStart = Date.parse(start) + 9 * 3600 * 1000  // 9:00
            const workEnd = Date.parse(start) + 18 * 3600 * 1000   // 18:00
            const slots = freeSlots(workStart, workEnd, busy, 30)
            result.freeSlots = slots.map(s => ({
              start: new Date(s.start).toISOString(), end: new Date(s.end).toISOString(),
            }))
          }

          return json(result)
        } catch (e) { return bad(e.message) }
      },
    )
  }

  // ── Triage de bandeja ───────────────────────────────────────────
  if (caps.readMail) {
    server.tool(
      'inbox_triage',
      'Clasifica tus correos no leidos por prioridad y marca los que piden accion tuya. Opcionalmente crea borradores de respuesta (no envia).',
      {
        top: z.number().int().min(1).max(50).optional().describe('Cuantos no-leidos revisar (default 25)'),
        draftReplies: z.boolean().optional().describe('Crear borradores para los que piden accion (requiere permiso de escritura)'),
      },
      async ({ top, draftReplies }) => {
        try {
          const me = await getMe().catch(() => ({ email: '' }))
          const max = top || 25
          const select = 'id,from,toRecipients,ccRecipients,subject,receivedDateTime,isRead,importance,bodyPreview'
          const path = `/me/messages?$select=${select}&$top=${max}&$orderby=receivedDateTime desc&$filter=isRead eq false`
          const unread = (await graphGetAll(path, max)).map(slimMail)
          const classified = sortByPriority(unread.map(m => ({ ...m, ...classifyEmail(m, me.email) })))

          const out = { count: classified.length, items: classified.map(m => ({
            id: m.id, from: m.from, subject: m.subject, priority: m.priority,
            needsMyAction: m.needsMyAction, reasons: m.reasons,
          })) }

          if (draftReplies && caps.draftMail) {
            const drafts = []
            for (const m of classified.filter(x => x.needsMyAction).slice(0, 5)) {
              const draft = await graphPost(`/me/messages/${encodeURIComponent(m.id)}/createReply`, {})
              drafts.push({ forSubject: m.subject, draftId: draft.id })
            }
            out.draftsCreated = drafts
            out.note = 'Borradores creados vacios listos para editar; ninguno fue enviado.'
          } else if (draftReplies && !caps.draftMail) {
            out.note = 'No se crearon borradores: falta el permiso Mail.ReadWrite.'
          }
          return json(out)
        } catch (e) { return bad(e.message) }
      },
    )
  }

  // ── Agendar desde correo ────────────────────────────────────────
  if (caps.readMail && caps.writeCalendar) {
    server.tool(
      'schedule_from_email',
      'A partir de un correo que pide reunion: propone un hueco libre y crea el evento (con Teams). Requiere confirm:true para crear.',
      {
        emailId: z.string().describe('id del correo que pide la reunion'),
        windowStart: z.string().describe('Inicio de la franja donde buscar hueco, ISO 8601'),
        windowEnd: z.string().describe('Fin de la franja, ISO 8601'),
        durationMinutes: z.number().int().min(15).optional().describe('Duracion de la reunion (default 30)'),
        onlineMeeting: z.boolean().optional().describe('Crear reunion de Teams (default true)'),
        confirm: z.boolean().optional(),
      },
      async (args) => {
        try {
          const dur = args.durationMinutes || 30
          const online = args.onlineMeeting !== false
          const mail = await graphGet(`/me/messages/${encodeURIComponent(args.emailId)}?$select=subject,from,toRecipients`)
          const requester = mail.from?.emailAddress?.address
          const subject = `Reunion: ${mail.subject || 'sin asunto'}`

          // Buscar primer hueco libre suficiente en la franja.
          const evPath = `/me/calendarView?startDateTime=${encodeURIComponent(args.windowStart)}` +
            `&endDateTime=${encodeURIComponent(args.windowEnd)}&$select=start,end&$top=100&$orderby=start/dateTime`
          const events = await graphGetAll(evPath, 100)
          const busy = events.map(eventToBusy)
          const slots = freeSlots(Date.parse(args.windowStart), Date.parse(args.windowEnd), busy, dur)
          if (!slots.length) {
            return ok('No encontre un hueco libre suficiente en esa franja. Prueba con otra franja o menor duracion.')
          }
          const chosenStart = new Date(slots[0].start)
          const chosenEnd = new Date(slots[0].start + dur * 60 * 1000)

          if (!isConfirmed(args)) {
            return ok(buildPreview('Agendar reunion desde correo', {
              'Titulo': subject,
              'Con': requester,
              'Inicio propuesto': chosenStart.toISOString(),
              'Fin propuesto': chosenEnd.toISOString(),
              'Teams': online ? 'si' : 'no',
            }))
          }

          const ev = await graphPost('/me/events', {
            subject,
            start: { dateTime: chosenStart.toISOString(), timeZone: 'UTC' },
            end: { dateTime: chosenEnd.toISOString(), timeZone: 'UTC' },
            attendees: requester ? [{ emailAddress: { address: requester }, type: 'required' }] : [],
            isOnlineMeeting: online,
            onlineMeetingProvider: online ? 'teamsForBusiness' : undefined,
          })
          return json({
            created: true,
            event: { id: ev.id, subject: ev.subject, start: ev.start?.dateTime, end: ev.end?.dateTime, joinUrl: ev.onlineMeeting?.joinUrl || '' },
            note: 'Evento creado. La invitacion se envio al solicitante.',
          })
        } catch (e) { return bad(e.message) }
      },
    )
  }

  // ── Pendientes de Teams ─────────────────────────────────────────
  if (caps.readChat) {
    server.tool(
      'teams_pending',
      'Revisa tus chats de Teams y lista los que tienen menciones o preguntas dirigidas a ti sin responder por ti.',
      { chats: z.number().int().min(1).max(30).optional().describe('Cuantos chats recientes revisar (default 15)') },
      async ({ chats }) => {
        try {
          const me = await getMe().catch(() => ({ id: '', name: '' }))
          const maxChats = chats || 15
          const chatList = await graphGetAll(
            `/me/chats?$expand=members&$top=${maxChats}&$orderby=lastUpdatedDateTime desc`, maxChats,
          )
          const pending = []
          for (const c of chatList) {
            const msgs = await graphGetAll(`/me/chats/${encodeURIComponent(c.id)}/messages?$top=10`, 10)
              .catch(() => [])
            if (!msgs.length) { continue }
            // Mensajes en orden: el mas reciente primero.
            const last = msgs[0]
            const lastFromMe = last?.from?.user?.id && me.id && last.from.user.id === me.id
            const mentionsMe = msgs.some(m =>
              (m.mentions || []).some(x => (x.mentionText || '').toLowerCase().includes((me.name || '').toLowerCase()) && me.name),
            )
            const hasQuestion = (last?.body?.content || '').includes('?')
            if (!lastFromMe && (mentionsMe || hasQuestion)) {
              pending.push({
                chatId: c.id,
                topic: c.topic || (c.members || []).map(m => m.displayName).filter(Boolean).join(', '),
                lastFrom: last?.from?.user?.displayName || '',
                lastText: (last?.body?.content || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 200),
                mentionsYou: mentionsMe,
                looksLikeQuestion: hasQuestion,
              })
            }
          }
          return json({ count: pending.length, pending })
        } catch (e) { return bad(e.message) }
      },
    )
  }
}
