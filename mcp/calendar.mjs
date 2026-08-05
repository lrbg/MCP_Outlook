/**
 * Herramientas de agenda (Outlook Calendar / Microsoft Graph).
 *   Lectura: calendar_list, calendar_find_free
 *   Escritura: calendar_create, calendar_update, calendar_cancel (confirm si hay invitados)
 */
import { z } from 'zod'
import { graphGet, graphGetAll, graphPost, graphPatch } from './graph.mjs'
import { buildPreview, isConfirmed } from './lib/preview.mjs'
import { freeSlots, eventToBusy } from './lib/freebusy.mjs'

const ok = (text) => ({ content: [{ type: 'text', text }] })
const bad = (text) => ({ content: [{ type: 'text', text }], isError: true })
const json = (obj) => ok(JSON.stringify(obj, null, 2))

function slimEvent(e) {
  return {
    id: e.id,
    subject: e.subject || '',
    start: e.start?.dateTime || '',
    end: e.end?.dateTime || '',
    timeZone: e.start?.timeZone || 'UTC',
    location: e.location?.displayName || '',
    organizer: e.organizer?.emailAddress?.address || '',
    attendees: (e.attendees || []).map(a => a.emailAddress?.address).filter(Boolean),
    isOnline: !!e.isOnlineMeeting,
    joinUrl: e.onlineMeeting?.joinUrl || '',
    webLink: e.webLink || '',
  }
}

function attendeeObjs(list) {
  return (list || []).map(addr => ({ emailAddress: { address: addr }, type: 'required' }))
}

export function registerCalendarTools(server, caps) {
  if (caps.readCalendar) {
    server.tool(
      'calendar_list',
      'Lista eventos de tu calendario en un rango de fechas (ISO 8601).',
      {
        start: z.string().describe('Inicio del rango, ISO 8601'),
        end: z.string().describe('Fin del rango, ISO 8601'),
        top: z.number().int().min(1).max(100).optional().describe('Maximo de eventos (default 50)'),
      },
      async ({ start, end, top }) => {
        try {
          const max = top || 50
          const select = 'id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting,webLink'
          const path = `/me/calendarView?startDateTime=${encodeURIComponent(start)}` +
            `&endDateTime=${encodeURIComponent(end)}&$select=${select}&$top=${max}&$orderby=start/dateTime`
          const items = await graphGetAll(path, max)
          return json({ count: items.length, events: items.map(slimEvent) })
        } catch (e) { return bad(e.message) }
      },
    )

    server.tool(
      'calendar_find_free',
      'Busca huecos libres en tu agenda dentro de una franja, de al menos N minutos.',
      {
        windowStart: z.string().describe('Inicio de la franja a revisar, ISO 8601'),
        windowEnd: z.string().describe('Fin de la franja, ISO 8601'),
        minMinutes: z.number().int().min(15).optional().describe('Duracion minima del hueco (default 30)'),
      },
      async ({ windowStart, windowEnd, minMinutes }) => {
        try {
          const select = 'id,subject,start,end'
          const path = `/me/calendarView?startDateTime=${encodeURIComponent(windowStart)}` +
            `&endDateTime=${encodeURIComponent(windowEnd)}&$select=${select}&$top=100&$orderby=start/dateTime`
          const events = await graphGetAll(path, 100)
          const busy = events.map(eventToBusy)
          const slots = freeSlots(Date.parse(windowStart), Date.parse(windowEnd), busy, minMinutes || 30)
          return json({
            count: slots.length,
            freeSlots: slots.map(s => ({ start: new Date(s.start).toISOString(), end: new Date(s.end).toISOString() })),
          })
        } catch (e) { return bad(e.message) }
      },
    )
  }

  if (caps.writeCalendar) {
    server.tool(
      'calendar_create',
      'Crea un evento en tu calendario. Puede ser reunion de Teams. Si hay invitados requiere confirm:true.',
      {
        subject: z.string().describe('Titulo del evento'),
        start: z.string().describe('Inicio, ISO 8601'),
        end: z.string().describe('Fin, ISO 8601'),
        timeZone: z.string().optional().describe('Zona horaria IANA (default UTC)'),
        attendees: z.array(z.string()).optional().describe('Invitados (direcciones)'),
        body: z.string().optional().describe('Descripcion/agenda del evento'),
        location: z.string().optional().describe('Ubicacion'),
        onlineMeeting: z.boolean().optional().describe('Crear reunion de Teams'),
        confirm: z.boolean().optional().describe('Requerido si hay invitados'),
      },
      async (args) => {
        const hasAttendees = (args.attendees || []).length > 0
        if (hasAttendees && !isConfirmed(args)) {
          return ok(buildPreview('Crear reunion con invitados', {
            'Titulo': args.subject,
            'Inicio': args.start,
            'Fin': args.end,
            'Invitados': (args.attendees || []).join(', '),
            'Teams': args.onlineMeeting ? 'si' : 'no',
            'Ubicacion': args.location,
          }))
        }
        try {
          const ev = await graphPost('/me/events', {
            subject: args.subject,
            start: { dateTime: args.start, timeZone: args.timeZone || 'UTC' },
            end: { dateTime: args.end, timeZone: args.timeZone || 'UTC' },
            attendees: attendeeObjs(args.attendees),
            body: args.body ? { contentType: 'Text', content: args.body } : undefined,
            location: args.location ? { displayName: args.location } : undefined,
            isOnlineMeeting: !!args.onlineMeeting,
            onlineMeetingProvider: args.onlineMeeting ? 'teamsForBusiness' : undefined,
          })
          return json({ created: true, event: slimEvent(ev) })
        } catch (e) { return bad(e.message) }
      },
    )

    server.tool(
      'calendar_update',
      'Modifica o mueve un evento existente. Si el evento tiene invitados requiere confirm:true.',
      {
        id: z.string().describe('id del evento'),
        subject: z.string().optional(),
        start: z.string().optional().describe('Nuevo inicio, ISO 8601'),
        end: z.string().optional().describe('Nuevo fin, ISO 8601'),
        timeZone: z.string().optional(),
        location: z.string().optional(),
        confirm: z.boolean().optional(),
      },
      async (args) => {
        try {
          const current = await graphGet(`/me/events/${encodeURIComponent(args.id)}?$select=subject,attendees`)
          const hasAttendees = (current.attendees || []).length > 0
          if (hasAttendees && !isConfirmed(args)) {
            return ok(buildPreview('Modificar reunion con invitados', {
              'Evento': current.subject || args.id,
              'Nuevo inicio': args.start,
              'Nuevo fin': args.end,
              'Nueva ubicacion': args.location,
            }))
          }
          const patch = {}
          if (args.subject) { patch.subject = args.subject }
          if (args.start) { patch.start = { dateTime: args.start, timeZone: args.timeZone || 'UTC' } }
          if (args.end) { patch.end = { dateTime: args.end, timeZone: args.timeZone || 'UTC' } }
          if (args.location) { patch.location = { displayName: args.location } }
          const ev = await graphPatch(`/me/events/${encodeURIComponent(args.id)}`, patch)
          return json({ updated: true, event: slimEvent(ev) })
        } catch (e) { return bad(e.message) }
      },
    )

    server.tool(
      'calendar_cancel',
      'Cancela un evento. Si tiene invitados requiere confirm:true (les llega la cancelacion).',
      {
        id: z.string().describe('id del evento'),
        comment: z.string().optional().describe('Mensaje para los invitados'),
        confirm: z.boolean().optional(),
      },
      async (args) => {
        try {
          const current = await graphGet(`/me/events/${encodeURIComponent(args.id)}?$select=subject,attendees`)
          const hasAttendees = (current.attendees || []).length > 0
          if (hasAttendees && !isConfirmed(args)) {
            return ok(buildPreview('Cancelar reunion con invitados', {
              'Evento': current.subject || args.id,
              'Mensaje': args.comment,
            }))
          }
          await graphPost(`/me/events/${encodeURIComponent(args.id)}/cancel`, {
            comment: args.comment || '',
          })
          return ok('Evento cancelado.')
        } catch (e) { return bad(e.message) }
      },
    )
  }
}
