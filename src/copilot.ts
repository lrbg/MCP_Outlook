import * as vscode from 'vscode'
import { RawEmail } from './bitacoraCore'
import { EmailBody } from './outlookActions'

async function pickModel(): Promise<any> {
  const lm: any = (vscode as any).lm
  if (!lm || typeof lm.selectChatModels !== 'function') {
    throw new Error('Tu VS Code no expone la API de modelos (vscode.lm). Actualiza VS Code.')
  }
  let models = await lm.selectChatModels({ vendor: 'copilot' })
  if (!models || models.length === 0) { models = await lm.selectChatModels({}) }
  if (!models || models.length === 0) {
    throw new Error('No hay ningun modelo de Copilot disponible. Verifica que Copilot este activo en VS Code.')
  }
  return models[0]
}

/** Asistente de agenda: recomendaciones sobre las reuniones del rango. */
export async function assistAgenda(meetings: any[], token: vscode.CancellationToken): Promise<string> {
  const model = await pickModel()
  const compact = meetings.map(m => ({ asunto: m.subject, inicio: m.start, fin: m.end, invitados: m.attendees, lugar: m.location }))
  const prompt =
    'Eres mi asistente de agenda. Abajo van mis reuniones (JSON). Responde en espanol, markdown breve:\n' +
    '1. **Hoy/proximo**: que sigue y que preparar.\n' +
    '2. **Alertas**: empalmes, dias saturados (muchas juntas seguidas), falta de buffer.\n' +
    '3. **Sugerencias**: bloques de foco o mover algo, si aplica.\n' +
    'No inventes reuniones que no esten. Se concreto y corto.\n\n' +
    'Reuniones:\n```json\n' + JSON.stringify(compact, null, 2) + '\n```'
  const resp = await model.sendRequest([vscode.LanguageModelChatMessage.User(prompt)], {}, token)
  let text = ''
  for await (const chunk of resp.text) { text += chunk }
  return text.trim() || '_Sin recomendaciones._'
}

/**
 * Resumen estructurado de los correos prioritarios en UNA sola llamada.
 * Devuelve un arreglo con { id, resumen, reunion:{es,cuando,modo,donde}, respuesta }.
 */
export async function summarizePriority(emails: any[], token: vscode.CancellationToken): Promise<any[]> {
  if (!emails.length) { return [] }
  const model = await pickModel()
  const compact = emails.map(e => ({
    id: e.id, de: e.sender || e.senderEmail || '', asunto: e.subject || '', recibido: e.received || '',
    para: e.to || '', cc: e.cc || '', yaRespondi: !!e.repliedAt,
    miRespuesta: (e.replyBody || '').replace(/\s+/g, ' ').slice(0, 300),
    cuerpo: (e.bodySnippet || '').replace(/\s+/g, ' ').slice(0, 700),
  }))
  const prompt =
    'Eres mi asistente de correo. Te doy mis correos prioritarios en JSON. ' +
    'Devuelve SOLO un arreglo JSON valido (sin texto extra, sin markdown), un objeto por correo, con EXACTAMENTE estas llaves:\n' +
    '- "id": el id tal cual del correo.\n' +
    '- "resumen": 1 frase CORTA, clara y util en espanol de que trata o que piden.\n' +
    '- "reunion": objeto { "es": true|false, "cuando": texto corto o "", "modo": "Teams"|"Presencial"|"", "donde": texto corto o "" }. es=true solo si el correo es sobre una llamada/junta/reunion. Infiere modo/donde del texto: Teams si hay link o menciona Teams/en linea; Presencial si menciona sala/piso/oficina/direccion.\n' +
    '- "respuesta": si ya respondi (yaRespondi=true), 1 linea CORTA de que respondi; si no, "".\n' +
    'No inventes datos que no esten. Responde unicamente el JSON.\n\n' +
    '```json\n' + JSON.stringify(compact) + '\n```'
  const resp = await model.sendRequest([vscode.LanguageModelChatMessage.User(prompt)], {}, token)
  let text = ''
  for await (const chunk of resp.text) { text += chunk }
  text = text.trim().replace(/^```(json)?/i, '').replace(/```$/, '').trim()
  try {
    const arr = JSON.parse(text)
    return Array.isArray(arr) ? arr : []
  } catch { return [] }
}

/** Redacta un borrador de respuesta al correo dado, para que el usuario lo edite. */
export async function draftReply(email: EmailBody, token: vscode.CancellationToken, instruction = ''): Promise<string> {
  const model = await pickModel()
  const prompt =
    'Redacta un BORRADOR de respuesta a este correo, en espanol, tono profesional y cordial, ' +
    'claro y conciso. Escribe SOLO el cuerpo de la respuesta (sin asunto, sin encabezados de correo, ' +
    'sin explicaciones). Deja el saludo y una despedida generica. Si falta informacion, deja un ' +
    '[dato pendiente] entre corchetes.' +
    (instruction ? `\nInstruccion adicional del usuario: ${instruction}` : '') +
    `\n\nCorreo original:\nDe: ${email.sender} <${email.senderEmail}>\nAsunto: ${email.subject}\n\n${email.body}`
  const resp = await model.sendRequest([vscode.LanguageModelChatMessage.User(prompt)], {}, token)
  let text = ''
  for await (const chunk of resp.text) { text += chunk }
  return text.trim()
}

/**
 * Genera las notas de la revision usando TU Copilot, via la API de modelos de
 * lenguaje de VS Code (`vscode.lm`). No abre el chat: la extension le manda el
 * prompt al modelo por su cuenta. La primera vez, VS Code pide tu consentimiento
 * para que la extension use el modelo.
 */
export async function summarizeInbox(inbox: RawEmail[], sent: RawEmail[], token: vscode.CancellationToken): Promise<string> {
  const model = await pickModel()

  const compactIn = inbox.map(e => ({
    de: e.sender || e.senderEmail || '', asunto: e.subject || '', recibido: e.received || '',
    noLeido: !!e.unread, extracto: (e.preview || '').replace(/\s+/g, ' ').slice(0, 300),
  }))
  const compactOut = sent.map(e => ({
    para: e.to || '', asunto: e.subject || '', enviado: e.received || '',
    extracto: (e.preview || '').replace(/\s+/g, ' ').slice(0, 200),
  }))

  const prompt =
    'Eres mi asistente de correo. Abajo van mis correos recientes de la BANDEJA DE ENTRADA y de ENVIADOS (JSON). ' +
    'Responde SIEMPRE en espanol, en markdown breve, con estas secciones:\n' +
    '1. **Requieren mi accion**: los 3-6 de la bandeja de entrada que piden algo de MI y que NO parezcan ya respondidos (revisa Enviados para inferirlo), con una linea de por que.\n' +
    '2. **Ya respondidos / en curso**: hilos donde ya envie respuesta (usa Enviados), 1-3 lineas.\n' +
    '3. **Temas/remitentes clave**: agrupa lo demas en 2-4 lineas.\n' +
    'No inventes correos que no esten en los datos. Se conciso.\n\n' +
    'Bandeja de entrada:\n```json\n' + JSON.stringify(compactIn, null, 2) + '\n```\n' +
    'Enviados:\n```json\n' + JSON.stringify(compactOut, null, 2) + '\n```'

  const resp = await model.sendRequest([vscode.LanguageModelChatMessage.User(prompt)], {}, token)
  let text = ''
  for await (const chunk of resp.text) { text += chunk }
  return text.trim() || '_Copilot no devolvio notas._'
}
