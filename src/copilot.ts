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
