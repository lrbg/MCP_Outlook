/**
 * Clasifica un correo segun si te toca actuar (sin vscode, testeable):
 *  - 'directed'    : vas en Para (To) por email o por nombre.
 *  - 'mentioned'   : te mencionan/taggean en asunto o cuerpo (@Batres, Rogelio).
 *  - 'informative' : ni dirigido ni mencionado (tipicamente solo en copia).
 */

export interface Me { name: string; email: string; tokens: string[] }

export interface ClassifiableEmail {
  subject?: string
  to?: string
  cc?: string
  bodySnippet?: string
}

export type PriorityLabel = 'directed' | 'mentioned' | 'informative'

export interface Classification { label: PriorityLabel; needsAction: boolean; onlyCc: boolean }

export function classifyPriority(email: ClassifiableEmail, me: Me): Classification {
  const to = (email.to || '').toLowerCase()
  const cc = (email.cc || '').toLowerCase()
  const hay = `${email.subject || ''} ${email.bodySnippet || ''}`.toLowerCase()
  const nameL = (me.name || '').toLowerCase()
  const emailL = (me.email || '').toLowerCase()

  const inField = (field: string) =>
    (!!emailL && field.includes(emailL)) || (!!nameL && field.includes(nameL))

  const directed = inField(to)
  const tokens = [nameL, ...(me.tokens || []).map(t => t.toLowerCase())].filter(Boolean)
  const mentioned = tokens.some(t => hay.includes(t))
  const onlyCc = !directed && inField(cc)

  const label: PriorityLabel = directed ? 'directed' : (mentioned ? 'mentioned' : 'informative')
  return { label, needsAction: directed || mentioned, onlyCc }
}

/** Etiqueta legible en espanol. */
export function labelText(l: PriorityLabel): string {
  return l === 'directed' ? 'Para ti' : l === 'mentioned' ? 'Te mencionan' : 'Informativo'
}
