/**
 * Gate de confirmacion para acciones de escritura hacia afuera (enviar correo,
 * responder, mandar a Teams, crear/cancelar reunion con invitados).
 * Logica pura: arma el texto de preview. La ejecucion real vive en cada tool.
 */

/**
 * Devuelve un bloque de preview legible. `action` describe que se hara,
 * `fields` es un objeto etiqueta -> valor con los datos relevantes.
 */
export function buildPreview(action, fields = {}) {
  const lines = [`ACCION PENDIENTE DE CONFIRMACION: ${action}`, '']
  for (const [label, value] of Object.entries(fields)) {
    if (value === undefined || value === null || value === '') { continue }
    lines.push(`${label}: ${value}`)
  }
  lines.push('')
  lines.push('Nada se ha enviado todavia. Para ejecutarlo, vuelve a llamar la')
  lines.push('misma herramienta con los mismos parametros y confirm: true.')
  return lines.join('\n')
}

/** ¿La llamada trae confirmacion explicita? */
export function isConfirmed(args) {
  return args?.confirm === true
}
