/**
 * Sanea la salida de PowerShell antes de JSON.parse. Windows PowerShell 5.1
 * (ConvertTo-Json) deja caracteres de control crudos dentro de las cadenas
 * (del cuerpo del correo), que rompen JSON.parse. Aqui se quitan en JS —
 * robusto, sin depender del regex de PowerShell. Cubre C0 (0x00-0x1F), DEL y
 * C1 (0x7F-0x9F), y separadores de linea/parrafo unicode.
 */
const BAD = new RegExp('[\\u0000-\\u001F\\u007F-\\u009F\\u2028\\u2029]', 'g')

export function cleanPsJson(out: string): string {
  return (out || '').replace(BAD, ' ')
}

/** Parsea la salida saneada. Vacio => []. */
export function parsePsArray(out: string): any[] {
  if (!out || !out.trim()) { return [] }
  const j = JSON.parse(cleanPsJson(out))
  return Array.isArray(j) ? j : [j]
}
