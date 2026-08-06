# Outlook MCP (correo + agenda)

Extensión de VS Code + servidor MCP que le da a tu agente de Copilot (o a
cualquier cliente MCP) acceso a tu **correo y agenda de Outlook** usando el
**Outlook de escritorio que ya tienes abierto y firmado**, por automatización COM
vía PowerShell. **No usa Microsoft Graph ni Entra ID**, así que evita por completo
el error `AADSTS65002` y el consentimiento de administrador del tenant.

> **Windows-only.** Requiere Windows con Outlook de escritorio instalado y con tu
> cuenta ya configurada. (El chat de Teams no está disponible en este modo: COM
> no lo alcanza.)

## Cómo funciona

1. El servidor MCP (`mcp/index.mjs`) ejecuta scripts de PowerShell que hablan con
   `Outlook.Application` (COM). Como Outlook ya está autenticado en tu sesión de
   Windows, no hay login, ni token, ni permisos de Graph que aprobar.
2. Los datos que tú das (destinatario, asunto…) se pasan por variables de entorno,
   no interpolados en el script, para evitar inyección de PowerShell.
3. Copilot Chat descubre el servidor por el provider MCP nativo de VS Code (o por
   `mcp.json` como respaldo).

## Instalación

**Windows:**
```bash
curl.exe -L -o outlook-mcp.vsix https://github.com/lrbg/MCP_Outlook/releases/latest/download/m365-mcp-plugin-0.3.0.vsix
code --install-extension outlook-mcp.vsix --force
```

Reinicia VS Code. Verás el icono del sobre "Microsoft 365" en la barra izquierda.
En su panel (o en la paleta `Ctrl+Shift+P`):

1. `Outlook: Registrar servidor MCP (Copilot Chat)` → en Copilot Chat presiona
   **Start** en el servidor "Microsoft 365".
2. Pregúntale al agente algo como "lista mis correos no leídos" o usa la tool
   `outlook_status` para verificar que Outlook responde.

## Herramientas

**Correo:** `outlook_list_emails`, `outlook_read_email`, `outlook_send_email`,
`outlook_reply_email`
**Agenda:** `outlook_list_calendar`, `outlook_shared_calendar`,
`outlook_create_event`, `outlook_find_free`
**Skills:** `daily_briefing` (correos importantes + reuniones de hoy),
`inbox_triage` (clasifica no-leídos por prioridad y acción)
**Diagnóstico:** `outlook_status`

### Conector con PolibioDesk (Anotador de Reuniones)

Opcional. Deja que el agente lea tus **minutas y transcripciones** del módulo de
minutas de PolibioDesk, las cruce con tu correo y agenda, y en planeación **liste
tus pendientes** (los extrae el agente del texto, distinguiendo *tus* acuerdos de
los de otros con tu identidad de Outlook).

Tools: `polibio_minutas_list`, `polibio_minuta_get`, `meeting_context`,
`my_action_items`.

Requiere una edge function de lectura en PolibioDesk (`minutas-read`,
`verify_jwt=false`, token propio). Para activarlo:

1. Ajustes: `m365.polibio.supabaseUrl` (ej. `https://TU-REF.supabase.co`) y
   `m365.polibio.anonKey`.
2. Comando `Outlook: Guardar token de Polibio (minutas)` — guarda el token de
   lectura (cifrado en SecretStorage).

El token y las URLs viven solo en tu configuración local, nunca en el repo.

## Seguridad

- **Confirmación en dos pasos** para enviar y responder correo: la herramienta
  primero devuelve un *preview* y solo envía si se la vuelve a llamar con
  `confirm: true`.
- Los datos de usuario se pasan por variables de entorno (sin inyección).
- El repo no contiene tenant, dominios de correo ni credenciales.

## Pruebas

```bash
npm test   # lógica pura: triage, huecos libres, confirmación
```

## Licencia

MIT
