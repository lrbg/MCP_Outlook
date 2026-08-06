# Outlook MCP (correo + agenda)

Extensión de VS Code + servidor MCP que le da a tu agente de Copilot acceso a tu
**correo y agenda de Outlook** usando el **Outlook de escritorio que ya tienes
abierto y firmado**, por automatización COM vía PowerShell. **No usa Microsoft
Graph ni Entra ID**, así que evita por completo el error `AADSTS65002` y el
consentimiento de administrador del tenant.

> **Windows-only.** Requiere Windows con Outlook de escritorio instalado y con tu
> cuenta ya configurada.

## Cómo funciona

El servidor MCP (`mcp/index.mjs`) ejecuta scripts de PowerShell que hablan con
`Outlook.Application` (COM). Como Outlook ya está autenticado en tu sesión de
Windows, no hay login, ni token, ni permisos que aprobar. Los datos que tú das
(destinatario, asunto…) se pasan por variables de entorno, no interpolados en el
script, para evitar inyección de PowerShell.

## Instalación

**Windows:**
```bash
curl.exe -L -o outlook-mcp.vsix https://github.com/lrbg/MCP_Outlook/releases/download/v0.4.0/m365-mcp-plugin-0.4.0.vsix
code --install-extension outlook-mcp.vsix --force
```

Reinicia VS Code. Verás el icono del sobre en la barra izquierda. En su panel (o
en la paleta `Ctrl+Shift+P`):

1. `Outlook: Registrar servidor MCP (Copilot Chat)` → en Copilot Chat presiona
   **Start** en el servidor "Outlook".
2. Pídele al agente, por ejemplo, "lista mis correos recientes".

## Herramientas

- `outlook_list_emails` — lista correos recientes (Inbox / SentItems / Drafts).
- `outlook_read_email` — lee un correo completo por su EntryID.
- `outlook_send_email` — envía un correo.
- `outlook_list_calendar` — próximos eventos de tu calendario.
- `outlook_shared_calendar` — calendario de otro usuario (con permisos).
- `outlook_create_event` — crea un evento en tu calendario.

## Licencia

MIT
