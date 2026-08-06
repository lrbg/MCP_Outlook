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

## Bitácora de bandeja (revisión diaria con Copilot)

Un panel dentro del plugin (icono del sobre → **Bitácora**) con una tabla tipo
bitácora: fecha, no-leídos, remitentes clave y **notas de Copilot**. Se llena
solo: un temporizador diario lee tus no-leídos (COM) y le pide a **tu Copilot**
—vía la API `vscode.lm`, sin abrir el chat— que los resuma. Corre mientras VS
Code esté abierto (se pone al día si estaba cerrado). La primera vez, VS Code
pide tu permiso para que la extensión use tu modelo de Copilot.

Ajustes: `m365.dailyReview.enabled` / `hour` / `maxEmails`.

## Manos de navegador (Playwright) + recetas

El plugin registra el **Playwright MCP oficial** de Microsoft (vía `npx`), que le
da al agente manos para abrir sitios, hacer clic, llenar formularios y leer la
página. Úsalo para tareas como "un correo me pide llenar los KPI en algo.com".

Las **recetas** son procedimientos en markdown (una por sitio) que el agente
puede leer para saber *cómo* operar un sitio. Tools del servidor de recetas:
`recipe_list`, `recipe_get`, `recipe_save`. Abre/edita la carpeta con el comando
`Outlook: Abrir carpeta de recetas`.

Seguridad: **el agente no teclea contraseñas** — tú inicias sesión en el sitio y
el agente opera ya dentro; conviene que confirmes las acciones finales (Guardar/
Enviar). Ajustes: `m365.playwright.enabled`, `m365.recipesDir`.

## Licencia

MIT
