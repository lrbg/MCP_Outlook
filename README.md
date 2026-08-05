# Microsoft 365 MCP (Outlook + Teams)

Extensión de VS Code + servidor MCP que le da a tu agente de Copilot (o a
cualquier cliente MCP) acceso a tu **correo, agenda y chats de Teams** a través
de Microsoft Graph, **reutilizando la sesión de Microsoft con la que ya estás
firmado en VS Code**. No registra ninguna app en Azure ni guarda credenciales:
el token vive solo en el almacenamiento local de la extensión.

## Cómo funciona

1. La extensión pide un token de Microsoft Graph con el proveedor `microsoft`
   integrado de VS Code (`vscode.authentication.getSession`), usando tu cuenta de
   organización.
2. Escribe ese token en `globalStorage/m365-config.json` (fuera del repo) y lo
   refresca solo cada ~40 min mientras VS Code está abierto.
3. El servidor MCP (`mcp/index.mjs`) lee ese token y expone las herramientas.
   Copilot Chat las descubre por el provider MCP nativo de VS Code (o por
   `mcp.json` como respaldo).

## Instalación (desarrollo)

```bash
npm install          # deps de la extensión
cd mcp && npm install # deps del servidor MCP
cd .. && npm run compile
```

En VS Code:

1. `M365: Iniciar sesión (Microsoft)` — inicia sesión y detecta qué permisos
   concede tu organización.
2. `M365: Registrar servidor MCP (Copilot Chat)` — registra el servidor.
3. En Copilot Chat presiona **Start** en el servidor "Microsoft 365".

Comandos útiles: `M365: Refrescar sesión`, `M365: Estado y permisos concedidos`.

## Permisos (scopes de Graph)

Se piden en una **escalera** de mayor a menor: si tu tenant no concede alguno,
se prueba un conjunto más chico. Cada herramienta se habilita según lo que se
haya concedido de verdad.

| Permiso | Para qué | Riesgo de consentimiento de admin |
|---|---|---|
| `Mail.ReadWrite` | leer correo, crear borradores | bajo |
| `Mail.Send` | enviar / responder correo | medio |
| `Calendars.ReadWrite` | leer y gestionar agenda | bajo |
| `Chat.ReadWrite` | leer y enviar chats de Teams | medio-alto |

Si un permiso avanzado necesita una app registrada por tu organización, puedes
apuntar a ella sin tocar el código con las opciones `m365.clientId` y
`m365.tenantId` (se inyectan como `VSCODE_CLIENT_ID` / `VSCODE_TENANT`). Esos
valores viven en tu configuración local, nunca en el repo.

## Herramientas

**Correo:** `mail_list`, `mail_read`, `mail_draft`, `mail_send`, `mail_reply`
**Agenda:** `calendar_list`, `calendar_find_free`, `calendar_create`,
`calendar_update`, `calendar_cancel`
**Teams:** `teams_chats_list`, `teams_messages_read`, `teams_send_message`
**Skills (compuestas):** `daily_briefing`, `inbox_triage`, `schedule_from_email`,
`teams_pending`
**Diagnóstico:** `m365_status`

## Seguridad

- **Confirmación en dos pasos** para todo lo que sale al exterior (enviar o
  responder correo, enviar a Teams, crear/mover/cancelar reuniones con
  invitados): la herramienta primero devuelve un *preview* y solo ejecuta si se
  la vuelve a llamar con `confirm: true`.
- El token de Graph nunca se versiona (`m365-config.json` está en `.gitignore`).
- El repo no contiene tenant, dominios de correo ni credenciales.

## Pruebas

```bash
npm test   # lógica pura: OData, triage, huecos libres, confirmación
```

## Licencia

MIT
