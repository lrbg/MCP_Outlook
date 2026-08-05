# m365MCP_PlugIn — Diseño (v1)

Fecha: 2026-08-05
Estado: aprobado para escribir plan de implementación

## Objetivo

Extensión de VS Code + servidor MCP standalone que le da al agente de Copilot
(o cualquier cliente MCP) acceso al **correo, agenda y chats de Teams** del
usuario, para coordinar correo y calendario desde el chat del agente.

Reusa el mismo esqueleto del plugin `genRocketMCP_PlugIn` (extensión `src/*.ts`
que escribe config local en `globalStorage` + servidor MCP `mcp/*.mjs` que
consume Copilot Chat / el provider MCP nativo de VS Code), **sin nada de
GenRocket**.

## No-objetivos (v1)

- OneDrive / SharePoint / archivos (posible v2; el nombre `m365` deja la puerta).
- Registrar una aplicación propia en Azure Entra ID.
- Panel/dashboard (a diferencia de GenRocket; se puede sumar después).
- Automatización sin supervisión: toda escritura hacia afuera se confirma.

## Autenticación

- **Proveedor Microsoft integrado de VS Code.** Se obtiene el token de Graph con
  `vscode.authentication.getSession('microsoft', scopes, { createIfNone })`,
  reusando la cuenta de organización con la que el usuario ya está firmado en
  VS Code (la misma que habilita su Copilot de equipo). No se registra app en
  Azure; el proveedor de VS Code aporta su client_id.
- **Scopes solicitados:** `User.Read`, `Mail.ReadWrite`, `Mail.Send`,
  `Calendars.ReadWrite`, `Chat.ReadWrite`. (`offline_access` lo maneja el
  proveedor de VS Code.)
- **Degradación elegante:** al activar, la extensión intenta obtener cada grupo
  de scopes. Un scope no concedido por el administrador del tenant **desactiva
  solo su(s) herramienta(s)**; las demás siguen funcionando. El config guarda
  qué grupos quedaron disponibles.
- **Riesgo conocido:** `Mail.Send` y `Chat.ReadWrite` pueden requerir
  consentimiento de administrador del tenant. `Mail.ReadWrite` (lectura) y
  `Calendars.ReadWrite` normalmente pasan sin admin. El diseño no falla en bloque
  si falta consentimiento.

## Manejo de token (opción B — archivo)

Idéntico patrón al de GenRocket (token de GitHub en config):

- La extensión escribe `globalStorage/m365-config.json` con `accessToken` y
  `expiresOn` (más metadatos: `account`, `availableScopes`, `mcpEntry`).
- La extensión **refresca proactivamente** el token: al activar, al ejecutar
  "Registrar servidor MCP", y con un timer (~cada 40 min) mientras VS Code esté
  abierto; cada refresco reescribe el archivo.
- El servidor MCP lee `accessToken` del archivo en cada llamada. Si el token ya
  expiró (o Graph responde 401), la tool devuelve un error claro y accionable:
  "token expirado, ejecuta el comando *M365: Refrescar sesión* en VS Code".
- La ruta del config se pasa al MCP por variable de entorno
  `M365_CONFIG_FILE` (en el provider MCP nativo y en el fallback `mcp.json`).
- El `accessToken` vive solo en `globalStorage`, **nunca en el repo**.

## Estructura del repo

Reusa el layout de `genRocketMCP_PlugIn`:

- `src/` — extensión VS Code (TypeScript):
  - `extension.ts` — activación, comandos, timer de refresco.
  - `auth.ts` — `getSession('microsoft', ...)`, detección de scopes disponibles.
  - `config.ts` — escritura/lectura de `m365-config.json` en `globalStorage`.
  - `mcpProvider.ts` — registro MCP nativo
    (`contributes.mcpServerDefinitionProviders` +
    `vscode.lm.registerMcpServerDefinitionProvider`, `McpStdioServerDefinition`
    → `mcp/index.mjs`, env `M365_CONFIG_FILE`), con feature-detect y fallback al
    comando que escribe `.vscode/mcp.json`.
- `mcp/` — servidor MCP standalone (ESM, stdio), sin dependencia de VS Code:
  - `index.mjs` — arranque del servidor MCP y registro de tools.
  - `graph.mjs` — cliente Graph (fetch + header `Authorization: Bearer`,
    manejo de 401/429, paginación `@odata.nextLink`), lee token del config.
  - `mail.mjs`, `calendar.mjs`, `teams.mjs` — tools base por dominio.
  - `skills.mjs` — tools compuestas (briefing, triage, agendar, pendientes).
  - `me.mjs` — identidad del usuario (`/me`) para separar "lo tuyo".
- `.vscode/mcp.json` — fallback para VS Code sin provider nativo.
- Empaquetado `.vsix` con `npm run package` (sin nombres de cliente en
  comentarios).

## Herramientas base (Microsoft Graph)

Correo:
- `mail_list` — filtra por remitente, asunto, fecha, no-leídos, carpeta; devuelve
  id, remitente, asunto, fecha, snippet, leído/no-leído.
- `mail_read` — cuerpo completo de un correo por id.
- `mail_draft` — crea borrador (no envía).
- `mail_send` — envía correo nuevo. **Requiere `confirm:true`** (ver Seguridad).
- `mail_reply` — responde a un correo. **Requiere `confirm:true`**.

Agenda:
- `calendar_list` — eventos en un rango de fechas.
- `calendar_find_free` — huecos libres en una franja (opcional: varios
  asistentes vía `getSchedule`).
- `calendar_create` — crea evento; opción `onlineMeeting:true` para Teams
  meeting. **Requiere `confirm:true` si hay invitados.**
- `calendar_update` — modifica/mueve un evento. **Requiere `confirm:true` si hay
  invitados.**
- `calendar_cancel` — cancela un evento. **Requiere `confirm:true` si hay
  invitados.**

Teams:
- `teams_chats_list` — lista chats recientes (1:1 y grupales).
- `teams_messages_read` — mensajes de un chat.
- `teams_send_message` — envía mensaje a un chat. **Requiere `confirm:true`.**

## Skills v1 (tools MCP compuestas)

Todas usan `/me` para distinguir *tus* pendientes/acuerdos de los de otros.

- `daily_briefing` — solo lectura. Combina correos importantes sin responder +
  reuniones del día + conflictos de horario + huecos libres → resumen.
- `inbox_triage` — lectura + borradores. Clasifica no-leídos por urgencia y
  remitente, marca los que piden acción tuya, y opcionalmente genera borradores
  de respuesta (no envía).
- `schedule_from_email` — escritura con confirmación. Lee un correo que pide
  reunión, busca hueco común, crea el evento (con Teams meeting) y prepara la
  respuesta de confirmación. El paso de crear/responder pasa por `confirm:true`.
- `teams_pending` — lectura. Chats de Teams con menciones/preguntas dirigidas a
  ti sin responder; opción de enlazar con la agenda ("responder antes de la junta
  de las 4").

## Seguridad

- **Gate de confirmación:** toda tool que produce un efecto externo (enviar o
  responder correo, enviar a Teams, crear/modificar/cancelar reunión con
  invitados) primero devuelve un **preview** estructurado de lo que haría y exige
  una segunda llamada con `confirm:true`. Nada sale en un solo paso. Alineado con
  la política de acciones que requieren permiso explícito.
- **Higiene del repo (regla estricta, como GenRocket):** el repo NO contiene
  tenant, dominio de correo, direcciones, ni credenciales. Todo lo específico del
  usuario vive en `globalStorage`. Validar con `git grep` antes de publicar.
- **Secretos:** `accessToken` solo en `globalStorage`; nunca en el repo ni en
  logs. El secreto/ruta del config se pasa por env al MCP.

## Errores y bordes

- **Token expirado / 401:** tool devuelve mensaje accionable ("refresca la
  sesión") en vez de fallar en crudo.
- **Scope faltante:** la tool afectada responde que la capacidad no está
  concedida por el admin y sugiere las alternativas de solo lectura.
- **429 / throttling de Graph:** respetar `Retry-After`, reintentos con backoff.
- **Paginación:** seguir `@odata.nextLink` hasta un límite razonable.

## Pruebas

- **Lógica pura testeable sin VS Code** (patrón GenRocket `dashboardCore.ts`):
  parsing de respuestas Graph, armado de filtros OData, clasificación de triage,
  cálculo de huecos libres, y ensamblado de previews de confirmación → tests
  unitarios (`node --test`) con respuestas Graph de ejemplo (fixtures).
- **Contrato de las tools:** validar shape de entrada/salida de cada tool.
- **Manual/E2E:** contra la cuenta real del usuario en VS Code (fuera de CI).

## Versionado

- v1 = MVP con lo anterior. Empaquetar `.vsix` limpio y publicar release.
- Futuro (no v1): OneDrive/SharePoint, dashboard de actividad, más skills
  (proteger tiempo, reprogramar en cadena, resumen semanal).
