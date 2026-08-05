#!/usr/bin/env node
/**
 * Servidor MCP de Microsoft 365 (Outlook + Teams) para VS Code / cualquier
 * cliente MCP. Reutiliza el token de Graph de la sesion de Microsoft de VS Code
 * (escrito por la extension en el archivo apuntado por M365_CONFIG_FILE).
 *
 * Las herramientas se registran segun los permisos realmente concedidos, para
 * degradar con elegancia cuando el administrador del tenant no otorga alguno.
 */
import { readFileSync } from 'node:fs'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { loadConfig, capabilities } from './config.mjs'
import { registerMailTools } from './mail.mjs'
import { registerCalendarTools } from './calendar.mjs'
import { registerTeamsTools } from './teams.mjs'
import { registerSkillTools } from './skills.mjs'
import { registerDiagnosticsTools } from './diagnostics.mjs'

let version = '0.0.0'
try { version = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version || version } catch { /* fallback */ }

const cfg = loadConfig()
const caps = capabilities(cfg.scopes)

const server = new McpServer({ name: 'm365-mcp', version })
registerDiagnosticsTools(server, caps)
registerMailTools(server, caps)
registerCalendarTools(server, caps)
registerTeamsTools(server, caps)
registerSkillTools(server, caps)

const transport = new StdioServerTransport()
await server.connect(transport)
