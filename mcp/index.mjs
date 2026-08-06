#!/usr/bin/env node
/**
 * Servidor MCP de Outlook (correo + agenda) para VS Code / cualquier cliente MCP.
 *
 * Motor: Outlook de ESCRITORIO por COM (PowerShell). NO usa Microsoft Graph ni
 * Entra ID — maneja el Outlook que ya tienes abierto y firmado, evitando el
 * AADSTS65002 / consentimiento de admin del tenant. Windows-only.
 *
 * Extra: conector opcional con el Anotador de minutas de PolibioDesk.
 */
import { readFileSync } from 'node:fs'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { loadConfig } from './config.mjs'
import { registerOutlookComTools } from './outlookCom.mjs'
import { registerPolibioTools } from './polibio.mjs'

let version = '0.0.0'
try { version = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version || version } catch { /* fallback */ }

const cfg = loadConfig()

const server = new McpServer({ name: 'outlook-mcp', version })
registerOutlookComTools(server)
registerPolibioTools(server, cfg.polibio)

const transport = new StdioServerTransport()
await server.connect(transport)
