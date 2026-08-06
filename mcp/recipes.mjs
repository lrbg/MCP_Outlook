#!/usr/bin/env node
/**
 * Servidor MCP de RECETAS de sitios: guarda y entrega procedimientos en markdown
 * ("como llenar los KPI en algo.com") para que el agente tenga el contexto de
 * como operar un sitio con las manos de Playwright. Carpeta en env RECIPES_DIR.
 */
import { Server } from '@modelcontextprotocol/sdk/server/index.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { CallToolRequestSchema, ListToolsRequestSchema, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js'
import { readdirSync, readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs'
import { join } from 'node:path'

const DIR = process.env.RECIPES_DIR || join(process.cwd(), 'recipes')
try { mkdirSync(DIR, { recursive: true }) } catch { /* ya existe */ }

const slug = (s) => String(s || '').toLowerCase().trim().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '')
const fileFor = (name) => join(DIR, `${slug(name)}.md`)
const firstLine = (txt) => (txt.split('\n').find(l => l.trim()) || '').replace(/^#+\s*/, '').slice(0, 120)

const server = new Server({ name: 'outlook-recipes', version: '1.0.0' }, { capabilities: { tools: {} } })

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: 'recipe_list',
      description: 'Lista las recetas de sitios guardadas (procedimientos paso a paso para operar sitios web).',
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'recipe_get',
      description: 'Devuelve el contenido (markdown) de una receta por su nombre, para seguir sus pasos.',
      inputSchema: {
        type: 'object',
        properties: { name: { type: 'string', description: 'Nombre de la receta (de recipe_list)' } },
        required: ['name'],
      },
    },
    {
      name: 'recipe_save',
      description: 'Guarda o actualiza una receta de sitio (procedimiento en markdown) para reutilizarla despues.',
      inputSchema: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'Nombre corto de la receta, ej. "kpi-algo"' },
          content: { type: 'string', description: 'Procedimiento en markdown (pasos para operar el sitio)' },
        },
        required: ['name', 'content'],
      },
    },
  ],
}))

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params
  try {
    if (name === 'recipe_list') {
      const files = existsSync(DIR) ? readdirSync(DIR).filter(f => f.endsWith('.md')) : []
      const recipes = files.map(f => {
        const txt = readFileSync(join(DIR, f), 'utf8')
        return { name: f.replace(/\.md$/, ''), title: firstLine(txt) }
      })
      return { content: [{ type: 'text', text: JSON.stringify({ dir: DIR, count: recipes.length, recipes }, null, 2) }] }
    }
    if (name === 'recipe_get') {
      const path = fileFor(args?.name)
      if (!existsSync(path)) { return { content: [{ type: 'text', text: `No existe la receta "${args?.name}". Usa recipe_list para ver las disponibles.` }], isError: true } }
      return { content: [{ type: 'text', text: readFileSync(path, 'utf8') }] }
    }
    if (name === 'recipe_save') {
      if (!args?.name || !args?.content) { throw new McpError(ErrorCode.InvalidParams, 'name y content son requeridos') }
      const path = fileFor(args.name)
      writeFileSync(path, String(args.content), 'utf8')
      return { content: [{ type: 'text', text: `Receta guardada: ${path}` }] }
    }
    throw new McpError(ErrorCode.MethodNotFound, `Herramienta desconocida: ${name}`)
  } catch (err) {
    if (err instanceof McpError) { throw err }
    throw new McpError(ErrorCode.InternalError, `Error en recetas: ${err?.message ?? err}`)
  }
})

async function main() {
  const transport = new StdioServerTransport()
  await server.connect(transport)
}
main()
