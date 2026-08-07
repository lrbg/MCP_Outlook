import * as vscode from 'vscode'
import { log } from './log'

/**
 * Metricas de commits del equipo via GitHub. Reutiliza la sesion de GitHub de
 * VS Code. Lista los repos activos de la org y, por cada repo, pide los commits
 * filtrando por autor (correo o usuario; la API de commits mapea la cuenta).
 * Con registro (canal "Outlook MCP"), timeout por peticion y concurrencia.
 */

export interface Commit { repo: string; message: string; date: string; sha: string; url: string }
export interface MemberResult { author: string; commits: Commit[]; error?: string }
export interface TeamResult { repoCount: number; members: MemberResult[] }
export interface RepoInfo { owner: string; name: string; full: string; pushed: string }

const REQ_TIMEOUT_MS = 20000
const MAX_REPOS = 150
const CONCURRENCY = 6

export async function getGithubToken(interactive = false): Promise<string> {
  try {
    const s = await vscode.authentication.getSession('github', ['repo', 'read:org'], interactive ? { createIfNone: true } : { silent: true } as any)
    return s?.accessToken || ''
  } catch (e: any) { log(`getGithubToken error: ${e?.message || e}`); return '' }
}

async function gh(url: string, token: string): Promise<{ status: number; json: any }> {
  const ctrl = new AbortController()
  const timer = setTimeout(() => ctrl.abort(), REQ_TIMEOUT_MS)
  try {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json', 'User-Agent': 'm365-mcp-plugin' },
      signal: ctrl.signal,
    })
    let json: any = null
    try { json = await res.json() } catch { /* vacio */ }
    return { status: res.status, json }
  } catch (e: any) {
    if (e?.name === 'AbortError') { throw new Error('GitHub no respondio a tiempo (timeout).') }
    throw e
  } finally { clearTimeout(timer) }
}

/** Repos de la org con push desde `sinceISO`, ordenados por push desc. */
export async function listActiveRepos(token: string, org: string, sinceISO: string): Promise<{ owner: string; name: string }[]> {
  const out: { owner: string; name: string }[] = []
  const sinceMs = Date.parse(sinceISO)
  for (let page = 1; page <= 6 && out.length < MAX_REPOS; page++) {
    log(`repos: pidiendo pagina ${page} de ${org}…`)
    const { status, json } = await gh(`https://api.github.com/orgs/${encodeURIComponent(org)}/repos?per_page=100&sort=pushed&direction=desc&page=${page}`, token)
    if (status === 404) { throw new Error(`No encuentro la organizacion "${org}" o sin acceso. Usa el login exacto (el de la URL github.com/ESTO).`) }
    if (status === 403) { throw new Error('GitHub 403 (permisos/limite). Reconecta concediendo acceso a la organizacion.') }
    if (status === 401) { throw new Error('Sesion de GitHub invalida (401). Vuelve a conectar.') }
    if (!Array.isArray(json) || json.length === 0) { log(`repos: pagina ${page} vacia (status ${status})`); break }
    let stop = false
    for (const r of json) {
      const pushed = Date.parse(r.pushed_at || r.updated_at || '')
      if (Number.isFinite(pushed) && Number.isFinite(sinceMs) && pushed < sinceMs) { stop = true; break }
      out.push({ owner: r.owner?.login || org, name: r.name })
      if (out.length >= MAX_REPOS) { break }
    }
    if (stop || json.length < 100) { break }
  }
  log(`repos: ${out.length} activos desde ${sinceISO}`)
  return out
}

/** Todos los repos de la org (para la lista de la izquierda). */
export async function listAllRepos(token: string, org: string, cap = 400): Promise<RepoInfo[]> {
  const out: RepoInfo[] = []
  for (let page = 1; page <= 8 && out.length < cap; page++) {
    const { status, json } = await gh(`https://api.github.com/orgs/${encodeURIComponent(org)}/repos?per_page=100&sort=pushed&direction=desc&page=${page}`, token)
    if (status === 404) { throw new Error(`No encuentro la organizacion "${org}" o sin acceso.`) }
    if (status === 403 || status === 401) { throw new Error(`GitHub ${status}: revisa permisos/reconecta.`) }
    if (!Array.isArray(json) || json.length === 0) { break }
    for (const r of json) { out.push({ owner: r.owner?.login || org, name: r.name, full: r.full_name || `${org}/${r.name}`, pushed: (r.pushed_at || '').slice(0, 10) }) }
    if (json.length < 100) { break }
  }
  log(`repos totales: ${out.length}`)
  return out
}

async function repoCommits(token: string, owner: string, repo: string, author: string, sinceISO: string, untilISO?: string): Promise<Commit[]> {
  const until = untilISO ? `&until=${untilISO}T23:59:59Z` : ''
  const url = `https://api.github.com/repos/${owner}/${repo}/commits?author=${encodeURIComponent(author)}&since=${sinceISO}T00:00:00Z${until}&per_page=100`
  const { status, json } = await gh(url, token)
  if (status === 409 || status === 404) { return [] }
  if (status === 403) { throw new Error('GitHub 403 (limite de tasa).') }
  if (!Array.isArray(json)) { return [] }
  return json.map((it: any) => ({
    repo: `${owner}/${repo}`,
    message: String(it.commit?.message || '').split('\n')[0].slice(0, 160),
    date: it.commit?.author?.date || it.commit?.committer?.date || '',
    sha: (it.sha || '').slice(0, 7),
    url: it.html_url || '',
  }))
}

/** Ejecuta tareas con concurrencia limitada. */
async function pool<T>(items: T[], limit: number, fn: (t: T, i: number) => Promise<void>): Promise<void> {
  let i = 0
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (i < items.length) { const idx = i++; await fn(items[idx], idx) }
  })
  await Promise.all(workers)
}

/** Commits del equipo: repos activos x autor, con registro y progreso. */
export async function getTeamCommits(
  token: string, org: string, authors: string[], sinceISO: string, untilISO?: string,
  onProgress?: (msg: string) => void,
): Promise<TeamResult> {
  log(`== GitHub metrics: org=${org} operadores=${authors.length} ${sinceISO}..${untilISO || 'hoy'} ==`)
  const repos = await listActiveRepos(token, org, sinceISO)
  onProgress?.(`${repos.length} repos activos · consultando commits…`)
  const members: MemberResult[] = []
  let grand = 0
  for (const a of authors) {
    const commits: Commit[] = []
    let err = ''
    let done = 0
    log(`autor ${a}: revisando ${repos.length} repos…`)
    await pool(repos, CONCURRENCY, async (r) => {
      try { const cs = await repoCommits(token, r.owner, r.name, a, sinceISO, untilISO); if (cs.length) { commits.push(...cs); grand += cs.length } }
      catch (e: any) { err = e?.message || String(e); log(`autor ${a} repo ${r.name}: ${err}`) }
      done++
      if (done % 10 === 0) { onProgress?.(`${a}: ${done}/${repos.length} repos · ${grand} commits`) }
    })
    commits.sort((x, y) => String(y.date).localeCompare(String(x.date)))
    log(`autor ${a}: ${commits.length} commits`)
    members.push({ author: a, commits, error: (err && !commits.length) ? err : undefined })
  }
  log(`== GitHub metrics: listo · ${grand} commits ==`)
  return { repoCount: repos.length, members }
}
