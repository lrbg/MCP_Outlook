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
export interface RepoStat { commits: number; prs: number | null }

const REQ_TIMEOUT_MS = 20000
const MAX_REPOS = 150
const CONCURRENCY = 3

/** Interpreta un 403: distingue SSO/SAML, limite de tasa y permisos. */
function msg403(headers: Headers, json: any): string {
  const sso = headers.get('x-github-sso') || ''
  if (/required/i.test(sso)) {
    const m = sso.match(/url=([^;,\s]+)/)
    return 'Tu sesion de GitHub NO esta autorizada para esta organizacion (SSO/SAML). Autorizala aqui y reintenta: ' + (m ? m[1] : `https://github.com/orgs/<org>/sso`)
  }
  const remaining = headers.get('x-ratelimit-remaining')
  const retry = headers.get('retry-after')
  if (remaining === '0' || retry) { return `Limite de tasa de GitHub. Espera ${retry ? retry + 's' : '~1 min'} y reintenta.` }
  const body = String(json?.message || '')
  if (/secondary rate/i.test(body)) { return 'Limite secundario de GitHub (muchas peticiones). Espera ~1 min y reintenta (o baja el rango).' }
  return body || 'GitHub 403.'
}

export async function getGithubToken(interactive = false): Promise<string> {
  try {
    const s = await vscode.authentication.getSession('github', ['repo', 'read:org'], interactive ? { createIfNone: true } : { silent: true } as any)
    return s?.accessToken || ''
  } catch (e: any) { log(`getGithubToken error: ${e?.message || e}`); return '' }
}

async function gh(url: string, token: string): Promise<{ status: number; json: any; headers: Headers }> {
  const ctrl = new AbortController()
  const timer = setTimeout(() => ctrl.abort(), REQ_TIMEOUT_MS)
  try {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json', 'User-Agent': 'm365-mcp-plugin' },
      signal: ctrl.signal,
    })
    let json: any = null
    try { json = await res.json() } catch { /* vacio */ }
    return { status: res.status, json, headers: res.headers }
  } catch (e: any) {
    if (e?.name === 'AbortError') { throw new Error('GitHub no respondio a tiempo (timeout).') }
    throw e
  } finally { clearTimeout(timer) }
}

/**
 * Detecta los usuarios (login) que estan commiteando en los repos activos del
 * rango, con su conteo. Sirve para EMU, donde el correo corporativo no mapea la
 * cuenta y hay que usar el usuario de GitHub.
 */
export async function detectAuthors(token: string, org: string, sinceISO: string, untilISO: string, onProgress?: (m: string) => void): Promise<{ login: string; count: number }[]> {
  const repos = await listActiveRepos(token, org, sinceISO)
  onProgress?.(`${repos.length} repos activos · detectando autores…`)
  const counts: Record<string, number> = {}
  let done = 0
  await pool(repos, CONCURRENCY, async (r) => {
    try {
      const { json } = await gh(`https://api.github.com/repos/${r.owner}/${r.name}/commits?since=${sinceISO}T00:00:00Z&until=${untilISO}T23:59:59Z&per_page=100`, token)
      if (Array.isArray(json)) {
        for (const it of json) {
          const login = it.author?.login || it.commit?.author?.name || '(sin usuario)'
          counts[login] = (counts[login] || 0) + 1
        }
      }
    } catch (e: any) { log(`detectAuthors ${r.name}: ${e?.message || e}`) }
    done++
    if (done % 10 === 0) { onProgress?.(`autores: ${done}/${repos.length} repos`) }
  })
  const out = Object.entries(counts).map(([login, count]) => ({ login, count })).sort((a, b) => b.count - a.count).slice(0, 100)
  log(`detectAuthors: ${out.length} usuarios`)
  return out
}

/** Igual que gh() pero tambien devuelve la cabecera Link (para contar por paginacion). */
async function ghFull(url: string, token: string): Promise<{ status: number; json: any; link: string }> {
  const ctrl = new AbortController()
  const timer = setTimeout(() => ctrl.abort(), REQ_TIMEOUT_MS)
  try {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json', 'User-Agent': 'm365-mcp-plugin' }, signal: ctrl.signal })
    let json: any = null
    try { json = await res.json() } catch { /* vacio */ }
    return { status: res.status, json, link: res.headers.get('link') || '' }
  } catch (e: any) {
    if (e?.name === 'AbortError') { throw new Error('GitHub no respondio a tiempo (timeout).') }
    throw e
  } finally { clearTimeout(timer) }
}

function lastPage(link: string): number | null {
  const m = link.match(/[?&]page=(\d+)[^>]*>;\s*rel="last"/)
  return m ? Number(m[1]) : null
}

/** Numero de commits del repo en el rango (todos los autores). Truco per_page=1 + Link. */
async function repoCommitCount(token: string, owner: string, name: string, sinceISO: string, untilISO: string): Promise<number> {
  const url = `https://api.github.com/repos/${owner}/${name}/commits?since=${sinceISO}T00:00:00Z&until=${untilISO}T23:59:59Z&per_page=1`
  const { status, json, link } = await ghFull(url, token)
  if (status === 409 || status === 404) { return 0 }
  const lp = lastPage(link)
  if (lp) { return lp }
  return Array.isArray(json) ? json.length : 0
}

let prSearchBlocked = false
/** Numero de PRs creados en el rango. null si la busqueda no esta disponible. */
async function repoPrCount(token: string, owner: string, name: string, sinceISO: string, untilISO: string): Promise<number | null> {
  if (prSearchBlocked) { return null }
  const q = `repo:${owner}/${name} type:pr created:${sinceISO}..${untilISO}`
  const { status, json } = await ghFull(`https://api.github.com/search/issues?q=${encodeURIComponent(q)}&per_page=1`, token)
  if (status === 403) { prSearchBlocked = true; return null }
  if (status !== 200) { return null }
  return typeof json?.total_count === 'number' ? json.total_count : null
}

/** Commits y PRs por repo (para los repos activos del mes). Devuelve mapa full->stat. */
export async function getReposStats(token: string, repos: RepoInfo[], sinceISO: string, untilISO: string, onProgress?: (m: string) => void): Promise<Record<string, RepoStat>> {
  prSearchBlocked = false
  const map: Record<string, RepoStat> = {}
  let done = 0
  await pool(repos, 3, async (r) => {
    let commits = 0; let prs: number | null = null
    try { commits = await repoCommitCount(token, r.owner, r.name, sinceISO, untilISO) } catch (e: any) { log(`stats ${r.name} commits: ${e?.message || e}`) }
    try { prs = await repoPrCount(token, r.owner, r.name, sinceISO, untilISO) } catch (e: any) { log(`stats ${r.name} prs: ${e?.message || e}`) }
    map[r.full] = { commits, prs }
    done++
    if (done % 10 === 0) { onProgress?.(`stats: ${done}/${repos.length} repos`) }
  })
  log(`stats: listo (${repos.length} repos)`)
  return map
}

/** Repos de la org con push desde `sinceISO`, ordenados por push desc. */
export async function listActiveRepos(token: string, org: string, sinceISO: string): Promise<{ owner: string; name: string }[]> {
  const out: { owner: string; name: string }[] = []
  const sinceMs = Date.parse(sinceISO)
  for (let page = 1; page <= 6 && out.length < MAX_REPOS; page++) {
    log(`repos: pidiendo pagina ${page} de ${org}…`)
    const { status, json, headers } = await gh(`https://api.github.com/orgs/${encodeURIComponent(org)}/repos?per_page=100&sort=pushed&direction=desc&page=${page}`, token)
    if (status === 404) { throw new Error(`No encuentro la organizacion "${org}" o sin acceso. Usa el login exacto (el de la URL github.com/ESTO).`) }
    if (status === 403) { throw new Error(msg403(headers, json)) }
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
    const { status, json, headers } = await gh(`https://api.github.com/orgs/${encodeURIComponent(org)}/repos?per_page=100&sort=pushed&direction=desc&page=${page}`, token)
    if (status === 404) { throw new Error(`No encuentro la organizacion "${org}" o sin acceso.`) }
    if (status === 401) { throw new Error('Sesion de GitHub invalida (401). Vuelve a conectar.') }
    if (status === 403) { throw new Error(msg403(headers, json)) }
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

let limited = false
/** Commits (hasta 100) de un repo en el rango, TODOS los autores. */
async function repoAllCommits(token: string, owner: string, name: string, sinceISO: string, untilISO: string): Promise<any[]> {
  if (limited) { return [] }
  const url = `https://api.github.com/repos/${owner}/${name}/commits?since=${sinceISO}T00:00:00Z&until=${untilISO}T23:59:59Z&per_page=100`
  const { status, json } = await gh(url, token)
  if (status === 403) { limited = true; return [] }
  if (status === 409 || status === 404) { return [] }
  return Array.isArray(json) ? json : []
}

/**
 * Reporte en UNA sola pasada por los repos activos (evita el 403 por exceso):
 * por repo cuenta commits (y PRs si includePRs) y reparte los commits por autor.
 */
export async function getGithubReport(
  token: string, org: string, authors: string[], sinceISO: string, untilISO: string,
  includePRs: boolean, onProgress?: (m: string) => void,
): Promise<{ repoCount: number; repos: any[]; members: MemberResult[]; limited: boolean }> {
  limited = false
  prSearchBlocked = false
  log(`== GitHub report: org=${org} operadores=${authors.length} ${sinceISO}..${untilISO} PRs=${includePRs} ==`)
  const active = await listActiveRepos(token, org, sinceISO)
  onProgress?.(`${active.length} repos activos · leyendo commits…`)
  const authSet = new Map<string, string>() // lower -> original
  authors.forEach(a => authSet.set(a.toLowerCase(), a))
  const memberMap: Record<string, Commit[]> = {}
  authors.forEach(a => { memberMap[a] = [] })
  const repos: any[] = []
  let done = 0
  let grand = 0
  await pool(active, CONCURRENCY, async (r) => {
    const raw = await repoAllCommits(token, r.owner, r.name, sinceISO, untilISO)
    for (const it of raw) {
      const login = String(it.author?.login || it.commit?.author?.name || '')
      const orig = authSet.get(login.toLowerCase())
      if (orig) {
        memberMap[orig].push({
          repo: `${r.owner}/${r.name}`,
          message: String(it.commit?.message || '').split('\n')[0].slice(0, 160),
          date: it.commit?.author?.date || it.commit?.committer?.date || '',
          sha: (it.sha || '').slice(0, 7), url: it.html_url || '',
        })
      }
    }
    grand += raw.length
    let prs: number | null = null
    if (includePRs) { try { prs = await repoPrCount(token, r.owner, r.name, sinceISO, untilISO) } catch { prs = null } }
    repos.push({ full: `${r.owner}/${r.name}`, pushed: (r as any).pushed || '', commits: raw.length, prs })
    done++
    if (done % 10 === 0) { onProgress?.(`${done}/${active.length} repos · ${grand} commits`) }
  })
  repos.sort((a, b) => (b.commits || 0) - (a.commits || 0))
  const members: MemberResult[] = authors.map(a => {
    const cs = memberMap[a].sort((x, y) => String(y.date).localeCompare(String(x.date)))
    return { author: a, commits: cs }
  })
  log(`== GitHub report: listo · ${grand} commits, limitado=${limited} ==`)
  return { repoCount: active.length, repos, members, limited }
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
