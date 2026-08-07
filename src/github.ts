import * as vscode from 'vscode'

/**
 * Metricas de commits del equipo via GitHub. Reutiliza la sesion de GitHub de
 * VS Code. Estrategia fiable: lista los repos activos de la organizacion y, por
 * cada repo, pide los commits filtrando por autor (la API de commits acepta
 * correo O usuario y mapea la cuenta, a diferencia de la Search API).
 */

export interface Commit { repo: string; message: string; date: string; sha: string; url: string }
export interface MemberResult { author: string; commits: Commit[]; error?: string }
export interface TeamResult { repoCount: number; members: MemberResult[] }

/** Token de la sesion de GitHub de VS Code. interactive=true puede abrir el login. */
export async function getGithubToken(interactive = false): Promise<string> {
  try {
    const s = await vscode.authentication.getSession('github', ['repo', 'read:org'], interactive ? { createIfNone: true } : { silent: true } as any)
    return s?.accessToken || ''
  } catch { return '' }
}

async function gh(url: string, token: string): Promise<{ status: number; json: any }> {
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json', 'User-Agent': 'm365-mcp-plugin' },
  })
  let json: any = null
  try { json = await res.json() } catch { /* puede venir vacio */ }
  return { status: res.status, json }
}

/** Repos de la org con actividad (push) desde `sinceISO`. Ordena por push desc. */
export async function listActiveRepos(token: string, org: string, sinceISO: string, maxRepos = 400): Promise<{ owner: string; name: string; full: string }[]> {
  const out: { owner: string; name: string; full: string }[] = []
  const sinceMs = Date.parse(sinceISO)
  for (let page = 1; page <= 5 && out.length < maxRepos; page++) {
    const { status, json } = await gh(`https://api.github.com/orgs/${encodeURIComponent(org)}/repos?per_page=100&sort=pushed&direction=desc&page=${page}`, token)
    if (status === 404) { throw new Error(`No encuentro la organizacion "${org}" o tu cuenta no tiene acceso. Revisa el nombre exacto (login) de la org.`) }
    if (status === 403) { throw new Error('GitHub rechazo la peticion (permisos o limite). Conecta GitHub concediendo acceso a la organizacion.') }
    if (status === 401) { throw new Error('Sesion de GitHub invalida. Vuelve a conectar.') }
    if (!Array.isArray(json) || json.length === 0) { break }
    let stop = false
    for (const r of json) {
      const pushed = Date.parse(r.pushed_at || r.updated_at || '')
      if (Number.isFinite(pushed) && Number.isFinite(sinceMs) && pushed < sinceMs) { stop = true; break }
      out.push({ owner: r.owner?.login || org, name: r.name, full: r.full_name || `${org}/${r.name}` })
    }
    if (stop || json.length < 100) { break }
  }
  return out.slice(0, maxRepos)
}

/** Commits de un autor (email o usuario) en un repo desde `sinceISO`. */
async function repoCommits(token: string, owner: string, repo: string, author: string, sinceISO: string): Promise<Commit[]> {
  const url = `https://api.github.com/repos/${owner}/${repo}/commits?author=${encodeURIComponent(author)}&since=${sinceISO}T00:00:00Z&per_page=100`
  const { status, json } = await gh(url, token)
  if (status === 409 || status === 404) { return [] } // repo vacio o sin acceso
  if (!Array.isArray(json)) { return [] }
  return json.map((it: any) => ({
    repo: `${owner}/${repo}`,
    message: String(it.commit?.message || '').split('\n')[0].slice(0, 160),
    date: it.commit?.author?.date || it.commit?.committer?.date || '',
    sha: (it.sha || '').slice(0, 7),
    url: it.html_url || '',
  }))
}

/** Commits del equipo: recorre repos activos y filtra por cada autor. */
export async function getTeamCommits(token: string, org: string, authors: string[], sinceISO: string): Promise<TeamResult> {
  const repos = await listActiveRepos(token, org, sinceISO)
  const members: MemberResult[] = []
  for (const a of authors) {
    const commits: Commit[] = []
    let err = ''
    for (const r of repos) {
      try { commits.push(...await repoCommits(token, r.owner, r.name, a, sinceISO)) }
      catch (e: any) { err = e?.message || String(e) }
    }
    commits.sort((x, y) => String(y.date).localeCompare(String(x.date)))
    members.push({ author: a, commits, error: err || undefined })
  }
  return { repoCount: repos.length, members }
}
