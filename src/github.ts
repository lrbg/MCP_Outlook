import * as vscode from 'vscode'

/**
 * Metricas de commits del equipo via GitHub. Reutiliza la sesion de GitHub de
 * VS Code (sin pedir token a mano). Busca commits por autor (email o usuario)
 * dentro de la organizacion con la Search Commits API.
 */

export interface Commit { repo: string; message: string; date: string; sha: string; url: string }

/** Token de la sesion de GitHub de VS Code. interactive=true puede abrir el login. */
export async function getGithubToken(interactive = false): Promise<string> {
  try {
    const s = await vscode.authentication.getSession('github', ['repo', 'read:org'], interactive ? { createIfNone: true } : { silent: true } as any)
    return s?.accessToken || ''
  } catch { return '' }
}

async function gh(url: string, token: string): Promise<any> {
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/vnd.github+json',
      'User-Agent': 'm365-mcp-plugin',
    },
  })
  if (res.status === 403) { throw new Error('GitHub rechazo la peticion (limite de tasa o permisos). Intenta en un minuto.') }
  if (res.status === 401) { throw new Error('Sesion de GitHub invalida. Vuelve a conectar GitHub.') }
  if (!res.ok) { throw new Error('GitHub HTTP ' + res.status) }
  return res.json()
}

/** Organizaciones del usuario (para elegir cual). */
export async function getOrgs(token: string): Promise<string[]> {
  const j = await gh('https://api.github.com/user/orgs?per_page=100', token)
  return Array.isArray(j) ? j.map((o: any) => o.login).filter(Boolean) : []
}

/** Commits de un autor (email o username) en la org desde `sinceISO` (YYYY-MM-DD). */
export async function searchCommits(token: string, org: string, author: string, sinceISO: string): Promise<Commit[]> {
  const isEmail = author.includes('@')
  const q = `org:${org} ${isEmail ? 'author-email' : 'author'}:${author} author-date:>=${sinceISO}`
  const url = `https://api.github.com/search/commits?q=${encodeURIComponent(q)}&per_page=100&sort=author-date&order=desc`
  const j = await gh(url, token)
  return (j.items || []).map((it: any) => ({
    repo: it.repository?.full_name || it.repository?.name || '',
    message: String(it.commit?.message || '').split('\n')[0].slice(0, 160),
    date: it.commit?.author?.date || it.commit?.committer?.date || '',
    sha: (it.sha || '').slice(0, 7),
    url: it.html_url || '',
  }))
}
