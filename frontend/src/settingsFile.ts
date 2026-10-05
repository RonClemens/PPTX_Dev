/** Reads AI connection settings out of a small file the user picks on their
 * own device ("Import from file" in Settings).
 *
 * This runs entirely in the browser: the file's text is parsed here and only
 * ever fills in the Settings form. It is never uploaded anywhere, and nothing
 * is saved until the user reviews it and presses "Save on this device".
 *
 * Accepted formats (all optional fields, unknown keys ignored):
 *   - a .env-style file:  ANTHROPIC_API_KEY=..., ANTHROPIC_BASE_URL=..., ...
 *   - a JSON object using the same names, or camelCase names (apiKey, baseUrl,
 *     authMode, extraHeaders, model), or Claude Code's {"env": { ... }} shape.
 */

import type { AiSettings, AuthMode } from './aiSettings.ts'

export const MAX_SETTINGS_FILE_BYTES = 100 * 1024

export interface ImportedSettings {
  settings: Partial<AiSettings>
  /** Human-readable names of the fields that were found. */
  found: string[]
  warnings: string[]
}

const FIELD_LABELS: Record<keyof AiSettings, string> = {
  apiKey: 'API key / token',
  baseUrl: 'base URL',
  authMode: 'authentication mode',
  extraHeaders: 'extra headers',
  model: 'model',
}

function parseAuthMode(value: string): AuthMode | null {
  const v = value.trim().toLowerCase().replace(/-/g, '_')
  if (['api_key', 'x_api_key', 'apikey', 'default'].includes(v)) return 'api_key'
  if (['bearer', 'token', 'auth_token', 'authorization'].includes(v)) return 'bearer'
  if (v === 'both') return 'both'
  return null
}

function parseEnvText(text: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const rawLine of text.replace(/^﻿/, '').split(/\r?\n/)) {
    let line = rawLine.trim()
    if (!line || line.startsWith('#')) continue
    if (line.startsWith('export ')) line = line.slice(7).trim()
    const eq = line.indexOf('=')
    if (eq <= 0) continue
    const key = line.slice(0, eq).trim()
    let value = line.slice(eq + 1).trim()
    const quote = value[0]
    if ((quote === '"' || quote === "'") && value.length >= 2 && value.endsWith(quote)) {
      value = value.slice(1, -1)
    } else {
      value = value.replace(/\s+#.*$/, '') // inline comment on an unquoted value
    }
    out[key] = value
  }
  return out
}

function flattenJson(obj: unknown): Record<string, string> {
  const out: Record<string, string> = {}
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return out
  const rec = obj as Record<string, unknown>
  // Claude Code's settings.json keeps these under "env".
  const source = rec.env && typeof rec.env === 'object' && !Array.isArray(rec.env) ? { ...rec, ...(rec.env as object) } : rec
  for (const [k, v] of Object.entries(source)) {
    if (typeof v === 'string' || typeof v === 'number') out[k] = String(v)
    else if (v && typeof v === 'object' && !Array.isArray(v) && /headers/i.test(k)) {
      out[k] = Object.entries(v as Record<string, unknown>)
        .map(([name, val]) => `${name}: ${String(val)}`)
        .join('\n')
    }
  }
  return out
}

/** Parse the text of a settings file. Throws an Error with a user-readable
 * message when it can't be read or contains nothing recognizable. */
export function parseSettingsFile(text: string): ImportedSettings {
  const trimmed = text.trim()
  if (!trimmed) throw new Error('That file is empty.')

  let raw: Record<string, string>
  if (trimmed.startsWith('{')) {
    try {
      raw = flattenJson(JSON.parse(trimmed))
    } catch {
      throw new Error('That file looks like JSON but could not be read as JSON.')
    }
  } else {
    raw = parseEnvText(trimmed)
  }

  // Case-insensitive, separator-insensitive lookup: API_KEY == apiKey == api-key.
  const norm = (k: string) => k.toLowerCase().replace(/[^a-z0-9]/g, '')
  const byName = new Map(Object.entries(raw).map(([k, v]) => [norm(k), v.trim()]))
  const pick = (...names: string[]) => {
    for (const n of names) {
      const v = byName.get(norm(n))
      if (v) return v
    }
    return ''
  }

  const settings: Partial<AiSettings> = {}
  const warnings: string[] = []

  const key = pick('ANTHROPIC_API_KEY', 'apiKey', 'api_key', 'ANTHROPIC_AUTH_TOKEN', 'authToken', 'token')
  const usedTokenAlias = !pick('ANTHROPIC_API_KEY', 'apiKey', 'api_key') && !!pick('ANTHROPIC_AUTH_TOKEN', 'authToken', 'token')
  if (key) settings.apiKey = key

  const baseUrl = pick('ANTHROPIC_BASE_URL', 'baseUrl', 'base_url')
  if (baseUrl) settings.baseUrl = baseUrl.replace(/\/+$/, '')

  const modeText = pick('PPTX_DEV_AUTH_MODE', 'authMode', 'auth_mode')
  if (modeText) {
    const mode = parseAuthMode(modeText)
    if (mode) settings.authMode = mode
    else warnings.push(`Ignored authentication mode "${modeText}" (use api_key, bearer or both).`)
  } else if (usedTokenAlias) {
    // ANTHROPIC_AUTH_TOKEN is what Claude Code uses for Bearer-token gateways.
    settings.authMode = 'bearer'
  }

  const headers = pick('PPTX_DEV_EXTRA_HEADERS', 'extraHeaders', 'extra_headers')
  if (headers) settings.extraHeaders = headers.replace(/\\n/g, '\n')

  const model = pick('PPTX_DEV_AI_MODEL', 'ANTHROPIC_MODEL', 'model')
  if (model) settings.model = model

  const found = (Object.keys(settings) as (keyof AiSettings)[]).map((k) => FIELD_LABELS[k])
  if (found.length === 0) {
    throw new Error(
      'No AI settings found in that file. Expected names like ANTHROPIC_API_KEY, ANTHROPIC_BASE_URL, ' +
        'PPTX_DEV_AUTH_MODE, PPTX_DEV_EXTRA_HEADERS and PPTX_DEV_AI_MODEL.',
    )
  }
  return { settings, found, warnings }
}
