/** The user's AI connection settings (API key / work token, base URL, auth
 * mode, extra headers, model).
 *
 * These live ONLY in this browser, on this device -- the server stores none
 * of them. They are attached, as `x-ai-*` request headers, to the handful of
 * requests that actually call the AI (see `aiRequestHeaders` and api.ts), used
 * for that one request, and discarded by the server.
 *
 * "Remember on this device" (default on) keeps them in localStorage so they
 * survive closing the browser; turned off they go in sessionStorage and are
 * forgotten when the tab closes. Everything is wrapped in try/catch: storage
 * can throw (private browsing, blocked site data) and the app must still work.
 */

export type AuthMode = 'api_key' | 'bearer' | 'both'

export interface AiSettings {
  apiKey: string
  baseUrl: string
  authMode: AuthMode
  /** One "Name: value" per line, or a JSON object. */
  extraHeaders: string
  model: string
}

export const EMPTY_AI_SETTINGS: AiSettings = {
  apiKey: '',
  baseUrl: '',
  authMode: 'api_key',
  extraHeaders: '',
  model: '',
}

const STORAGE_KEY = 'pptx_dev_ai_settings'
const REMEMBER_KEY = 'pptx_dev_ai_remember'
// Keys written by earlier versions, which also cached these in localStorage.
const LEGACY_KEYS = {
  apiKey: 'pptx_dev_anthropic_api_key',
  baseUrl: 'pptx_dev_anthropic_base_url',
  model: 'pptx_dev_ai_model',
  authMode: 'pptx_dev_auth_mode',
  extraHeaders: 'pptx_dev_extra_headers',
}

function safe<T>(fn: () => T, fallback: T): T {
  try {
    return fn()
  } catch {
    return fallback
  }
}

function isAuthMode(v: unknown): v is AuthMode {
  return v === 'api_key' || v === 'bearer' || v === 'both'
}

function normalize(raw: Partial<AiSettings> | null): AiSettings {
  const r = raw ?? {}
  return {
    apiKey: typeof r.apiKey === 'string' ? r.apiKey : '',
    baseUrl: typeof r.baseUrl === 'string' ? r.baseUrl : '',
    authMode: isAuthMode(r.authMode) ? r.authMode : 'api_key',
    extraHeaders: typeof r.extraHeaders === 'string' ? r.extraHeaders : '',
    model: typeof r.model === 'string' ? r.model : '',
  }
}

/** Whether settings are kept across browser restarts (default: yes). */
export function isRemembered(): boolean {
  return safe(() => localStorage.getItem(REMEMBER_KEY) !== '0', true)
}

/** One-time move of values an earlier version cached under separate keys. */
function migrateLegacy(): void {
  safe(() => {
    const legacy: Partial<AiSettings> = {
      apiKey: localStorage.getItem(LEGACY_KEYS.apiKey) || '',
      baseUrl: localStorage.getItem(LEGACY_KEYS.baseUrl) || '',
      model: localStorage.getItem(LEGACY_KEYS.model) || '',
      authMode: (localStorage.getItem(LEGACY_KEYS.authMode) as AuthMode) || 'api_key',
      extraHeaders: localStorage.getItem(LEGACY_KEYS.extraHeaders) || '',
    }
    const any = Object.values(LEGACY_KEYS).some((k) => localStorage.getItem(k) !== null)
    if (!any) return
    if (!localStorage.getItem(STORAGE_KEY)) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(normalize(legacy)))
    }
    for (const k of Object.values(LEGACY_KEYS)) localStorage.removeItem(k)
  }, undefined)
}

export function loadAiSettings(): AiSettings {
  migrateLegacy()
  const fromSession = safe(() => sessionStorage.getItem(STORAGE_KEY), null)
  const fromLocal = safe(() => localStorage.getItem(STORAGE_KEY), null)
  const raw = fromSession ?? fromLocal
  if (!raw) return { ...EMPTY_AI_SETTINGS }
  return normalize(safe(() => JSON.parse(raw), null))
}

export function saveAiSettings(settings: AiSettings, remember: boolean): void {
  const json = JSON.stringify(normalize(settings))
  safe(() => localStorage.setItem(REMEMBER_KEY, remember ? '1' : '0'), undefined)
  // Write to exactly one place, and remove it from the other.
  const [keep, drop] = remember ? [localStorage, sessionStorage] : [sessionStorage, localStorage]
  safe(() => keep.setItem(STORAGE_KEY, json), undefined)
  safe(() => drop.removeItem(STORAGE_KEY), undefined)
}

export function clearAiSettings(): void {
  safe(() => localStorage.removeItem(STORAGE_KEY), undefined)
  safe(() => sessionStorage.removeItem(STORAGE_KEY), undefined)
}

/** UTF-8 safe base64 (btoa alone throws on non-Latin-1). */
function b64(text: string): string {
  const bytes = new TextEncoder().encode(text)
  let bin = ''
  for (const b of bytes) bin += String.fromCharCode(b)
  return btoa(bin)
}

/** The `x-ai-*` headers carrying these settings for ONE request. Empty fields
 * are omitted so the server falls back to its own (environment) defaults. */
export function aiHeadersFor(s: AiSettings): Record<string, string> {
  const h: Record<string, string> = {}
  if (s.apiKey.trim()) h['x-ai-key'] = s.apiKey.trim()
  if (s.baseUrl.trim()) h['x-ai-base-url'] = s.baseUrl.trim()
  if (s.authMode !== 'api_key') h['x-ai-auth-mode'] = s.authMode
  if (s.extraHeaders.trim()) h['x-ai-extra-headers'] = b64(s.extraHeaders.trim())
  if (s.model.trim()) h['x-ai-model'] = s.model.trim()
  return h
}

export function aiRequestHeaders(): Record<string, string> {
  return aiHeadersFor(loadAiSettings())
}

/** First 6 and last 4 characters, for showing that a key is saved. */
export function maskKey(key: string): string {
  if (!key) return ''
  return key.length <= 10 ? '•'.repeat(key.length) : `${key.slice(0, 6)}…${key.slice(-4)}`
}
