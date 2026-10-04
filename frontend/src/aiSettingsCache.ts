/** Browser-side cache of the AI connection settings (API key, base URL,
 * model, auth mode, extra headers) entered into the Settings modal, so they survive a page reload or
 * a backend restart without retyping -- the backend itself only holds these
 * in process memory (see backend/app/api/settings.py) and forgets them the
 * moment it restarts. Wrapped in try/catch since localStorage can throw
 * (private browsing, disabled storage) and this is a convenience, not
 * something that should ever break the app. */

const KEYS = {
  apiKey: 'pptx_dev_anthropic_api_key',
  baseUrl: 'pptx_dev_anthropic_base_url',
  model: 'pptx_dev_ai_model',
  authMode: 'pptx_dev_auth_mode',
  extraHeaders: 'pptx_dev_extra_headers',
}

export interface CachedAiSettings {
  apiKey: string
  baseUrl: string
  model: string
  authMode: string
  extraHeaders: string
}

export function loadCachedAiSettings(): CachedAiSettings {
  try {
    return {
      apiKey: localStorage.getItem(KEYS.apiKey) || '',
      baseUrl: localStorage.getItem(KEYS.baseUrl) || '',
      model: localStorage.getItem(KEYS.model) || '',
      authMode: localStorage.getItem(KEYS.authMode) || '',
      extraHeaders: localStorage.getItem(KEYS.extraHeaders) || '',
    }
  } catch {
    return { apiKey: '', baseUrl: '', model: '', authMode: '', extraHeaders: '' }
  }
}

function setOrRemove(key: string, value: string) {
  try {
    if (value) localStorage.setItem(key, value)
    else localStorage.removeItem(key)
  } catch {
    // ignore -- caching is best-effort
  }
}

export const cacheApiKey = (value: string) => setOrRemove(KEYS.apiKey, value)
export const cacheBaseUrl = (value: string) => setOrRemove(KEYS.baseUrl, value)
export const cacheModel = (value: string) => setOrRemove(KEYS.model, value)
export const cacheAuthMode = (value: string) => setOrRemove(KEYS.authMode, value)
export const cacheExtraHeaders = (value: string) => setOrRemove(KEYS.extraHeaders, value)
