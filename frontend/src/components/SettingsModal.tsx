import { useEffect, useState } from 'react'
import {
  clearAiSettings,
  EMPTY_AI_SETTINGS,
  isRemembered,
  loadAiSettings,
  maskKey,
  saveAiSettings,
  type AiSettings,
} from '../aiSettings'
import { api, type ServerDefaults, type TestConnectionResult } from '../api'
import { MAX_SETTINGS_FILE_BYTES, parseSettingsFile } from '../settingsFile'

type TestOutcome = TestConnectionResult | { ok: false; error: string }

interface Props {
  onClose: () => void
  authorName: string
  onAuthorNameChange: (name: string) => void
}

export default function SettingsModal({ onClose, authorName, onAuthorNameChange }: Props) {
  const [nameInput, setNameInput] = useState(authorName)
  const [nameSaved, setNameSaved] = useState(false)

  const [saved, setSaved] = useState<AiSettings>(loadAiSettings)
  const [form, setForm] = useState<AiSettings>(loadAiSettings)
  const [remember, setRemember] = useState(isRemembered)
  const [savedJustNow, setSavedJustNow] = useState(false)
  const [server, setServer] = useState<ServerDefaults | null>(null)
  const [importNote, setImportNote] = useState<{ ok: boolean; text: string } | null>(null)
  const [testBusy, setTestBusy] = useState(false)
  const [testResult, setTestResult] = useState<TestOutcome | null>(null)

  useEffect(() => {
    api.getServerDefaults().then(setServer).catch(() => setServer(null))
  }, [])

  const set = <K extends keyof AiSettings>(key: K, value: AiSettings[K]) => setForm((f) => ({ ...f, [key]: value }))
  const dirty = JSON.stringify(form) !== JSON.stringify(saved) || remember !== isRemembered()
  const anySaved = Object.entries(saved).some(([k, v]) => v && !(k === 'authMode' && v === 'api_key'))

  function save() {
    const trimmed: AiSettings = {
      ...form,
      apiKey: form.apiKey.trim(),
      baseUrl: form.baseUrl.trim(),
      extraHeaders: form.extraHeaders.trim(),
      model: form.model.trim(),
    }
    saveAiSettings(trimmed, remember)
    setSaved(trimmed)
    setForm(trimmed)
    setSavedJustNow(true)
    setTimeout(() => setSavedJustNow(false), 2500)
  }

  /** Read a settings file the user picked on THIS device. It is parsed right
   * here in the browser and only fills the form below -- never uploaded, and
   * not saved until the user presses "Save on this device". */
  async function importFromFile(file: File | undefined) {
    if (!file) return
    setImportNote(null)
    try {
      if (file.size > MAX_SETTINGS_FILE_BYTES) throw new Error('That file is too large to be a settings file.')
      const { settings, found, warnings } = parseSettingsFile(await file.text())
      setForm((f) => ({ ...f, ...settings }))
      setTestResult(null)
      setImportNote({
        ok: true,
        text:
          `Loaded ${found.join(', ')} from ${file.name}. Review below, then press “Save on this device”. ` +
          'The file stayed on this device.' +
          (warnings.length ? ' ' + warnings.join(' ') : ''),
      })
    } catch (e) {
      setImportNote({ ok: false, text: (e as Error).message })
    }
  }

  function forgetAll() {
    clearAiSettings()
    setSaved(EMPTY_AI_SETTINGS)
    setForm(EMPTY_AI_SETTINGS)
    setTestResult(null)
  }

  async function testConnection() {
    setTestBusy(true)
    setTestResult(null)
    try {
      // Tests the form's CURRENT values, saved or not.
      setTestResult(await api.testAnthropicConnection(form))
    } catch (e) {
      setTestResult({ ok: false, error: (e as Error).message })
    } finally {
      setTestBusy(false)
    }
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h3>Settings</h3>
          <button className="modal-close" onClick={onClose}>
            ×
          </button>
        </div>

        <label className="modal-label">Your name</label>
        <div className="modal-row">
          <input
            className="modal-input"
            placeholder="Your name"
            value={nameInput}
            onChange={(e) => setNameInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && nameInput.trim()) {
                onAuthorNameChange(nameInput.trim())
                setNameSaved(true)
                setTimeout(() => setNameSaved(false), 2500)
              }
            }}
          />
          <button
            disabled={!nameInput.trim()}
            onClick={() => {
              onAuthorNameChange(nameInput.trim())
              setNameSaved(true)
              setTimeout(() => setNameSaved(false), 2500)
            }}
          >
            Save
          </button>
        </div>
        {nameSaved && <p className="modal-success">Saved.</p>}

        <hr className="modal-divider" />

        <label className="modal-label">AI connection</label>
        <div className="privacy-note">
          🔒 Stored <strong>only on this device</strong> (in this browser). The server never saves your key,
          base URL or headers: they are sent with each AI request, used once, and discarded.
        </div>

        {server?.serverKeyConfigured && (
          <p className="modal-hint">
            This server also has its own key set in its environment; it is used only for fields you leave
            blank here.
          </p>
        )}

        <div className="modal-row">
          <label className="file-import-button">
            Import from file…
            <input
              type="file"
              hidden
              onChange={(e) => {
                importFromFile(e.target.files?.[0])
                e.target.value = '' // allow re-picking the same file
              }}
            />
          </label>
        </div>
        <p className="modal-hint">
          Pick a small settings file from this device (a .env or .json with ANTHROPIC_API_KEY,
          ANTHROPIC_BASE_URL, …). It is read in your browser only: never uploaded, and not saved until you
          press Save.
        </p>
        {importNote && (
          <p className={importNote.ok ? 'modal-success' : 'error-text'}>{importNote.text}</p>
        )}

        <label className="modal-sublabel">API key / auth token</label>
        <input
          type="password"
          className="modal-input modal-full"
          placeholder="sk-ant-... or your gateway token"
          autoComplete="off"
          value={form.apiKey}
          onChange={(e) => set('apiKey', e.target.value)}
        />
        {saved.apiKey && (
          <p className="modal-hint">
            Saved on this device: <code>{maskKey(saved.apiKey)}</code>
          </p>
        )}

        <label className="modal-sublabel">API base URL</label>
        <input
          className="modal-input modal-full"
          placeholder={server?.baseUrl || 'https://your-gateway.example.com  (blank = api.anthropic.com)'}
          value={form.baseUrl}
          onChange={(e) => set('baseUrl', e.target.value)}
        />
        <p className="modal-hint">
          Host plus any path prefix your gateway uses, without the trailing /v1 — it is added automatically.
        </p>

        <label className="modal-sublabel">Authentication</label>
        <select
          className="modal-input modal-select"
          value={form.authMode}
          onChange={(e) => set('authMode', e.target.value as AiSettings['authMode'])}
        >
          <option value="api_key">API key (x-api-key header) — api.anthropic.com</option>
          <option value="bearer">Bearer token (Authorization header) — most work gateways</option>
          <option value="both">Both headers</option>
        </select>
        <p className="modal-hint">If Test Connection reports HTTP 401, try the other option.</p>

        <label className="modal-sublabel">Extra headers (optional)</label>
        <textarea
          className="modal-input modal-textarea"
          rows={3}
          placeholder={'X-Tenant-Id: my-team\nOcp-Apim-Subscription-Key: ...'}
          value={form.extraHeaders}
          onChange={(e) => set('extraHeaders', e.target.value)}
        />
        <p className="modal-hint">One “Name: value” per line, for gateways that need more than a token.</p>

        <label className="modal-sublabel">Model</label>
        <input
          className="modal-input modal-full"
          placeholder={server?.defaultModel || 'claude-sonnet-5'}
          value={form.model}
          onChange={(e) => set('model', e.target.value)}
        />
        <p className="modal-hint">
          The model id your gateway serves. Bedrock-style gateways use ids like
          anthropic.claude-sonnet-4-5-20250929-v1:0.
        </p>

        <label className="modal-checkbox">
          <input type="checkbox" checked={remember} onChange={(e) => setRemember(e.target.checked)} />
          Remember on this device (untick to forget when this tab closes)
        </label>

        <div className="modal-row">
          <button className="primary" disabled={!dirty} onClick={save}>
            Save on this device
          </button>
          {anySaved && <button onClick={forgetAll}>Forget all AI settings</button>}
        </div>
        {savedJustNow && <p className="modal-success">Saved on this device.</p>}

        <hr className="modal-divider" />

        <label className="modal-label">Test connection</label>
        <button
          disabled={testBusy}
          onClick={testConnection}
          title="Sends one minimal real request using the values above (saved or not)"
        >
          {testBusy ? 'Testing…' : 'Test Connection'}
        </button>

        {testResult && testResult.ok && (
          <p className="modal-success">
            ✓ Connected — model <code>{testResult.model}</code>
            {testResult.baseUrl ? (
              <>
                {' '}
                via <code>{testResult.baseUrl}</code>
              </>
            ) : null}{' '}
            ({testResult.latencyMs}ms)
          </p>
        )}
        {testResult && !testResult.ok && <p className="error-text">{testResult.error}</p>}
      </div>
    </div>
  )
}
