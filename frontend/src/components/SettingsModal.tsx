import { useEffect, useState } from 'react'
import { cacheApiKey, cacheBaseUrl, cacheModel } from '../aiSettingsCache'
import { api, type ApiKeyStatus, type TestConnectionResult } from '../api'

type TestOutcome = TestConnectionResult | { ok: false; error: string }

interface Props {
  onClose: () => void
  authorName: string
  onAuthorNameChange: (name: string) => void
}

export default function SettingsModal({ onClose, authorName, onAuthorNameChange }: Props) {
  const [status, setStatus] = useState<ApiKeyStatus | null>(null)
  const [input, setInput] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [savedJustNow, setSavedJustNow] = useState(false)
  const [nameInput, setNameInput] = useState(authorName)
  const [nameSaved, setNameSaved] = useState(false)
  const [baseUrlInput, setBaseUrlInput] = useState('')
  const [baseUrlBusy, setBaseUrlBusy] = useState(false)
  const [baseUrlError, setBaseUrlError] = useState<string | null>(null)
  const [baseUrlSavedJustNow, setBaseUrlSavedJustNow] = useState(false)
  const [modelInput, setModelInput] = useState('')
  const [modelBusy, setModelBusy] = useState(false)
  const [modelError, setModelError] = useState<string | null>(null)
  const [modelSavedJustNow, setModelSavedJustNow] = useState(false)
  const [testBusy, setTestBusy] = useState(false)
  const [testResult, setTestResult] = useState<TestOutcome | null>(null)

  useEffect(() => {
    api
      .getApiKeyStatus()
      .then((s) => {
        setStatus(s)
        setBaseUrlInput(s.baseUrl || '')
        setModelInput(s.model)
      })
      .catch((e) => setError((e as Error).message))
  }, [])

  async function save() {
    setBusy(true)
    setError(null)
    try {
      const s = await api.setApiKey(input)
      setStatus(s)
      cacheApiKey(input.trim())
      setInput('')
      setSavedJustNow(true)
      setTimeout(() => setSavedJustNow(false), 2500)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  async function clear() {
    setBusy(true)
    setError(null)
    try {
      const s = await api.setApiKey('')
      setStatus(s)
      cacheApiKey('')
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  async function saveBaseUrl() {
    setBaseUrlBusy(true)
    setBaseUrlError(null)
    try {
      const s = await api.setBaseUrl(baseUrlInput.trim())
      setStatus(s)
      cacheBaseUrl(s.baseUrl || '')
      setBaseUrlSavedJustNow(true)
      setTimeout(() => setBaseUrlSavedJustNow(false), 2500)
    } catch (e) {
      setBaseUrlError((e as Error).message)
    } finally {
      setBaseUrlBusy(false)
    }
  }

  async function clearBaseUrl() {
    setBaseUrlBusy(true)
    setBaseUrlError(null)
    try {
      const s = await api.setBaseUrl('')
      setStatus(s)
      setBaseUrlInput('')
      cacheBaseUrl('')
    } catch (e) {
      setBaseUrlError((e as Error).message)
    } finally {
      setBaseUrlBusy(false)
    }
  }

  async function saveModel() {
    setModelBusy(true)
    setModelError(null)
    try {
      const s = await api.setModel(modelInput.trim())
      setStatus(s)
      setModelInput(s.model)
      cacheModel(s.modelSource === 'runtime' ? s.model : '')
      setModelSavedJustNow(true)
      setTimeout(() => setModelSavedJustNow(false), 2500)
    } catch (e) {
      setModelError((e as Error).message)
    } finally {
      setModelBusy(false)
    }
  }

  async function resetModel() {
    setModelBusy(true)
    setModelError(null)
    try {
      const s = await api.setModel('')
      setStatus(s)
      setModelInput(s.model)
      cacheModel('')
    } catch (e) {
      setModelError((e as Error).message)
    } finally {
      setModelBusy(false)
    }
  }

  async function testConnection() {
    setTestBusy(true)
    setTestResult(null)
    try {
      const result = await api.testAnthropicConnection()
      setTestResult(result)
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

        <label className="modal-label">Anthropic API key</label>
        <p className="modal-hint">Cached in this browser (not encrypted).</p>

        {status && (
          <div className={`key-status ${status.configured ? 'key-status-ok' : 'key-status-empty'}`}>
            {status.configured ? (
              <>
                Configured ({status.source === 'runtime' ? 'set here' : 'from environment'}):{' '}
                <code>{status.masked}</code>
              </>
            ) : (
              'No API key configured — "Ask AI" will fail until one is set.'
            )}
          </div>
        )}

        <div className="modal-row">
          <input
            type="password"
            className="modal-input"
            placeholder="sk-ant-..."
            title="Used for &quot;Ask AI&quot;. For a real deployment, set ANTHROPIC_API_KEY via your platform's secrets manager instead."
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && input.trim()) save()
            }}
            disabled={busy}
          />
          <button disabled={busy || !input.trim()} onClick={save}>
            Save
          </button>
        </div>

        {savedJustNow && <p className="modal-success">Saved.</p>}
        {error && <p className="error-text">{error}</p>}

        {status?.configured && (
          <button className="modal-clear" disabled={busy} onClick={clear}>
            Clear key
          </button>
        )}

        <hr className="modal-divider" />

        <label className="modal-label">API base URL</label>
        <p className="modal-hint">For a Bedrock/gateway proxy instead of the public API. No /v1 suffix.</p>

        {status && (
          <div className={`key-status ${status.baseUrl ? 'key-status-ok' : 'key-status-empty'}`}>
            {status.baseUrl ? (
              <>
                Using ({status.baseUrlSource === 'runtime' ? 'set here' : 'from environment'}):{' '}
                <code>{status.baseUrl}</code>
              </>
            ) : (
              'Using the default Anthropic API endpoint.'
            )}
          </div>
        )}

        <div className="modal-row">
          <input
            className="modal-input"
            placeholder="https://your-bedrock-gateway.example.mil"
            title="Leave blank for the default (api.anthropic.com). Just the host — no /v1 suffix, it's added automatically."
            value={baseUrlInput}
            onChange={(e) => setBaseUrlInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') saveBaseUrl()
            }}
            disabled={baseUrlBusy}
          />
          <button disabled={baseUrlBusy} onClick={saveBaseUrl}>
            Save
          </button>
        </div>

        {baseUrlSavedJustNow && <p className="modal-success">Saved.</p>}
        {baseUrlError && <p className="error-text">{baseUrlError}</p>}

        {status?.baseUrlSource === 'runtime' && (
          <button className="modal-clear" disabled={baseUrlBusy} onClick={clearBaseUrl}>
            Reset to default
          </button>
        )}

        <hr className="modal-divider" />

        <label className="modal-label">Model</label>
        <p className="modal-hint">Bedrock gateways use different model IDs than the direct API.</p>

        {status && (
          <div className={`key-status ${status.modelSource === 'runtime' ? 'key-status-ok' : 'key-status-empty'}`}>
            Using ({status.modelSource === 'runtime' ? 'set here' : 'from environment/default'}):{' '}
            <code>{status.model}</code>
          </div>
        )}

        <div className="modal-row">
          <input
            className="modal-input"
            placeholder={status?.defaultModel || 'claude-sonnet-5'}
            title='e.g. anthropic.claude-sonnet-4-5-20250929-v1:0 for a Bedrock gateway'
            value={modelInput}
            onChange={(e) => setModelInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && modelInput.trim()) saveModel()
            }}
            disabled={modelBusy}
          />
          <button disabled={modelBusy || !modelInput.trim()} onClick={saveModel}>
            Save
          </button>
        </div>

        {modelSavedJustNow && <p className="modal-success">Saved.</p>}
        {modelError && <p className="error-text">{modelError}</p>}

        {status?.modelSource === 'runtime' && (
          <button className="modal-clear" disabled={modelBusy} onClick={resetModel}>
            Reset to default
          </button>
        )}

        <hr className="modal-divider" />

        <label className="modal-label">Test connection</label>
        <button
          disabled={testBusy}
          onClick={testConnection}
          title="Sends one minimal real request using the settings above (save changes first)"
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
