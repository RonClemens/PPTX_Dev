import { useRef, useState } from 'react'
import type { DocumentMeta } from '../types'

interface Props {
  documents: DocumentMeta[]
  activeDocId: string | null
  onSelect: (docId: string) => void
  onUpload: (file: File) => Promise<void>
  onDelete: (docId: string) => Promise<void>
  onOpenSettings: () => void
}

const BUILD_TIME = new Date(__BUILD_TIME__)

export default function UploadPanel({
  documents,
  activeDocId,
  onSelect,
  onUpload,
  onDelete,
  onOpenSettings,
}: Props) {
  const fileInput = useRef<HTMLInputElement>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  async function handleFile(file: File | undefined) {
    if (!file) return
    setBusy(true)
    setError(null)
    try {
      await onUpload(file)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
      if (fileInput.current) fileInput.current.value = ''
    }
  }

  return (
    <div className="upload-panel">
      <div className="upload-panel-heading">
        <h2>PPTX Review Assistant</h2>
        <button className="settings-button" title="Settings" onClick={onOpenSettings}>
          ⚙
        </button>
      </div>
      <label className="upload-dropzone">
        <input
          ref={fileInput}
          type="file"
          accept=".pptx"
          onChange={(e) => handleFile(e.target.files?.[0])}
          disabled={busy}
        />
        {busy ? 'Uploading…' : 'Upload a .pptx file'}
      </label>
      {error && <p className="error-text">{error}</p>}

      <h3>Documents</h3>
      <ul className="document-list">
        {documents.map((d) => (
          <li
            key={d.id}
            className={d.id === activeDocId ? 'doc-item doc-item-active' : 'doc-item'}
          >
            <span onClick={() => onSelect(d.id)}>{d.title || d.filename}</span>
            <button className="doc-delete" title="Delete" onClick={() => onDelete(d.id)}>
              ×
            </button>
          </li>
        ))}
        {documents.length === 0 && <li className="empty-hint">No documents yet.</li>}
      </ul>

      <p className="deploy-timestamp" title={BUILD_TIME.toISOString()}>
        Deployed {BUILD_TIME.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' })}
      </p>
    </div>
  )
}
