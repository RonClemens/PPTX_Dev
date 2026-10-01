import { useLayoutEffect, useRef, useState } from 'react'

export interface SelectionRange {
  startRunId: string
  startOffset: number
  endRunId: string
  endOffset: number
  text: string
  rect: { top: number; left: number; bottom: number; right: number }
}

interface Props {
  selection: SelectionRange
  onAddComment: (text: string) => Promise<void>
  onChangeText: (newText: string) => Promise<void>
  onClose: () => void
}

type Mode = 'menu' | 'comment' | 'edit'

export default function SelectionPopup({ selection, onAddComment, onChangeText, onClose }: Props) {
  const [mode, setMode] = useState<Mode>('menu')
  const [commentText, setCommentText] = useState('')
  const [editText, setEditText] = useState(selection.text)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const popupRef = useRef<HTMLDivElement>(null)
  const left = Math.min(Math.max(selection.rect.left, 8), window.innerWidth - 280)
  const [top, setTop] = useState(Math.min(selection.rect.bottom + 8, window.innerHeight - 60))

  // The form (textarea + buttons) is much taller than the one-line menu, so
  // after every mode change measure the real height and keep the whole popup
  // on screen: below the selection if it fits, otherwise above it, otherwise
  // pinned to the bottom edge. Without this, a selection low in the viewport
  // pushed the form's buttons off-screen where they couldn't be clicked.
  useLayoutEffect(() => {
    const height = popupRef.current?.offsetHeight ?? 0
    const below = selection.rect.bottom + 8
    if (below + height <= window.innerHeight - 8) setTop(below)
    else if (selection.rect.top - height - 8 >= 8) setTop(selection.rect.top - height - 8)
    else setTop(Math.max(8, window.innerHeight - height - 8))
  }, [mode, selection.rect.bottom, selection.rect.top, error])

  async function submitComment() {
    if (!commentText.trim()) return
    setBusy(true)
    setError(null)
    try {
      await onAddComment(commentText.trim())
      onClose()
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  async function submitEdit() {
    if (!editText.trim()) return
    setBusy(true)
    setError(null)
    try {
      await onChangeText(editText)
      onClose()
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <div
      className="selection-popup"
      ref={popupRef}
      style={{ top, left }}
      onMouseDown={(e) => e.stopPropagation()}
    >
      {mode === 'menu' && (
        <div className="selection-popup-menu">
          <button onClick={() => setMode('comment')}>💬 Add Comment</button>
          <button onClick={() => setMode('edit')}>✏️ Change Text</button>
        </div>
      )}

      {mode === 'comment' && (
        <div className="selection-popup-form">
          <textarea
            autoFocus
            placeholder="Comment..."
            value={commentText}
            disabled={busy}
            onChange={(e) => setCommentText(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Escape') onClose()
              if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submitComment()
            }}
          />
          <div className="selection-popup-actions">
            <button onClick={onClose} disabled={busy}>
              Cancel
            </button>
            <button className="primary" onClick={submitComment} disabled={busy || !commentText.trim()}>
              Add Comment
            </button>
          </div>
        </div>
      )}

      {mode === 'edit' && (
        <div className="selection-popup-form">
          <textarea
            autoFocus
            value={editText}
            disabled={busy}
            onChange={(e) => setEditText(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Escape') onClose()
              if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submitEdit()
            }}
          />
          <div className="selection-popup-actions">
            <button onClick={onClose} disabled={busy}>
              Cancel
            </button>
            <button className="primary" onClick={submitEdit} disabled={busy || !editText.trim()}>
              Replace
            </button>
          </div>
        </div>
      )}

      {error && <p className="error-text selection-popup-error">{error}</p>}
    </div>
  )
}
