import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { CommentEntry, Slide, SlideSize, ViewMode } from '../types'
import OutlineView from './OutlineView'
import SelectionPopup, { type SelectionRange } from './SelectionPopup'
import SlideCanvas from './SlideCanvas'

interface Props {
  docId: string
  slides: Slide[]
  slideSize: SlideSize
  comments: Record<string, CommentEntry>
  viewMode: ViewMode
  /** A slide id or shape id to scroll to. */
  scrollToId: string | null
  onScrolled: () => void
  onJumpToSlide: (slideId: string) => void
  selectedCommentId: string | null
  onSelectComment: (id: string) => void
  onReplaceRange: (
    startRunId: string,
    startOffset: number,
    endRunId: string,
    endOffset: number,
    newText: string,
  ) => Promise<void>
  onAddComment: (
    startRunId: string,
    startOffset: number,
    endRunId: string,
    endOffset: number,
    text: string,
  ) => Promise<void>
  onAddSlideComment: (slideId: string, text: string) => Promise<void>
}

function resolveRunPosition(node: Node, offset: number): { runId: string; charOffset: number } | null {
  if (node.nodeType === Node.TEXT_NODE) {
    const el = (node.parentElement as Element | null)?.closest('[data-run-id]')
    if (!el) return null
    return { runId: el.getAttribute('data-run-id')!, charOffset: offset }
  }
  if (node.nodeType === Node.ELEMENT_NODE) {
    const runEl = (node as Element).closest('[data-run-id]')
    if (runEl) {
      const textLen = (runEl.textContent || '').length
      return { runId: runEl.getAttribute('data-run-id')!, charOffset: offset === 0 ? 0 : textLen }
    }
  }
  return null
}

function SlideCommentButton({ slideId, onAdd }: { slideId: string; onAdd: Props['onAddSlideComment'] }) {
  const [open, setOpen] = useState(false)
  const [text, setText] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  async function submit() {
    if (!text.trim()) return
    setBusy(true)
    setError(null)
    try {
      await onAdd(slideId, text.trim())
      setText('')
      setOpen(false)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  if (!open) {
    return (
      <button className="slide-header-button" onClick={() => setOpen(true)} title="Add a comment on this whole slide">
        💬 Comment on slide
      </button>
    )
  }
  return (
    <div className="slide-comment-form">
      <textarea
        autoFocus
        placeholder="Comment on this slide…"
        value={text}
        disabled={busy}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Escape') setOpen(false)
          if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submit()
        }}
      />
      <div className="slide-comment-form-actions">
        <button onClick={() => setOpen(false)} disabled={busy}>
          Cancel
        </button>
        <button className="primary" onClick={submit} disabled={busy || !text.trim()}>
          Add Comment
        </button>
      </div>
      {error && <p className="error-text">{error}</p>}
    </div>
  )
}

export default function DocumentViewer({
  docId,
  slides,
  slideSize,
  comments,
  viewMode,
  scrollToId,
  onScrolled,
  onJumpToSlide,
  selectedCommentId,
  onSelectComment,
  onReplaceRange,
  onAddComment,
  onAddSlideComment,
}: Props) {
  const rootRef = useRef<HTMLDivElement>(null)
  const [selection, setSelection] = useState<SelectionRange | null>(null)
  const [notesOpen, setNotesOpen] = useState<Record<string, boolean>>({})

  useEffect(() => {
    if (viewMode !== 'slides') return

    let debounceTimer: ReturnType<typeof setTimeout> | null = null

    function processSelection() {
      // Focus inside our own popup (typing in the textarea, or the moment a
      // button in it is clicked) -- never let a resulting selectionchange
      // (e.g. an input's own internal selection) clear the popup out from
      // under the user.
      if ((document.activeElement as Element | null)?.closest?.('.selection-popup')) return

      const sel = window.getSelection()
      if (!sel || sel.isCollapsed || sel.rangeCount === 0) {
        setSelection(null)
        return
      }
      const range = sel.getRangeAt(0)
      if (!rootRef.current?.contains(range.commonAncestorContainer)) {
        return
      }

      const start = resolveRunPosition(range.startContainer, range.startOffset)
      const end = resolveRunPosition(range.endContainer, range.endOffset)
      const text = sel.toString()
      if (!start || !end || !text) {
        setSelection(null)
        return
      }

      const rect = range.getBoundingClientRect()
      setSelection({
        startRunId: start.runId,
        startOffset: start.charOffset,
        endRunId: end.runId,
        endOffset: end.charOffset,
        text,
        rect: { top: rect.top, left: rect.left, bottom: rect.bottom, right: rect.right },
      })
    }

    function handleSelectionChange() {
      // selectionchange fires on every intermediate step of a drag (mouse
      // drag on desktop, or dragging the native selection handles on
      // mobile, which don't otherwise emit mouseup/touchend on our DOM at
      // all) -- debounce so we act once the selection settles rather than
      // on every frame, and so it works identically for touch selection.
      if (debounceTimer) clearTimeout(debounceTimer)
      debounceTimer = setTimeout(processSelection, 150)
    }

    function handleScroll() {
      setSelection(null)
    }

    document.addEventListener('selectionchange', handleSelectionChange)
    window.addEventListener('scroll', handleScroll, true)
    return () => {
      if (debounceTimer) clearTimeout(debounceTimer)
      document.removeEventListener('selectionchange', handleSelectionChange)
      window.removeEventListener('scroll', handleScroll, true)
    }
  }, [viewMode])

  useLayoutEffect(() => {
    if (!scrollToId || viewMode !== 'slides') return
    const el =
      rootRef.current?.querySelector(`[data-shape-id="${scrollToId}"]`) ??
      rootRef.current?.querySelector(`.slide-frame[data-slide-id="${scrollToId}"]`)
    el?.scrollIntoView({ behavior: 'smooth', block: 'center' })
    onScrolled()
  }, [scrollToId, viewMode, onScrolled])

  if (viewMode === 'outline') {
    return (
      <div className="document-viewer view-outline" ref={rootRef}>
        <OutlineView slides={slides} onJumpToSlide={onJumpToSlide} />
      </div>
    )
  }

  if (viewMode === 'grid') {
    return (
      <div className="document-viewer view-grid" ref={rootRef}>
        <div className="slide-grid">
          {slides.map((slide) => {
            const open = Object.values(comments).filter(
              (c) => c.slideId === slide.id && !c.parentId && c.kind !== 'edit' && !c.done,
            ).length
            return (
              <button key={slide.id} className="slide-grid-item" onClick={() => onJumpToSlide(slide.id)}>
                <SlideCanvas
                  docId={docId}
                  slide={slide}
                  slideSize={slideSize}
                  comments={comments}
                  selectedCommentId={null}
                  onSelectComment={() => {}}
                  inert
                />
                <span className="slide-grid-caption">
                  <span className="slide-grid-number">{slide.index}</span>
                  <span className="slide-grid-title">{slide.title || '(untitled)'}</span>
                  {open > 0 && <span className="slide-grid-badge">💬 {open}</span>}
                </span>
              </button>
            )
          })}
        </div>
      </div>
    )
  }

  return (
    <div className="document-viewer view-slides" ref={rootRef}>
      {slides.map((slide) => (
        <section className="slide-frame" data-slide-id={slide.id} key={slide.id}>
          <header className="slide-header">
            <span className="slide-header-title">
              <span className="slide-header-number">Slide {slide.index}</span>
              {slide.title && <span className="slide-header-name">{slide.title}</span>}
              {slide.hidden && <span className="outline-hidden-badge">hidden</span>}
            </span>
            <SlideCommentButton slideId={slide.id} onAdd={onAddSlideComment} />
          </header>
          <SlideCanvas
            docId={docId}
            slide={slide}
            slideSize={slideSize}
            comments={comments}
            selectedCommentId={selectedCommentId}
            onSelectComment={onSelectComment}
          />
          {slide.notes && (
            <div className="slide-notes">
              <button
                className="slide-notes-toggle"
                onClick={() => setNotesOpen((n) => ({ ...n, [slide.id]: !n[slide.id] }))}
              >
                {notesOpen[slide.id] ? '▾' : '▸'} Speaker notes
              </button>
              {notesOpen[slide.id] && <p className="slide-notes-text">{slide.notes}</p>}
            </div>
          )}
        </section>
      ))}

      {selection && (
        <SelectionPopup
          selection={selection}
          onClose={() => setSelection(null)}
          onAddComment={(text) =>
            onAddComment(selection.startRunId, selection.startOffset, selection.endRunId, selection.endOffset, text)
          }
          onChangeText={(newText) =>
            onReplaceRange(
              selection.startRunId,
              selection.startOffset,
              selection.endRunId,
              selection.endOffset,
              newText,
            )
          }
        />
      )}
    </div>
  )
}
