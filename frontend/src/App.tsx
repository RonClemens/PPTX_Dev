import { useEffect, useState } from 'react'
import './App.css'
import { api } from './api'
import AdjudicationQueue from './components/AdjudicationQueue'
import AuthorGate from './components/AuthorGate'
import CommentsSidebar from './components/CommentsSidebar'
import DocumentViewer from './components/DocumentViewer'
import SettingsModal from './components/SettingsModal'
import UploadPanel from './components/UploadPanel'
import { loadCachedAiSettings } from './aiSettingsCache'
import { scrollTargetForComment } from './documentUtils'
import type { Adjudication, AppMode, DocumentMeta, DocumentPayload, ViewMode } from './types'

const AUTHOR_NAME_KEY = 'pptx_dev_author_name'
const ACTIVE_DOC_KEY = 'pptx_dev_active_doc_id'
const VIEW_MODE_KEY = 'pptx_dev_view_mode'
const APP_MODE_KEY = 'pptx_dev_app_mode'
const COMMENTS_WIDTH_KEY = 'pptx_dev_comments_width'
const COMMENTS_WIDTH_MIN = 260
const COMMENTS_WIDTH_MAX = 640
const COMMENTS_WIDTH_DEFAULT = 340

function clampCommentsWidth(w: number): number {
  return Math.min(COMMENTS_WIDTH_MAX, Math.max(COMMENTS_WIDTH_MIN, w))
}

function isViewMode(v: string | null): v is ViewMode {
  return v === 'slides' || v === 'grid' || v === 'outline'
}

function isAppMode(v: string | null): v is AppMode {
  return v === 'review' || v === 'adjudicate'
}

export default function App() {
  const [authorName, setAuthorName] = useState<string>(
    () => localStorage.getItem(AUTHOR_NAME_KEY) || '',
  )
  const [documents, setDocuments] = useState<DocumentMeta[]>([])
  const [activeDocId, setActiveDocId] = useState<string | null>(null)
  const [payload, setPayload] = useState<DocumentPayload | null>(null)
  const [selectedCommentId, setSelectedCommentId] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [adjudications, setAdjudications] = useState<Record<string, Adjudication>>({})
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [docsDrawerOpen, setDocsDrawerOpen] = useState(false)
  const [commentsDrawerOpen, setCommentsDrawerOpen] = useState(false)
  const [viewMode, setViewModeState] = useState<ViewMode>(() => {
    const stored = localStorage.getItem(VIEW_MODE_KEY)
    return isViewMode(stored) ? stored : 'slides'
  })
  const [scrollToId, setScrollToId] = useState<string | null>(null)
  const [contextCommentId, setContextCommentId] = useState<string | null>(null)
  const [contextScrollToId, setContextScrollToId] = useState<string | null>(null)
  const [appMode, setAppModeState] = useState<AppMode>(() => {
    const stored = localStorage.getItem(APP_MODE_KEY)
    return isAppMode(stored) ? stored : 'review'
  })
  const [aiReviewBusy, setAiReviewBusy] = useState(false)
  const [commentsWidth, setCommentsWidth] = useState<number>(() => {
    const stored = Number(localStorage.getItem(COMMENTS_WIDTH_KEY))
    return Number.isFinite(stored) && stored > 0 ? clampCommentsWidth(stored) : COMMENTS_WIDTH_DEFAULT
  })
  const [resizing, setResizing] = useState(false)

  function setViewMode(mode: ViewMode) {
    localStorage.setItem(VIEW_MODE_KEY, mode)
    setViewModeState(mode)
  }

  function setAppMode(mode: AppMode) {
    localStorage.setItem(APP_MODE_KEY, mode)
    setAppModeState(mode)
    setCommentsDrawerOpen(false)
  }

  function jumpToSlide(slideId: string) {
    setViewMode('slides')
    setScrollToId(slideId)
  }

  function jumpToComment(commentId: string) {
    setSelectedCommentId(commentId)
    const target = payload && scrollTargetForComment(payload.document, commentId)
    if (target) {
      // Outline and Grid modes don't render the full slide, so there's
      // nothing to scroll to there -- switch to the Slides view.
      if (viewMode !== 'slides') setViewMode('slides')
      setScrollToId(target)
    }
    // On mobile the comments drawer covers the document, so the jump would
    // otherwise happen invisibly behind it.
    setCommentsDrawerOpen(false)
  }

  function openCommentContext(commentId: string) {
    setContextCommentId(commentId)
    const target = payload && scrollTargetForComment(payload.document, commentId)
    setContextScrollToId(target ?? null)
  }

  function startResizingComments(e: React.PointerEvent) {
    e.preventDefault()
    const startX = e.clientX
    const startWidth = commentsWidth
    setResizing(true)

    function onMove(ev: PointerEvent) {
      // Dragging left (toward the document) widens the right-hand panel.
      setCommentsWidth(clampCommentsWidth(startWidth + (startX - ev.clientX)))
    }
    function onUp() {
      document.removeEventListener('pointermove', onMove)
      document.removeEventListener('pointerup', onUp)
      setResizing(false)
      setCommentsWidth((w) => {
        localStorage.setItem(COMMENTS_WIDTH_KEY, String(w))
        return w
      })
    }
    document.addEventListener('pointermove', onMove)
    document.addEventListener('pointerup', onUp)
  }

  useEffect(() => {
    ;(async () => {
      const docs = await refreshDocumentList()
      await restoreLastOpenDocument(docs)
    })()
    restoreCachedAiSettings()
  }, [])

  async function restoreCachedAiSettings() {
    const cached = loadCachedAiSettings()
    const overrides: {
      api_key?: string
      base_url?: string
      model?: string
      auth_mode?: string
      extra_headers?: string
    } = {}
    if (cached.apiKey) overrides.api_key = cached.apiKey
    if (cached.baseUrl) overrides.base_url = cached.baseUrl
    if (cached.model) overrides.model = cached.model
    if (cached.authMode) overrides.auth_mode = cached.authMode
    if (cached.extraHeaders) overrides.extra_headers = cached.extraHeaders
    if (Object.keys(overrides).length === 0) return
    try {
      await api.restoreAiSettings(overrides)
    } catch {
      // Ignore -- if this fails, it'll surface naturally the moment the
      // user actually tries an AI feature, with its own inline error.
    }
  }

  // The document itself, its comments, decisions and chats all already
  // persist server-side (JSON files on a mounted disk) -- but without this,
  // a plain browser refresh still dropped the user back to the empty
  // "Upload a .pptx" screen, since nothing remembered which document (of
  // possibly several already uploaded) they had open. Re-opening the last
  // active document on load closes that gap.
  async function restoreLastOpenDocument(docs: DocumentMeta[]) {
    const cachedId = localStorage.getItem(ACTIVE_DOC_KEY)
    if (!cachedId) return
    if (!docs.some((d) => d.id === cachedId)) {
      localStorage.removeItem(ACTIVE_DOC_KEY)
      return
    }
    try {
      await openDocument(cachedId)
    } catch {
      // Stale/unreadable doc -- fall back to the empty state rather than
      // getting stuck retrying it on every future load.
      localStorage.removeItem(ACTIVE_DOC_KEY)
    }
  }

  async function refreshDocumentList() {
    const res = await api.listDocuments()
    setDocuments(res.documents)
    return res.documents
  }

  async function openDocument(docId: string) {
    setError(null)
    const p = await api.getDocument(docId)
    setActiveDocId(docId)
    setPayload(p)
    setSelectedCommentId(null)
    setScrollToId(null)
    localStorage.setItem(ACTIVE_DOC_KEY, docId)
  }

  async function handleUpload(file: File) {
    const p = await api.uploadDocument(file)
    await refreshDocumentList()
    setActiveDocId(p.meta.id)
    setPayload(p)
    setSelectedCommentId(null)
    setScrollToId(null)
    localStorage.setItem(ACTIVE_DOC_KEY, p.meta.id)
  }

  async function handleDelete(docId: string) {
    await api.deleteDocument(docId)
    if (docId === activeDocId) {
      setActiveDocId(null)
      setPayload(null)
      localStorage.removeItem(ACTIVE_DOC_KEY)
    }
    await refreshDocumentList()
  }

  function applyResult(fn: () => Promise<DocumentPayload>) {
    return async () => {
      try {
        setError(null)
        const p = await fn()
        setPayload(p)
        if (p.adjudication && selectedCommentId) {
          setAdjudications((prev) => ({ ...prev, [selectedCommentId]: p.adjudication! }))
        }
      } catch (e) {
        setError((e as Error).message)
        // Re-throw so the button/component that triggered this (which may
        // have its own local busy/error state, e.g. SelectionPopup or a
        // CommentsSidebar card) can also show inline feedback, not just the
        // page-level banner.
        throw e
      }
    }
  }

  function saveAuthorName(name: string) {
    localStorage.setItem(AUTHOR_NAME_KEY, name)
    setAuthorName(name)
  }

  if (!authorName) {
    return <AuthorGate onSubmit={saveAuthorName} />
  }

  if (!payload) {
    return (
      <div className="app-shell app-shell-empty">
        <UploadPanel
          documents={documents}
          activeDocId={activeDocId}
          onSelect={openDocument}
          onUpload={handleUpload}
          onDelete={handleDelete}
          onOpenSettings={() => setSettingsOpen(true)}
        />
        <div className="empty-state">
          <p>Upload a .pptx (with or without PowerPoint comments) to get started.</p>
        </div>
        {settingsOpen && (
          <SettingsModal
            onClose={() => setSettingsOpen(false)}
            authorName={authorName}
            onAuthorNameChange={saveAuthorName}
          />
        )}
      </div>
    )
  }

  const docId = activeDocId!
  const commentCount = Object.keys(payload.document.comments).length
  const anyDrawerOpen = docsDrawerOpen || commentsDrawerOpen || contextCommentId !== null
  const closeDrawers = () => {
    setDocsDrawerOpen(false)
    setCommentsDrawerOpen(false)
    setContextCommentId(null)
  }

  return (
    <div
      className="app-shell app-shell-loaded"
      style={{ '--comments-width': `${commentsWidth}px` } as React.CSSProperties}
    >
      <div className={`drawer drawer-left ${docsDrawerOpen ? 'drawer-open' : ''}`}>
        <button className="drawer-close" onClick={() => setDocsDrawerOpen(false)}>
          ✕
        </button>
        <UploadPanel
          documents={documents}
          activeDocId={activeDocId}
          onSelect={(id) => {
            openDocument(id)
            setDocsDrawerOpen(false)
          }}
          onUpload={async (file) => {
            await handleUpload(file)
            setDocsDrawerOpen(false)
          }}
          onDelete={handleDelete}
          onOpenSettings={() => {
            setDocsDrawerOpen(false)
            setSettingsOpen(true)
          }}
        />
      </div>

      <div className="main-panel">
        <div className="toolbar">
          <button
            className="mobile-toolbar-button"
            title="Documents"
            onClick={() => setDocsDrawerOpen(true)}
          >
            ☰
          </button>
          <h2>{payload.meta.title}</h2>
          <div className="toolbar-actions">
            <div className="mode-toggle">
              <button
                className={appMode === 'review' ? 'mode-tab mode-tab-active' : 'mode-tab'}
                onClick={() => setAppMode('review')}
              >
                📝 Review
              </button>
              <button
                className={appMode === 'adjudicate' ? 'mode-tab mode-tab-active' : 'mode-tab'}
                onClick={() => setAppMode('adjudicate')}
              >
                ✅ Adjudicate
              </button>
            </div>
            {appMode === 'review' && (
              <>
                <select
                  className="view-mode-select"
                  value={viewMode}
                  onChange={(e) => setViewMode(e.target.value as ViewMode)}
                  title="View"
                >
                  <option value="slides">Slides</option>
                  <option value="grid">Slide Sorter</option>
                  <option value="outline">Outline</option>
                </select>
                <button
                  className="ai-button"
                  disabled={aiReviewBusy}
                  onClick={async () => {
                    setAiReviewBusy(true)
                    setError(null)
                    try {
                      const p = await api.runAiReview(docId, 'AI Reviewer')
                      setPayload(p)
                    } catch (e) {
                      setError((e as Error).message)
                    } finally {
                      setAiReviewBusy(false)
                    }
                  }}
                >
                  {aiReviewBusy ? 'Reviewing…' : '✨ AI Review'}
                </button>
                <button
                  className="mobile-toolbar-button"
                  title="Comments"
                  onClick={() => setCommentsDrawerOpen(true)}
                >
                  💬 {commentCount}
                </button>
              </>
            )}
            <button
              onClick={async () => {
                if (!confirm('Discard all edits and reset to the original upload?')) return
                const p = await api.resetDocument(docId)
                setPayload(p)
              }}
            >
              Reset
            </button>
            <a className="export-button" href={api.exportUrl(docId)}>
              Export .pptx
            </a>
          </div>
        </div>
        {error && <p className="error-text">{error}</p>}
        {appMode === 'adjudicate' ? (
          <AdjudicationQueue
            docId={docId}
            document={payload.document}
            meta={payload.meta}
            decisions={payload.decisions}
            chats={payload.chats}
            commentRefs={payload.commentRefs}
            authorName={authorName}
            onDocumentUpdate={setPayload}
            onViewContext={openCommentContext}
          />
        ) : (
          <DocumentViewer
            docId={docId}
            slides={payload.document.slides}
            slideSize={payload.document.slideSize}
            comments={payload.document.comments}
            viewMode={viewMode}
            scrollToId={scrollToId}
            onScrolled={() => setScrollToId(null)}
            onJumpToSlide={jumpToSlide}
            selectedCommentId={selectedCommentId}
            onSelectComment={(id) => {
              setSelectedCommentId(id)
              setCommentsDrawerOpen(false)
            }}
            onReplaceRange={async (startRunId, startOffset, endRunId, endOffset, newText) =>
              applyResult(() =>
                api.replaceTextRange(docId, startRunId, startOffset, endRunId, endOffset, newText, authorName),
              )()
            }
            onAddComment={async (startRunId, startOffset, endRunId, endOffset, text) =>
              applyResult(() =>
                api.createComment(docId, startRunId, startOffset, endRunId, endOffset, text, authorName),
              )()
            }
            onAddSlideComment={async (slideId, text) =>
              applyResult(() => api.createSlideComment(docId, slideId, text, authorName))()
            }
          />
        )}
      </div>

      {(appMode === 'review' || (appMode === 'adjudicate' && contextCommentId)) && (
        <div
          className={`resize-handle ${resizing ? 'resize-handle-active' : ''}`}
          onPointerDown={startResizingComments}
          title="Drag to resize"
        />
      )}

      {appMode === 'review' && (
        <div className={`drawer drawer-right ${commentsDrawerOpen ? 'drawer-open' : ''}`}>
          <button className="drawer-close" onClick={() => setCommentsDrawerOpen(false)}>
            ✕
          </button>
          <CommentsSidebar
            comments={payload.document.comments}
            slides={payload.document.slides}
            selectedCommentId={selectedCommentId}
            onSelectComment={jumpToComment}
            onReply={async (commentId: string, text: string) =>
              applyResult(() => api.replyToComment(docId, commentId, text, authorName))()
            }
            onResolve={async (commentId: string, done: boolean) =>
              applyResult(() => api.resolveComment(docId, commentId, done, authorName))()
            }
            onAdjudicate={async (commentId: string) => {
              setSelectedCommentId(commentId)
              try {
                setError(null)
                const p = await api.adjudicateComment(docId, commentId)
                setPayload(p)
                if (p.adjudication) {
                  setAdjudications((prev) => ({ ...prev, [commentId]: p.adjudication! }))
                }
              } catch (e) {
                setError((e as Error).message)
                // Re-throw so the comment card's own "Ask AI" button (which
                // tracks its own busy/error state) can show inline feedback
                // too, not just the page-level banner.
                throw e
              }
            }}
            lastAdjudicationByComment={adjudications}
          />
        </div>
      )}

      {appMode === 'adjudicate' && contextCommentId && (
        <div className="drawer drawer-right drawer-open">
          <button className="drawer-close" onClick={() => setContextCommentId(null)}>
            ✕
          </button>
          <div className="context-drawer">
            <div className="context-drawer-header">
              <h3>Comment source</h3>
              <button className="context-drawer-close" onClick={() => setContextCommentId(null)}>
                ✕ Close
              </button>
            </div>
            <DocumentViewer
              docId={docId}
              slides={
                // Only the comment's own slide: the drawer is a "where is this
                // comment" view, not a second copy of the whole deck.
                payload.document.slides.filter(
                  (sl) => sl.id === payload.document.comments[contextCommentId]?.slideId,
                )
              }
              slideSize={payload.document.slideSize}
              comments={payload.document.comments}
              viewMode="slides"
              scrollToId={contextScrollToId}
              onScrolled={() => setContextScrollToId(null)}
              onJumpToSlide={() => {}}
              selectedCommentId={contextCommentId}
              onSelectComment={() => {}}
              onReplaceRange={async (startRunId, startOffset, endRunId, endOffset, newText) =>
                applyResult(() =>
                  api.replaceTextRange(docId, startRunId, startOffset, endRunId, endOffset, newText, authorName),
                )()
              }
              onAddComment={async (startRunId, startOffset, endRunId, endOffset, text) =>
                applyResult(() =>
                  api.createComment(docId, startRunId, startOffset, endRunId, endOffset, text, authorName),
                )()
              }
              onAddSlideComment={async (slideId, text) =>
                applyResult(() => api.createSlideComment(docId, slideId, text, authorName))()
              }
            />
          </div>
        </div>
      )}

      {anyDrawerOpen && <div className="drawer-backdrop" onClick={closeDrawers} />}
      {settingsOpen && (
        <SettingsModal
          onClose={() => setSettingsOpen(false)}
          authorName={authorName}
          onAuthorNameChange={saveAuthorName}
        />
      )}
    </div>
  )
}
