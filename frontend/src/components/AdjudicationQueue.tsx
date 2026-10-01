import { useEffect, useRef, useState } from 'react'
import { api } from '../api'
import SlideCanvas from './SlideCanvas'
import { findAnchorForComment, paragraphText, shapeParagraphs, type CommentAnchor } from '../documentUtils'
import type {
  AdjudicationSuggestion,
  ChatTurn,
  CommentEntry,
  CommentRefGuess,
  DecisionRecord,
  DecisionValue,
  DocumentMeta,
  DocumentModel,
  DocumentPayload,
  SuggestionError,
  XlsxImportResult,
} from '../types'

const DECISION_LABELS: Record<DecisionValue, string> = {
  accept: 'Accept',
  reject: 'Reject',
  info_only: 'Info Only',
  defer: 'Defer',
}

interface Props {
  docId: string
  document: DocumentModel
  meta: DocumentMeta
  decisions: Record<string, DecisionRecord>
  chats: Record<string, ChatTurn[]>
  commentRefs: Record<string, CommentRefGuess>
  authorName: string
  onDocumentUpdate: (payload: DocumentPayload) => void
  onViewContext: (commentId: string) => void
}

function AnchorPreview({
  anchor,
  comment,
  docId,
  document,
}: {
  anchor: CommentAnchor
  comment: CommentEntry
  docId: string
  document: DocumentModel
}) {
  const { slide, shape } = anchor
  const paragraphs = shape ? shapeParagraphs(shape) : []
  const quote = comment.quote?.trim() || ''
  return (
    <div className="adjudication-context">
      <div className="adjudication-context-where">
        Slide {slide.index}
        {slide.title ? ` — ${slide.title}` : ''}
        {!shape && ' · whole slide'}
        {shape && shape.type === 'table' && ' · table'}
        {shape && shape.type === 'picture' && ' · picture'}
      </div>
      {paragraphs.map((p) => {
        const text = paragraphText(p)
        if (!text.trim()) return null
        const at = quote ? text.indexOf(quote) : -1
        return (
          <p key={p.id} className="adjudication-context-text">
            {at >= 0 ? (
              <>
                {text.slice(0, at)}
                <span className="context-highlight">{quote}</span>
                {text.slice(at + quote.length)}
              </>
            ) : (
              <span className={quote ? undefined : 'context-highlight'}>{text}</span>
            )}
          </p>
        )
      })}
      <div className="adjudication-slide-thumb">
        <SlideCanvas
          docId={docId}
          slide={slide}
          slideSize={document.slideSize}
          comments={document.comments}
          selectedCommentId={null}
          onSelectComment={() => {}}
          inert
          highlightShapeId={shape?.id ?? null}
        />
      </div>
    </div>
  )
}

function isSuggestionError(s: AdjudicationSuggestion | SuggestionError | undefined): s is SuggestionError {
  return !!s && 'error' in s
}

interface CardProps {
  docId: string
  document: DocumentModel
  comment: CommentEntry
  anchor: CommentAnchor | null
  suggestion: AdjudicationSuggestion | SuggestionError | undefined
  decision: DecisionRecord | undefined
  refGuess: CommentRefGuess | undefined
  chat: ChatTurn[]
  authorName: string
  onApplied: (payload: DocumentPayload) => void
  onSkip: () => void
  onViewContext: () => void
}

function QueueCard({
  docId, document, comment, anchor, suggestion, decision, refGuess, chat, authorName, onApplied, onSkip, onViewContext,
}: CardProps) {
  const sugg = !isSuggestionError(suggestion) ? suggestion : undefined
  const [replacementText, setReplacementText] = useState(sugg?.replacement_text || '')
  // The text being replaced: the reviewer's own selection when there is one
  // (fixed), otherwise whatever the AI proposed (editable -- PowerPoint
  // comments attach to a whole shape unless text was selected).
  const [originalInput, setOriginalInput] = useState(sugg?.original_text || '')
  const [replyText, setReplyText] = useState(sugg?.reply || '')
  const [manualReply, setManualReply] = useState('')
  const [decisionValue, setDecisionValue] = useState<DecisionValue | null>(decision?.decision ?? null)
  const [reasonText, setReasonText] = useState(decision?.reason ?? '')
  const [refText, setRefText] = useState(decision?.ref ?? refGuess?.ref ?? '')
  const [chatOpen, setChatOpen] = useState(chat.length > 0)
  const [chatInput, setChatInput] = useState('')
  const [chatBusy, setChatBusy] = useState(false)
  const [chatError, setChatError] = useState<string | null>(null)
  const [copiedTurnIndex, setCopiedTurnIndex] = useState<number | null>(null)
  const [collapsed, setCollapsed] = useState(false)
  const [busy, setBusy] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    setDecisionValue(decision?.decision ?? null)
    setReasonText(decision?.reason ?? '')
    setRefText(decision?.ref ?? refGuess?.ref ?? '')
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [decision?.decision, decision?.reason, decision?.ref, refGuess?.ref])

  useEffect(() => {
    setReplacementText(sugg?.replacement_text || '')
    setOriginalInput(sugg?.original_text || '')
    setReplyText(sugg?.reply || '')
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sugg?.replacement_text, sugg?.original_text, sugg?.reply])

  const originalText = comment.quote || originalInput

  async function applySuggestion() {
    if (!sugg) return
    setBusy('apply')
    setError(null)
    try {
      const payload = await api.applyAdjudication(
        docId,
        comment.id,
        sugg.action,
        replyText,
        sugg.action === 'edit' ? replacementText : null,
        sugg.action === 'edit' ? originalText || null : null,
        sugg.reasoning,
        'AI Assistant',
      )
      onApplied(payload)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  async function submitReply() {
    if (!manualReply.trim()) return
    setBusy('reply')
    setError(null)
    try {
      const payload = await api.replyToComment(docId, comment.id, manualReply.trim(), authorName)
      onApplied(payload)
      setManualReply('')
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  async function submitResolve() {
    setBusy('resolve')
    setError(null)
    try {
      const payload = await api.resolveComment(docId, comment.id, !comment.done, authorName)
      onApplied(payload)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  async function submitDecision() {
    if (!decisionValue || !reasonText.trim()) return
    setBusy('decision')
    setError(null)
    try {
      const payload = await api.saveDecision(
        docId, comment.id, decisionValue, reasonText.trim(), refText.trim(), authorName,
      )
      onApplied(payload)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  async function sendChat() {
    const message = chatInput.trim()
    if (!message) return
    setChatBusy(true)
    setChatError(null)
    setChatInput('')
    try {
      const payload = await api.sendChatMessage(docId, comment.id, message, authorName)
      onApplied(payload)
    } catch (e) {
      setChatError((e as Error).message)
    } finally {
      setChatBusy(false)
    }
  }

  async function copyChatTurn(index: number, text: string) {
    try {
      await navigator.clipboard.writeText(text)
      setCopiedTurnIndex(index)
      setTimeout(() => setCopiedTurnIndex((i) => (i === index ? null : i)), 1500)
    } catch {
      // Clipboard API unavailable (non-secure context, permissions denied) --
      // the text is still plain, selectable prose, just without the button.
    }
  }

  return (
    <div className={`queue-card${collapsed ? ' queue-card-collapsed' : ''}`}>
      <div className="queue-card-header">
        <button
          type="button"
          className="queue-collapse-toggle"
          onClick={() => setCollapsed((c) => !c)}
          title={collapsed ? 'Expand' : 'Collapse'}
        >
          {collapsed ? '▸' : '▾'}
        </button>
        <span className="comment-author">{comment.author}</span>
        {comment.date && <span className="comment-date">{comment.date.slice(0, 10)}</span>}
        {decision && (
          <span className={`queue-card-decision-badge queue-card-decision-badge-${decision.decision}`}>
            {DECISION_LABELS[decision.decision]}
          </span>
        )}
        {collapsed && <span className="queue-card-collapsed-preview">{comment.text}</span>}
        <button className="queue-skip" onClick={onSkip} title="Hide from this queue for now">
          Skip ›
        </button>
      </div>

      {!collapsed && (
        <>
          <div
            className="comment-text queue-comment-text-clickable"
            onClick={onViewContext}
            title="View this comment's source text in the document"
          >
            {comment.text}
          </div>

          {anchor && <AnchorPreview anchor={anchor} comment={comment} docId={docId} document={document} />}

          {!suggestion && <p className="empty-hint">Analyzing…</p>}
          {isSuggestionError(suggestion) && <p className="error-text">{suggestion.error}</p>}

      {sugg && (
        <div className={`queue-suggestion queue-suggestion-${sugg.action}`}>
          <div className="queue-suggestion-badge">
            {sugg.action === 'edit' && '✏️ AI proposes an edit'}
            {sugg.action === 'reply' && '💬 AI proposes a reply'}
            {sugg.action === 'no_change' && '⚠️ AI has no confident suggestion'}
          </div>

          {sugg.action === 'edit' && (
            <div className="queue-edit-field">
              {comment.quote ? (
                <span className="run-del">{comment.quote}</span>
              ) : (
                <input
                  className="queue-replacement-input queue-original-input"
                  value={originalInput}
                  placeholder="Exact text on the slide to replace"
                  onChange={(e) => setOriginalInput(e.target.value)}
                />
              )}
              <span className="queue-arrow">→</span>
              <input
                className="queue-replacement-input"
                value={replacementText}
                onChange={(e) => setReplacementText(e.target.value)}
              />
            </div>
          )}

          <textarea
            className="queue-reply-textarea"
            value={replyText}
            onChange={(e) => setReplyText(e.target.value)}
          />
          <p className="queue-suggestion-reasoning">{sugg.reasoning}</p>

          <div className="queue-actions">
            <button className="ai-button" disabled={busy !== null} onClick={applySuggestion}>
              {busy === 'apply' ? 'Applying…' : '✓ Apply'}
            </button>
          </div>
        </div>
      )}

      <div className="queue-manual">
        <div className="queue-manual-form">
          <input
            className="comment-reply-input"
            placeholder="Reply..."
            value={manualReply}
            onChange={(e) => setManualReply(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && manualReply.trim()) submitReply()
            }}
          />
          <button disabled={busy !== null || !manualReply.trim()} onClick={submitReply}>
            {busy === 'reply' ? 'Replying…' : 'Reply'}
          </button>
          <button disabled={busy !== null} onClick={submitResolve}>
            {busy === 'resolve' ? 'Working…' : comment.done ? 'Reopen' : 'Resolve'}
          </button>
        </div>
      </div>

      <div className="queue-decision">
        <div className="queue-decision-header">
          <span>Adjudication decision</span>
          {decision && (
            <span className="queue-decision-meta">
              {DECISION_LABELS[decision.decision]} by {decision.author} on {decision.decidedAt.slice(0, 10)}
            </span>
          )}
        </div>
        <div className="queue-decision-buttons">
          {(Object.keys(DECISION_LABELS) as DecisionValue[]).map((v) => (
            <button
              key={v}
              type="button"
              className={`queue-decision-btn queue-decision-btn-${v}${decisionValue === v ? ' active' : ''}`}
              onClick={() => setDecisionValue(v)}
            >
              {DECISION_LABELS[v]}
            </button>
          ))}
        </div>
        <textarea
          className="queue-decision-reason"
          placeholder="Rationale for this decision (required)…"
          value={reasonText}
          onChange={(e) => setReasonText(e.target.value)}
        />
        {(refGuess?.ref || refGuess?.para) && (
          <p className="queue-decision-hint">
            Detected from document:{' '}
            {refGuess.ref && (
              <>
                Ref # <strong>{refGuess.ref}</strong>
              </>
            )}
            {refGuess.ref && refGuess.para && ' · '}
            {refGuess.para && (
              <>
                Para # <strong>{refGuess.para}</strong>
              </>
            )}
          </p>
        )}
        <input
          className="queue-decision-ref"
          placeholder="Ref # (optional, e.g. spec section 3.2.2)"
          value={refText}
          onChange={(e) => setRefText(e.target.value)}
        />
        <button
          className="ai-button"
          disabled={busy !== null || !decisionValue || !reasonText.trim()}
          onClick={submitDecision}
        >
          {busy === 'decision' ? 'Saving…' : decision ? 'Update Decision' : 'Save Decision'}
        </button>
      </div>

      <div className="queue-chat">
        <button type="button" className="queue-chat-toggle" onClick={() => setChatOpen((o) => !o)}>
          💬 Discuss with AI{chat.length > 0 ? ` (${chat.length})` : ''} {chatOpen ? '▾' : '▸'}
        </button>
        {chatOpen && (
          <div className="queue-chat-panel">
            {chat.length === 0 && (
              <p className="empty-hint">
                Talk through this comment with the AI -- useful for developing a rationale, or when
                it had no confident suggestion above.
              </p>
            )}
            {chat.map((turn, i) => (
              <div key={i} className={`chat-turn chat-turn-${turn.role}`}>
                <div className="chat-turn-head">
                  <span className="chat-turn-role">
                    {turn.role === 'user' ? turn.author || authorName : 'AI Assistant'}
                  </span>
                  {turn.role === 'assistant' && (
                    <button
                      type="button"
                      className="chat-copy-button"
                      onClick={() => copyChatTurn(i, turn.content)}
                      title="Copy this reply"
                    >
                      {copiedTurnIndex === i ? 'Copied ✓' : '⧉ Copy'}
                    </button>
                  )}
                </div>
                <p>{turn.content}</p>
              </div>
            ))}
            <div className="queue-chat-input-row">
              <textarea
                className="queue-chat-input"
                placeholder="Ask the AI about this comment…"
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault()
                    sendChat()
                  }
                }}
              />
              <button disabled={chatBusy || !chatInput.trim()} onClick={sendChat}>
                {chatBusy ? '…' : 'Send'}
              </button>
            </div>
            {chatError && <p className="error-text">{chatError}</p>}
          </div>
        )}
      </div>

          {error && <p className="error-text">{error}</p>}
        </>
      )}
    </div>
  )
}

export default function AdjudicationQueue({
  docId, document, meta, decisions, chats, commentRefs, authorName, onDocumentUpdate, onViewContext,
}: Props) {
  const [suggestions, setSuggestions] = useState<Record<string, AdjudicationSuggestion | SuggestionError>>({})
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [skipped, setSkipped] = useState<Set<string>>(new Set())
  const [usedFullDocContext, setUsedFullDocContext] = useState<boolean | null>(null)
  const [docRefInput, setDocRefInput] = useState(meta.doc_ref ?? '')
  const [revInput, setRevInput] = useState(meta.rev ?? '')
  const [savingMeta, setSavingMeta] = useState(false)
  const [importBusy, setImportBusy] = useState(false)
  const [importResult, setImportResult] = useState<XlsxImportResult | null>(null)
  const importFileInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    setDocRefInput(meta.doc_ref ?? '')
    setRevInput(meta.rev ?? '')
  }, [meta.id])

  const slideNumber = (slideId: string) => document.slides.find((sl) => sl.id === slideId)?.index ?? 0
  const openComments = Object.values(document.comments)
    .filter((c) => !c.parentId && !c.done && c.kind !== 'edit' && !skipped.has(c.id))
    .sort((a, b) => slideNumber(a.slideId) - slideNumber(b.slideId) || (a.date || '').localeCompare(b.date || ''))

  async function analyzeAll() {
    setLoading(true)
    setError(null)
    try {
      const res = await api.suggestAdjudications(docId)
      setSuggestions(res.suggestions)
      setUsedFullDocContext(res.usedFullDocumentContext)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setLoading(false)
    }
  }

  async function saveDocMeta() {
    setSavingMeta(true)
    setError(null)
    try {
      const payload = await api.setDocumentMeta(docId, { doc_ref: docRefInput.trim(), rev: revInput.trim() })
      onDocumentUpdate(payload)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setSavingMeta(false)
    }
  }

  async function handleImportFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    e.target.value = '' // allow re-selecting the same file again later
    if (!file) return
    setImportBusy(true)
    setError(null)
    setImportResult(null)
    try {
      const payload = await api.importCrmXlsx(docId, file, authorName)
      onDocumentUpdate(payload)
      setImportResult(payload.xlsxImport ?? null)
    } catch (err) {
      setError((err as Error).message)
    } finally {
      setImportBusy(false)
    }
  }

  return (
    <div className="adjudication-queue">
      <div className="adjudication-queue-header">
        <h2>Adjudicate ({openComments.length} open)</h2>
        <button className="ai-button" disabled={loading} onClick={analyzeAll}>
          {loading ? 'Analyzing all comments…' : '✨ Analyze All with AI'}
        </button>
      </div>
      <div className="crm-meta-bar">
        <label>
          Document ref
          <input
            value={docRefInput}
            onChange={(e) => setDocRefInput(e.target.value)}
            placeholder="e.g. UTIC-MK710-A012-0003"
          />
        </label>
        <label>
          Rev
          <input value={revInput} onChange={(e) => setRevInput(e.target.value)} placeholder="e.g. X1" />
        </label>
        <button disabled={savingMeta} onClick={saveDocMeta}>
          {savingMeta ? 'Saving…' : 'Save'}
        </button>
        <a className="export-button" href={api.exportCrmUrl(docId)}>
          ⬇ Export Comment Resolution Matrix (TSV)
        </a>
        <button disabled={importBusy} onClick={() => importFileInputRef.current?.click()}>
          {importBusy ? 'Importing…' : '⬆ Import Filled Matrix (.xlsx)'}
        </button>
        <input
          ref={importFileInputRef}
          type="file"
          accept=".xlsx,.xlsm"
          className="crm-import-input"
          onChange={handleImportFile}
        />
      </div>
      {importResult && (
        <div className="crm-import-result">
          <p>
            <strong>{importResult.applied.length}</strong> decision
            {importResult.applied.length === 1 ? '' : 's'} imported
            {importResult.applied.some((a) => a.created) && (
              <>
                {' '}
                (<strong>{importResult.applied.filter((a) => a.created).length}</strong> as brand-new
                comments -- no matching comment existed in the document, so one was created and
                anchored via that row's Ref #/Para #)
              </>
            )}
            {importResult.skipped.length > 0 && (
              <>
                , <strong>{importResult.skipped.length}</strong> row
                {importResult.skipped.length === 1 ? '' : 's'} skipped
              </>
            )}
            .
            <button className="crm-import-dismiss" onClick={() => setImportResult(null)}>
              Dismiss
            </button>
          </p>
          {importResult.skippedSummary.length > 0 && (
            <ul>
              {importResult.skippedSummary.map((group) => (
                <li key={group.reason}>
                  {group.count} row{group.count === 1 ? '' : 's'} ({group.rows.join(', ')}
                  {group.count > group.rows.length ? ', …' : ''}): {group.reason}
                </li>
              ))}
            </ul>
          )}
        </div>
      )}
      {usedFullDocContext === false && (
        <p className="empty-hint">
          Document is large — AI suggestions used per-section context rather than the whole document.
        </p>
      )}
      {error && <p className="error-text">{error}</p>}

      {openComments.length === 0 && (
        <p className="empty-hint">
          {skipped.size > 0
            ? 'All remaining comments were skipped. Switch modes and back to reset.'
            : 'No open comments to adjudicate.'}
        </p>
      )}

      {openComments.map((c) => (
        <QueueCard
          key={c.id}
          docId={docId}
          comment={c}
          document={document}
          anchor={findAnchorForComment(document, c.id)}
          suggestion={suggestions[c.id]}
          decision={decisions[c.id]}
          refGuess={commentRefs[c.id]}
          chat={chats[c.id] || []}
          authorName={authorName}
          onApplied={onDocumentUpdate}
          onViewContext={() => onViewContext(c.id)}
          onSkip={() => setSkipped((prev) => new Set(prev).add(c.id))}
        />
      ))}
    </div>
  )
}
