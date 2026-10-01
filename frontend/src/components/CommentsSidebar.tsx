import { useState } from 'react'
import type { CommentEntry, Slide } from '../types'

interface Props {
  comments: Record<string, CommentEntry>
  slides: Slide[]
  selectedCommentId: string | null
  onSelectComment: (id: string) => void
  onReply: (commentId: string, text: string) => Promise<void>
  onResolve: (commentId: string, done: boolean) => Promise<void>
  onAdjudicate: (commentId: string) => Promise<void>
  lastAdjudicationByComment: Record<string, { action: string; reply: string; reasoning: string }>
}

function ThreadCard({
  comment,
  replies,
  slideNumber,
  ...rest
}: {
  comment: CommentEntry
  replies: CommentEntry[]
  slideNumber: number | undefined
} & Omit<Props, 'comments' | 'slides'>) {
  const [replyText, setReplyText] = useState('')
  const [busy, setBusy] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const isSelected = rest.selectedCommentId === comment.id
  const adjudication = rest.lastAdjudicationByComment[comment.id]

  async function run(label: string, fn: () => Promise<void>) {
    setBusy(label)
    setError(null)
    try {
      await fn()
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(null)
    }
  }

  return (
    <div
      className={`comment-card ${isSelected ? 'comment-card-selected' : ''} ${comment.done ? 'comment-card-done' : ''}`}
      onClick={() => rest.onSelectComment(comment.id)}
    >
      <div className="comment-header">
        {slideNumber !== undefined && <span className="comment-slide-chip">Slide {slideNumber}</span>}
        <span className="comment-author">{comment.author}</span>
        {comment.date && <span className="comment-date">{comment.date.slice(0, 10)}</span>}
        {comment.done && <span className="comment-badge">resolved</span>}
      </div>
      {comment.quote && <div className="comment-quote">“{comment.quote}”</div>}
      <div className="comment-text">{comment.text}</div>

      {replies.map((r) => (
        <div className="comment-reply" key={r.id}>
          <div className="comment-header">
            <span className="comment-author">{r.author}</span>
            {r.date && <span className="comment-date">{r.date.slice(0, 10)}</span>}
          </div>
          <div className="comment-text">{r.text}</div>
        </div>
      ))}

      {adjudication && (
        <div className="adjudication-note">
          <strong>AI: {adjudication.action}</strong>
          <div>{adjudication.reasoning}</div>
        </div>
      )}

      <div className="comment-actions" onClick={(e) => e.stopPropagation()}>
        <input
          className="comment-reply-input"
          placeholder="Reply..."
          value={replyText}
          onChange={(e) => setReplyText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && replyText.trim()) {
              run('reply', async () => {
                await rest.onReply(comment.id, replyText)
                setReplyText('')
              })
            }
          }}
        />
        <button
          disabled={busy !== null || !replyText.trim()}
          onClick={() =>
            run('reply', async () => {
              await rest.onReply(comment.id, replyText)
              setReplyText('')
            })
          }
        >
          {busy === 'reply' ? 'Replying…' : 'Reply'}
        </button>
        <button
          disabled={busy !== null}
          onClick={() => run('resolve', () => rest.onResolve(comment.id, !comment.done))}
        >
          {comment.done ? 'Reopen' : 'Resolve'}
        </button>
        <button
          className="ai-button"
          disabled={busy !== null}
          onClick={() => run('ai', () => rest.onAdjudicate(comment.id))}
        >
          {busy === 'ai' ? 'Thinking…' : '✨ Ask AI'}
        </button>
      </div>

      {error && (
        <p className="error-text comment-error" onClick={(e) => e.stopPropagation()}>
          {error}
        </p>
      )}
    </div>
  )
}

export default function CommentsSidebar({ comments, slides, ...rest }: Props) {
  const slideIndex: Record<string, number> = {}
  for (const sl of slides) slideIndex[sl.id] = sl.index

  const all = Object.values(comments)
  const byDeckOrder = (a: CommentEntry, b: CommentEntry) =>
    (slideIndex[a.slideId] ?? 0) - (slideIndex[b.slideId] ?? 0) || (a.date || '').localeCompare(b.date || '')
  const topLevel = all.filter((c) => !c.parentId && c.kind !== 'edit').sort(byDeckOrder)
  const edits = all.filter((c) => c.kind === 'edit').sort(byDeckOrder)
  const repliesByParent: Record<string, CommentEntry[]> = {}
  for (const c of all) {
    if (c.parentId) {
      repliesByParent[c.parentId] = repliesByParent[c.parentId] || []
      repliesByParent[c.parentId].push(c)
    }
  }

  return (
    <div className="comments-sidebar">
      <h3>Comments ({topLevel.length})</h3>
      {topLevel.length === 0 && <p className="empty-hint">No comments in this presentation.</p>}
      {topLevel.map((c) => (
        <ThreadCard
          key={c.id}
          comment={c}
          replies={repliesByParent[c.id] || []}
          slideNumber={slideIndex[c.slideId]}
          {...rest}
        />
      ))}

      {edits.length > 0 && (
        <details className="edit-records">
          <summary>Text edits ({edits.length})</summary>
          <p className="empty-hint">
            PowerPoint has no Track Changes, so each text edit is recorded as an automatic, already-resolved
            comment.
          </p>
          {edits.map((c) => (
            <div
              key={c.id}
              className="edit-record"
              onClick={() => rest.onSelectComment(c.id)}
              title="Jump to the edited text"
            >
              <span className="comment-slide-chip">Slide {slideIndex[c.slideId]}</span>
              <span className="comment-author">{c.author}</span>
              {c.date && <span className="comment-date">{c.date.slice(0, 10)}</span>}
              <div className="comment-text">{c.text}</div>
            </div>
          ))}
        </details>
      )}
    </div>
  )
}
