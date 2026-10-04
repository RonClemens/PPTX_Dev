import { aiHeadersFor, aiRequestHeaders, type AiSettings, type AuthMode } from './aiSettings'
import type { DecisionValue, DocumentMeta, DocumentPayload, SuggestAdjudicationsResponse } from './types'

// Only these requests call the AI, so only these carry the user's AI
// credentials (from this device's storage). Everything else -- uploads,
// comments, exports -- never sees them.
const AI_PATHS = /\/(adjudicate|suggest-adjudications|ai-review|chat|anthropic-key\/test)$/

async function req<T>(path: string, init?: RequestInit, extraHeaders?: Record<string, string>): Promise<T> {
  const headers: Record<string, string> = {
    ...(init?.body && !(init.body instanceof FormData) ? { 'Content-Type': 'application/json' } : {}),
    ...((init?.headers as Record<string, string> | undefined) || {}),
    ...(extraHeaders ?? (AI_PATHS.test(path) ? aiRequestHeaders() : {})),
  }
  const res = await fetch(path, { ...init, headers })
  if (!res.ok) {
    const body = await res.json().catch(() => ({}))
    throw new Error(body.detail || `request failed: ${res.status}`)
  }
  return res.json()
}

export const api = {
  listDocuments: () => req<{ documents: DocumentMeta[] }>('/api/documents'),

  uploadDocument: (file: File) => {
    const form = new FormData()
    form.append('file', file)
    return req<DocumentPayload>('/api/documents', { method: 'POST', body: form })
  },

  getDocument: (docId: string) => req<DocumentPayload>(`/api/documents/${docId}`),

  deleteDocument: (docId: string) =>
    req<{ ok: boolean }>(`/api/documents/${docId}`, { method: 'DELETE' }),

  resetDocument: (docId: string) =>
    req<DocumentPayload>(`/api/documents/${docId}/reset`, { method: 'POST' }),

  exportUrl: (docId: string) => `/api/documents/${docId}/export`,

  mediaUrl: (docId: string, slideId: string, relId: string) =>
    `/api/documents/${docId}/media/${slideId}/${relId}`,

  replaceTextRange: (
    docId: string,
    startRunId: string,
    startOffset: number,
    endRunId: string,
    endOffset: number,
    newText: string,
    author: string,
  ) =>
    req<DocumentPayload>(`/api/documents/${docId}/edits/replace-range`, {
      method: 'POST',
      body: JSON.stringify({
        start_run_id: startRunId,
        start_offset: startOffset,
        end_run_id: endRunId,
        end_offset: endOffset,
        new_text: newText,
        author,
      }),
    }),

  /** Comment on a text selection (stored against the selection's shape). */
  createComment: (
    docId: string,
    startRunId: string,
    startOffset: number,
    endRunId: string,
    endOffset: number,
    text: string,
    author: string,
  ) =>
    req<DocumentPayload>(`/api/documents/${docId}/comments`, {
      method: 'POST',
      body: JSON.stringify({
        start_run_id: startRunId,
        start_offset: startOffset,
        end_run_id: endRunId,
        end_offset: endOffset,
        text,
        author,
      }),
    }),

  /** Comment on a whole slide (or one shape on it). */
  createSlideComment: (docId: string, slideId: string, text: string, author: string, shapeId?: string) =>
    req<DocumentPayload>(`/api/documents/${docId}/comments`, {
      method: 'POST',
      body: JSON.stringify({ slide_id: slideId, shape_id: shapeId ?? null, text, author }),
    }),

  replyToComment: (docId: string, commentId: string, text: string, author: string) =>
    req<DocumentPayload>(`/api/documents/${docId}/comments/${commentId}/reply`, {
      method: 'POST',
      body: JSON.stringify({ text, author }),
    }),

  resolveComment: (docId: string, commentId: string, done: boolean, author: string) =>
    req<DocumentPayload>(`/api/documents/${docId}/comments/${commentId}/resolve`, {
      method: 'POST',
      body: JSON.stringify({ done, author }),
    }),

  saveDecision: (
    docId: string,
    commentId: string,
    decision: DecisionValue,
    reason: string,
    ref: string,
    author: string,
  ) =>
    req<DocumentPayload>(`/api/documents/${docId}/comments/${commentId}/decision`, {
      method: 'POST',
      body: JSON.stringify({ decision, reason, ref, author }),
    }),

  exportCrmUrl: (docId: string) => `/api/documents/${docId}/comments/export`,

  importCrmXlsx: (docId: string, file: File, author: string) => {
    const form = new FormData()
    form.append('file', file)
    form.append('author', author)
    return req<DocumentPayload>(`/api/documents/${docId}/comments/import-decisions`, {
      method: 'POST',
      body: form,
    })
  },

  setDocumentMeta: (docId: string, fields: { doc_ref?: string; rev?: string }) =>
    req<DocumentPayload>(`/api/documents/${docId}/meta`, {
      method: 'PATCH',
      body: JSON.stringify(fields),
    }),

  adjudicateComment: (docId: string, commentId: string) =>
    req<DocumentPayload>(`/api/documents/${docId}/comments/${commentId}/adjudicate`, {
      method: 'POST',
      body: JSON.stringify({ author: 'AI Assistant' }),
    }),

  suggestAdjudications: (docId: string) =>
    req<SuggestAdjudicationsResponse>(`/api/documents/${docId}/suggest-adjudications`, {
      method: 'POST',
    }),

  applyAdjudication: (
    docId: string,
    commentId: string,
    action: 'edit' | 'reply' | 'no_change',
    reply: string,
    replacementText: string | null,
    originalText: string | null,
    reasoning: string,
    author: string,
  ) =>
    req<DocumentPayload>(`/api/documents/${docId}/comments/${commentId}/apply-adjudication`, {
      method: 'POST',
      body: JSON.stringify({
        action,
        reply,
        replacement_text: replacementText,
        original_text: originalText,
        reasoning,
        author,
      }),
    }),

  sendChatMessage: (docId: string, commentId: string, message: string, author: string) =>
    req<DocumentPayload>(`/api/documents/${docId}/comments/${commentId}/chat`, {
      method: 'POST',
      body: JSON.stringify({ message, author }),
    }),

  runAiReview: (docId: string, author: string) =>
    req<DocumentPayload>(`/api/documents/${docId}/ai-review`, {
      method: 'POST',
      body: JSON.stringify({ author }),
    }),

  /** What the SERVER's environment supplies as fallback defaults. It never
   * holds anything the user typed. */
  getServerDefaults: () => req<ServerDefaults>('/api/settings/anthropic-key'),

  /** One minimal real request with the given settings (the form's current
   * values, saved or not) -- sent as headers for this request only. */
  testAnthropicConnection: (settings: AiSettings) =>
    req<TestConnectionResult>('/api/settings/anthropic-key/test', { method: 'POST' }, aiHeadersFor(settings)),
}

export interface TestConnectionResult {
  ok: true
  model: string
  baseUrl: string | null
  latencyMs: number
  responseId: string
}

export interface ServerDefaults {
  storesCredentials: false
  serverKeyConfigured: boolean
  baseUrl: string | null
  authMode: AuthMode
  /** Names only. */
  extraHeaders: string[]
  model: string
  defaultModel: string
}
