import type { DocumentModel, Paragraph, Shape, Slide } from './types'

export interface CommentAnchor {
  slide: Slide
  /** The shape the comment is attached to, or null for a whole-slide comment. */
  shape: Shape | null
}

/** Where a comment lives: its slide and (if any) shape. Null only if the
 * slide no longer exists. */
export function findAnchorForComment(document: DocumentModel, commentId: string): CommentAnchor | null {
  const comment = document.comments[commentId]
  if (!comment) return null
  const slide = document.slides.find((s) => s.id === comment.slideId)
  if (!slide) return null
  const shape = comment.shapeId ? (slide.shapes.find((s) => s.id === comment.shapeId) ?? null) : null
  return { slide, shape }
}

/** Every paragraph of a shape: its own, or all table-cell paragraphs. */
export function shapeParagraphs(shape: Shape): Paragraph[] {
  if (shape.type === 'shape') return shape.paragraphs
  if (shape.type === 'table') return shape.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs))
  return []
}

export function paragraphText(p: Paragraph): string {
  return p.runs.map((r) => r.text ?? '').join('')
}

/** DOM id to scroll to for a comment: its shape if it has one, else its slide. */
export function scrollTargetForComment(document: DocumentModel, commentId: string): string | null {
  const anchor = findAnchorForComment(document, commentId)
  if (!anchor) return null
  return anchor.shape ? anchor.shape.id : anchor.slide.id
}

/** Black or white, whichever reads better on the given hex background. */
export function contrastingTextColor(hexBg: string): string {
  const h = hexBg.replace('#', '')
  if (h.length !== 6) return '000000'
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16))
  return 0.299 * r + 0.587 * g + 0.114 * b < 140 ? 'FFFFFF' : '000000'
}
