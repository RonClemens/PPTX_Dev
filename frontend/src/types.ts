export interface Run {
  id: string
  type: 'text' | 'break'
  text?: string
  /** Font size in points, with inheritance and autofit scaling already applied. */
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  /** Hex RGB without '#'. */
  color?: string
  /** True for a PowerPoint field run (slide number, date, ...). */
  field?: boolean
}

export interface Paragraph {
  id: string
  type: 'paragraph'
  level: number
  alignment?: 'center' | 'right' | 'justify'
  /** Bullet character ('#' = auto-numbered). */
  bullet?: string
  /** Left margin / first-line indent, in EMU. */
  marL?: number
  indent?: number
  lineSpacing?: number
  /** Space before, in points. */
  spaceBefore?: number
  /** Font size (pt) for an empty paragraph's line height. */
  defaultSize?: number
  runs: Run[]
}

export interface TableCell {
  paragraphs: Paragraph[]
  colSpan?: number
  rowSpan?: number
  hMerge?: boolean
  vMerge?: boolean
  fill?: string
}

export interface TableRow {
  h: number
  cells: TableCell[]
}

interface ShapeBase {
  id: string
  name: string
  descr: string
  /** Position and size in EMU (914400 per inch), in slide coordinates. */
  x: number
  y: number
  w: number
  h: number
  rot?: number
  background?: boolean
}

export interface TextShape extends ShapeBase {
  type: 'shape'
  placeholder?: string
  isTitle?: boolean
  fill?: string
  line?: { color: string; w: number }
  geom?: string
  paragraphs: Paragraph[]
  anchor?: 'top' | 'middle' | 'bottom'
  insets?: { l: number; t: number; r: number; b: number }
  nowrap?: boolean
}

export interface TableShape extends ShapeBase {
  type: 'table'
  cols: number[]
  rows: TableRow[]
  firstRow: boolean
  headerFill?: string
}

export interface PictureShape extends ShapeBase {
  type: 'picture'
  relId: string | null
}

export interface LineShape extends ShapeBase {
  type: 'line'
  line: { color: string; w: number }
}

export interface OtherShape extends ShapeBase {
  type: 'other'
  label: string
}

export type Shape = TextShape | TableShape | PictureShape | LineShape | OtherShape

export interface Slide {
  id: string
  /** 1-based slide number. */
  index: number
  title: string | null
  layout: string | null
  hidden: boolean
  background: string
  shapes: Shape[]
  backgroundShapes: Shape[]
  notes: string
}

export interface CommentEntry {
  id: string
  author: string
  initials: string | null
  date: string | null
  text: string
  done: boolean
  parentId: string | null
  slideId: string
  /** The shape the comment is attached to, or null for a whole-slide comment. */
  shapeId: string | null
  /** Text the reviewer selected when commenting, if any. */
  quote: string | null
  /** 'edit' = an automatic record of a text edit (PowerPoint has no Track Changes). */
  kind: 'comment' | 'edit'
  pos: { x: number; y: number }
}

export interface SlideSize {
  widthEmu: number
  heightEmu: number
  widthIn: number
  heightIn: number
}

export interface DocumentModel {
  slides: Slide[]
  comments: Record<string, CommentEntry>
  slideSize: SlideSize
}

export type ViewMode = 'slides' | 'grid' | 'outline'

export interface Chunk {
  id: string
  heading: string | null
  level: number
  startIndex: number
  endIndex: number
  markdown: string
  itemIds: string[]
  slideId: string
}

export interface DocumentMeta {
  id: string
  filename: string
  title: string
  created_at: string
  updated_at: string
  doc_ref?: string | null
  rev?: string | null
}

export type DecisionValue = 'accept' | 'reject' | 'info_only' | 'defer'

export interface DecisionRecord {
  decision: DecisionValue
  reason: string
  ref: string
  author: string
  decidedAt: string
}

export interface ChatTurn {
  role: 'user' | 'assistant'
  content: string
  ts: string
  author?: string
}

export interface Adjudication {
  action: 'edit' | 'reply' | 'no_change'
  reply: string
  reasoning: string
  original_text: string | null
  replacement_text: string | null
  change_id: string | null
  reply_comment_id: string | null
}

export interface CommentRefGuess {
  ref: string | null
  para: string | null
}

export interface DocumentPayload {
  meta: DocumentMeta
  document: DocumentModel
  chunks: Chunk[]
  decisions: Record<string, DecisionRecord>
  chats: Record<string, ChatTurn[]>
  commentRefs: Record<string, CommentRefGuess>
  adjudication?: Adjudication
  aiReview?: { added: number; skipped: number; commentIds: string[] }
  decision?: { comment_id: string } & DecisionRecord
  chat?: { comment_id: string; history: ChatTurn[] }
  xlsxImport?: XlsxImportResult
}

export interface XlsxImportResult {
  applied: { row: number; comment_id: string; decision: DecisionValue; created: boolean }[]
  skipped: { row: number; reason: string }[]
  skippedSummary: { reason: string; count: number; rows: number[] }[]
}

export type AppMode = 'review' | 'adjudicate'

export interface AdjudicationSuggestion {
  action: 'edit' | 'reply' | 'no_change'
  reply: string
  reasoning: string
  original_text: string | null
  replacement_text: string | null
}

export interface SuggestionError {
  error: string
}

export interface SuggestAdjudicationsResponse {
  suggestions: Record<string, AdjudicationSuggestion | SuggestionError>
  usedFullDocumentContext: boolean
}
