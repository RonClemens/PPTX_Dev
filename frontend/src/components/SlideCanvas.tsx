import type { CSSProperties } from 'react'
import { api } from '../api'
import { contrastingTextColor } from '../documentUtils'
import type { CommentEntry, Paragraph, Run, Shape, Slide, SlideSize, TableShape, TextShape } from '../types'

/**
 * Renders one slide as absolutely-positioned shapes inside a box that keeps
 * the slide's aspect ratio. Every length is expressed relative to the slide:
 * positions/sizes as percentages, font sizes and insets in `cqw` (1% of the
 * slide's own width, via CSS container queries), so the slide scales
 * perfectly to any width -- full size in the main view, thumbnail in the
 * grid -- with no JavaScript measuring or transforms.
 */

interface SlideMetrics {
  widthEmu: number
  heightEmu: number
  widthPt: number
}

function metricsOf(size: SlideSize): SlideMetrics {
  return { widthEmu: size.widthEmu, heightEmu: size.heightEmu, widthPt: size.widthIn * 72 }
}

const pct = (v: number, total: number) => `${(v / total) * 100}%`
const cqwFromPt = (pt: number, m: SlideMetrics) => `${(pt / m.widthPt) * 100}cqw`
const cqwFromEmu = (emu: number, m: SlideMetrics) => `${(emu / m.widthEmu) * 100}cqw`

interface CommonProps {
  docId: string
  slide: Slide
  metrics: SlideMetrics
  defaultColor: string
  comments: CommentEntry[]
  selectedCommentId: string | null
  onSelectComment: (id: string) => void
  /** Thumbnail / preview mode: nothing is selectable or clickable. */
  inert: boolean
}

function RunView({ run, m, defaultColor }: { run: Run; m: SlideMetrics; defaultColor: string }) {
  if (run.type === 'break') return <br />
  const style: CSSProperties = {
    color: `#${run.color ?? defaultColor}`,
  }
  if (run.size) style.fontSize = cqwFromPt(run.size, m)
  if (run.bold) style.fontWeight = 700
  if (run.italic) style.fontStyle = 'italic'
  if (run.underline) style.textDecoration = 'underline'
  return (
    <span className="slide-run" data-run-id={run.id} style={style}>
      {run.text}
    </span>
  )
}

function ParagraphView({
  p,
  m,
  defaultColor,
}: {
  p: Paragraph
  m: SlideMetrics
  defaultColor: string
}) {
  const style: CSSProperties = {}
  if (p.alignment) style.textAlign = p.alignment
  if (p.lineSpacing) style.lineHeight = 1.2 * p.lineSpacing
  if (p.spaceBefore) style.marginTop = cqwFromPt(p.spaceBefore, m)
  const marL = p.marL ?? 0
  const indent = p.indent ?? 0
  if (marL) style.paddingLeft = cqwFromEmu(marL, m)
  if (indent) style.textIndent = cqwFromEmu(indent, m)
  if (p.bullet) {
    // Hanging bullet: the ::before box (not selectable text, so it never
    // leaks into a text selection) is as wide as the hanging indent.
    ;(style as Record<string, string>)['--bullet-width'] = cqwFromEmu(Math.abs(indent) || 342900, m)
  }
  const size = p.runs.find((r) => r.size)?.size ?? p.defaultSize
  // The paragraph's own font size sets its line box "strut"; without it a
  // line of small text would still be as tall as the browser default.
  if (size) style.fontSize = cqwFromPt(size, m)
  const empty = p.runs.length === 0 || p.runs.every((r) => r.type === 'text' && !r.text)
  return (
    <div
      className="slide-para"
      data-para-id={p.id}
      data-bullet={p.bullet ? (p.bullet === '#' ? '•' : p.bullet) : undefined}
      style={style}
    >
      {empty ? (
        <span>&#8203;</span>
      ) : (
        p.runs.map((r) => <RunView key={r.id} run={r} m={m} defaultColor={defaultColor} />)
      )}
    </div>
  )
}

function boxStyle(shape: Shape, m: SlideMetrics): CSSProperties {
  const style: CSSProperties = {
    left: pct(shape.x, m.widthEmu),
    top: pct(shape.y, m.heightEmu),
    width: pct(shape.w, m.widthEmu),
    height: pct(shape.h, m.heightEmu),
  }
  if (shape.rot) style.transform = `rotate(${shape.rot}deg)`
  return style
}

function TextShapeView({ shape, ...c }: { shape: TextShape } & CommonProps) {
  const m = c.metrics
  const style: CSSProperties = {
    ...boxStyle(shape, m),
    justifyContent: shape.anchor === 'middle' ? 'center' : shape.anchor === 'bottom' ? 'flex-end' : 'flex-start',
  }
  if (shape.fill) style.background = `#${shape.fill}`
  if (shape.line) style.border = `${Math.max(1, shape.line.w / 12700)}px solid #${shape.line.color}`
  if (shape.geom === 'ellipse') style.borderRadius = '50%'
  else if (shape.geom === 'roundRect') style.borderRadius = '8%'
  if (shape.insets) {
    style.padding = `${cqwFromEmu(shape.insets.t, m)} ${cqwFromEmu(shape.insets.r, m)} ${cqwFromEmu(shape.insets.b, m)} ${cqwFromEmu(shape.insets.l, m)}`
  }
  // A filled shape is its own background; otherwise pick a text color that
  // reads on the slide background.
  const defaultColor = shape.fill ? contrastingTextColor(shape.fill) : c.defaultColor
  return (
    <ShapeFrame shape={shape} style={style} {...c}>
      {shape.paragraphs.map((p) => (
        <ParagraphView key={p.id} p={p} m={m} defaultColor={defaultColor} />
      ))}
    </ShapeFrame>
  )
}

function TableShapeView({ shape, ...c }: { shape: TableShape } & CommonProps) {
  const m = c.metrics
  const totalW = shape.cols.reduce((a, b) => a + b, 0) || 1
  const totalH = shape.rows.reduce((a, r) => a + r.h, 0) || 1
  return (
    <ShapeFrame shape={shape} style={boxStyle(shape, m)} {...c}>
      <table className="slide-table">
        <colgroup>
          {shape.cols.map((w, i) => (
            <col key={i} style={{ width: `${(w / totalW) * 100}%` }} />
          ))}
        </colgroup>
        <tbody>
          {shape.rows.map((row, ri) => {
            const header = shape.firstRow && ri === 0
            return (
              <tr key={ri} style={{ height: `${(row.h / totalH) * 100}%` }}>
                {row.cells.map((cell, ci) => {
                  if (cell.hMerge || cell.vMerge) return null
                  const fill = cell.fill ?? (header ? shape.headerFill : undefined)
                  const color = fill ? contrastingTextColor(fill) : c.defaultColor
                  return (
                    <td
                      key={ci}
                      colSpan={cell.colSpan}
                      rowSpan={cell.rowSpan}
                      style={{
                        background: fill ? `#${fill}` : undefined,
                        fontWeight: header ? 700 : undefined,
                        padding: `${cqwFromEmu(45720, m)} ${cqwFromEmu(91440, m)}`,
                      }}
                    >
                      {cell.paragraphs.map((p) => (
                        <ParagraphView key={p.id} p={p} m={m} defaultColor={color} />
                      ))}
                    </td>
                  )
                })}
              </tr>
            )
          })}
        </tbody>
      </table>
    </ShapeFrame>
  )
}

/** Wrapper that carries the shape id (scroll target), the commented-shape
 * outline, and click-to-select-comment behavior. */
function ShapeFrame({
  shape,
  style,
  children,
  slide,
  comments,
  selectedCommentId,
  onSelectComment,
  inert,
}: {
  shape: Shape
  style: CSSProperties
  children?: React.ReactNode
} & CommonProps) {
  const mine = comments.filter((cm) => cm.shapeId === shape.id)
  const selected = mine.some((cm) => cm.id === selectedCommentId)
  const classes = [
    'slide-shape',
    `slide-shape-${shape.type}`,
    mine.length > 0 ? 'slide-shape-commented' : '',
    selected ? 'slide-shape-commented-selected' : '',
  ]
    .filter(Boolean)
    .join(' ')
  return (
    <div
      className={classes}
      data-shape-id={shape.id}
      data-slide-id={slide.id}
      style={style}
      onClick={() => {
        if (!inert && mine.length > 0) onSelectComment(mine[0].id)
      }}
    >
      {children}
    </div>
  )
}

function PictureView({ shape, ...c }: { shape: Extract<Shape, { type: 'picture' }> } & CommonProps) {
  return (
    <ShapeFrame shape={shape} style={boxStyle(shape, c.metrics)} {...c}>
      {shape.relId && (
        <img
          className="slide-picture"
          src={api.mediaUrl(c.docId, c.slide.id, shape.relId)}
          alt={shape.descr || shape.name || 'picture'}
          draggable={false}
        />
      )}
    </ShapeFrame>
  )
}

function BackgroundShapeView({ shape, m }: { shape: Shape; m: SlideMetrics }) {
  if (shape.type === 'shape') {
    const style: CSSProperties = boxStyle(shape, m)
    if (shape.fill) style.background = `#${shape.fill}`
    if (shape.line) style.border = `${Math.max(1, shape.line.w / 12700)}px solid #${shape.line.color}`
    if (shape.geom === 'ellipse') style.borderRadius = '50%'
    return <div className="slide-bg-shape" style={style} />
  }
  if (shape.type === 'line') return <LineSvg shape={shape} m={m} />
  return null
}

function LineSvg({ shape, m }: { shape: Extract<Shape, { type: 'line' }>; m: SlideMetrics }) {
  return (
    <svg className="slide-line" style={boxStyle(shape, m)} viewBox="0 0 100 100" preserveAspectRatio="none">
      <line
        x1="0"
        y1="0"
        x2={shape.w >= shape.h ? 100 : 0}
        y2={shape.h > shape.w ? 100 : 0}
        stroke={`#${shape.line.color}`}
        strokeWidth={Math.max(1, shape.line.w / 12700)}
        vectorEffect="non-scaling-stroke"
      />
    </svg>
  )
}

interface Props {
  docId: string
  slide: Slide
  slideSize: SlideSize
  comments: Record<string, CommentEntry>
  selectedCommentId: string | null
  onSelectComment: (id: string) => void
  /** Show comment pins on the slide. */
  showPins?: boolean
  /** Thumbnail / preview: no selection, clicks or pins interaction. */
  inert?: boolean
  /** Extra emphasis for one shape (e.g. the shape a queued comment is on). */
  highlightShapeId?: string | null
}

export default function SlideCanvas({
  docId,
  slide,
  slideSize,
  comments,
  selectedCommentId,
  onSelectComment,
  showPins = true,
  inert = false,
  highlightShapeId = null,
}: Props) {
  const m = metricsOf(slideSize)
  const slideComments = Object.values(comments).filter((c) => c.slideId === slide.id)
  const common: CommonProps = {
    docId,
    slide,
    metrics: m,
    defaultColor: contrastingTextColor(slide.background),
    comments: slideComments,
    selectedCommentId,
    onSelectComment,
    inert,
  }
  const pins = slideComments.filter((c) => !c.parentId && c.kind !== 'edit')

  return (
    <div
      className={`slide-canvas${inert ? ' slide-canvas-inert' : ''}`}
      data-slide-id={slide.id}
      style={{
        aspectRatio: `${slideSize.widthEmu} / ${slideSize.heightEmu}`,
        background: `#${slide.background}`,
      }}
    >
      {slide.backgroundShapes.map((s) => (
        <BackgroundShapeView key={s.id} shape={s} m={m} />
      ))}
      {slide.shapes.map((shape) => {
        switch (shape.type) {
          case 'shape':
            return <TextShapeView key={shape.id} shape={shape} {...common} />
          case 'table':
            return <TableShapeView key={shape.id} shape={shape} {...common} />
          case 'picture':
            return <PictureView key={shape.id} shape={shape} {...common} />
          case 'line':
            return <LineSvg key={shape.id} shape={shape} m={m} />
          case 'other':
            return (
              <ShapeFrame key={shape.id} shape={shape} style={boxStyle(shape, m)} {...common}>
                <div className="slide-other">{shape.label}</div>
              </ShapeFrame>
            )
        }
      })}
      {highlightShapeId &&
        slide.shapes
          .filter((s) => s.id === highlightShapeId)
          .map((s) => <div key="hl" className="slide-highlight" style={boxStyle(s, m)} />)}
      {showPins &&
        !inert &&
        pins.map((c, i) => (
          <button
            key={c.id}
            className={[
              'comment-pin',
              c.done ? 'comment-pin-done' : '',
              c.id === selectedCommentId ? 'comment-pin-selected' : '',
            ]
              .filter(Boolean)
              .join(' ')}
            style={{
              left: `calc(${pct(c.pos.x, m.widthEmu)} + ${i * 1.6}cqw)`,
              top: pct(c.pos.y, m.heightEmu),
            }}
            title={`${c.author}: ${c.text}`}
            onClick={(e) => {
              e.stopPropagation()
              onSelectComment(c.id)
            }}
          >
            {i + 1}
          </button>
        ))}
    </div>
  )
}
