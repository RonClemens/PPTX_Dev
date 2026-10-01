import { useState } from 'react'
import { paragraphText } from '../documentUtils'
import type { Shape, Slide } from '../types'

interface Props {
  slides: Slide[]
  onJumpToSlide: (slideId: string) => void
}

function shapeLines(shape: Shape): { text: string; level: number }[] {
  if (shape.type === 'shape') {
    return shape.paragraphs
      .map((p) => ({ text: paragraphText(p), level: p.level }))
      .filter((l) => l.text.trim())
  }
  if (shape.type === 'table') {
    return shape.rows.map((row) => ({
      text: row.cells.map((c) => c.paragraphs.map(paragraphText).join(' ').trim()).join('  |  '),
      level: 0,
    }))
  }
  if (shape.type === 'picture') return [{ text: `[Picture: ${shape.descr || shape.name}]`, level: 0 }]
  if (shape.type === 'other') return [{ text: `[${shape.label}]`, level: 0 }]
  return []
}

/** PowerPoint-style outline: one collapsible node per slide (number + title),
 * its text underneath. Read/navigate only -- clicking a title jumps to that
 * slide in the Slides view. */
export default function OutlineView({ slides, onJumpToSlide }: Props) {
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({})

  return (
    <div className="outline-view">
      {slides.map((slide) => {
        const isCollapsed = collapsed[slide.id]
        const lines = slide.shapes.filter((s) => !(s.type === 'shape' && s.isTitle)).flatMap(shapeLines)
        return (
          <div className="outline-node" key={slide.id}>
            <div className="outline-heading-row">
              <button
                className="outline-toggle"
                disabled={lines.length === 0}
                onClick={() => setCollapsed((c) => ({ ...c, [slide.id]: !c[slide.id] }))}
                title={isCollapsed ? 'Expand' : 'Collapse'}
              >
                {lines.length === 0 ? '·' : isCollapsed ? '▸' : '▾'}
              </button>
              <span className="outline-slide-number">{slide.index}</span>
              <span className="outline-heading-text outline-level-1" onClick={() => onJumpToSlide(slide.id)}>
                {slide.title || '(untitled slide)'}
              </span>
              {slide.hidden && <span className="outline-hidden-badge">hidden</span>}
            </div>
            {!isCollapsed && lines.length > 0 && (
              <div className="outline-children">
                {lines.map((l, i) => (
                  <div
                    key={i}
                    className="outline-body-item"
                    style={{ paddingLeft: 44 + l.level * 18 }}
                    onClick={() => onJumpToSlide(slide.id)}
                  >
                    {l.text}
                  </div>
                ))}
              </div>
            )}
          </div>
        )
      })}
    </div>
  )
}
