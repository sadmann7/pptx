import React from 'react'
import { useSlide } from '../context'

export interface NotesProps {
  className?: string
  style?: React.CSSProperties
  /** Placeholder text shown when the current slide has no notes */
  emptyText?: string
}

/**
 * <Presentation.Notes>
 *
 * Renders the speaker notes for the current slide as plain text.
 * Notes are extracted from the notesSlide XML during parsing.
 */
export function Notes({ className, style, emptyText = 'No notes for this slide.' }: NotesProps) {
  const { slide } = useSlide()

  const containerStyle: React.CSSProperties = {
    padding: '12px 16px',
    fontFamily: 'sans-serif',
    fontSize: '13px',
    lineHeight: '1.6',
    color: '#374151',
    whiteSpace: 'pre-wrap',
    overflow: 'auto',
    ...style,
  }

  const notes = slide?.notes

  return (
    <div className={className} style={containerStyle} role="note">
      {notes ? notes : <span style={{ color: '#9ca3af' }}>{emptyText}</span>}
    </div>
  )
}
