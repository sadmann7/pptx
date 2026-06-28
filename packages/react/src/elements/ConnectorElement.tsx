import React from 'react'
import type { ConnectorShape, ThemeColors } from '@pptx/parser'
import { strokeToSVGAttrs } from '../render/color'
import { elementStyle } from '../render/transform'

interface ConnectorElementProps {
  element: ConnectorShape
  theme: ThemeColors
}

export function ConnectorElement({ element, theme }: ConnectorElementProps) {
  const outer: React.CSSProperties = {
    ...elementStyle(element),
    pointerEvents: 'none',
  }

  const strokeAttrs = strokeToSVGAttrs(element.stroke, theme)

  return (
    <div style={outer} data-element-type="connector" data-element-id={element.id}>
      <svg
        style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', overflow: 'visible' }}
        viewBox="0 0 100 100"
        preserveAspectRatio="none"
        aria-hidden="true"
      >
        <line
          x1={0}
          y1={0}
          x2={100}
          y2={100}
          stroke={strokeAttrs.stroke !== 'none' ? strokeAttrs.stroke : '#000'}
          strokeWidth={strokeAttrs.strokeWidth}
          {...(strokeAttrs.strokeDasharray ? { strokeDasharray: strokeAttrs.strokeDasharray } : {})}
        />
      </svg>
    </div>
  )
}
