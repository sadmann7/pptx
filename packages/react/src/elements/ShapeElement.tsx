import React from 'react'
import type { GeometricShape, ThemeColors } from '@pptx/parser'
import { fillToSVG, strokeToSVGAttrs } from '../render/color'
import { bodyStyle } from '../render/text'
import { getShapePath } from '../render/shapes'
import { elementStyle } from '../render/transform'
import { ParagraphElement } from './shared/ParagraphElement'

interface ShapeElementProps {
  element: GeometricShape
  theme: ThemeColors
}

export function ShapeElement({ element, theme }: ShapeElementProps) {
  const outer: React.CSSProperties = {
    ...elementStyle(element),
    position: 'absolute',
  }

  const shape = getShapePath(element.shapeType)
  const fill = fillToSVG(element.fill, theme)
  const strokeAttrs = strokeToSVGAttrs(element.stroke, theme)

  // Shapes that are just lines / connectors have no fill
  const isLine = element.shapeType === 'line' || element.shapeType.includes('Connector')

  const svgStyle: React.CSSProperties = {
    position: 'absolute',
    inset: 0,
    width: '100%',
    height: '100%',
    overflow: 'visible',
  }

  const sharedProps = {
    fill: isLine ? 'none' : fill,
    stroke: strokeAttrs.stroke,
    strokeWidth: strokeAttrs.strokeWidth,
    ...(strokeAttrs.strokeDasharray ? { strokeDasharray: strokeAttrs.strokeDasharray } : {}),
  }

  return (
    <div style={outer} data-element-type="shape" data-element-id={element.id}>
      <svg
        style={svgStyle}
        viewBox="0 0 100 100"
        preserveAspectRatio="none"
        aria-hidden="true"
      >
        {renderShapeElement(shape, sharedProps)}
      </svg>

      {element.body && (
        <div
          style={{
            position: 'absolute',
            inset: 0,
            ...bodyStyle(element.body.properties, theme),
          }}
        >
          {element.body.paragraphs.map((p, i) => (
            <ParagraphElement key={i} paragraph={p} theme={theme} />
          ))}
        </div>
      )}
    </div>
  )
}

function renderShapeElement(
  shape: ReturnType<typeof getShapePath>,
  sharedProps: Record<string, string>,
): React.ReactElement {
  const allProps = { ...shape.attrs, ...sharedProps }

  switch (shape.element) {
    case 'rect':
      return (
        <rect
          {...allProps}
          rx={shape.rx}
          ry={shape.ry}
        />
      )
    case 'ellipse':
      return <ellipse {...allProps} />
    case 'polygon':
      return <polygon {...allProps} />
    case 'path':
      return <path {...allProps} />
    case 'line':
      return <line {...allProps} />
  }
}
