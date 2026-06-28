import type { Color, Fill, Stroke, ThemeColors } from '@pptx/parser'
import { resolveColor } from '@pptx/parser'

export function toCSS(color: Color | undefined, theme: ThemeColors): string {
  if (!color) return 'transparent'
  return resolveColor(color, theme)
}

export function fillToCSS(fill: Fill | undefined, theme: ThemeColors): string {
  if (!fill || fill.type === 'none') return 'transparent'
  if (fill.type === 'solid') return toCSS(fill.color, theme)
  if (fill.type === 'gradient') {
    const stops = fill.stops
      .map((s) => `${toCSS(s.color, theme)} ${(s.position * 100).toFixed(1)}%`)
      .join(', ')
    const angle = fill.angle ?? 0
    // CSS gradient angle: 0° = bottom-to-top; OOXML: 0° = left-to-right
    const cssAngle = (90 - angle + 360) % 360
    return `linear-gradient(${cssAngle}deg, ${stops})`
  }
  // pattern falls back to foreground color
  if (fill.type === 'pattern' && fill.fgColor) return toCSS(fill.fgColor, theme)
  return 'transparent'
}

/** SVG fill attribute value */
export function fillToSVG(fill: Fill | undefined, theme: ThemeColors): string {
  if (!fill || fill.type === 'none') return 'none'
  if (fill.type === 'solid') return toCSS(fill.color, theme)
  // SVG gradient would require defs — fall back to first stop for now
  if (fill.type === 'gradient' && fill.stops.length > 0) {
    return toCSS(fill.stops[0]!.color, theme)
  }
  if (fill.type === 'pattern' && fill.fgColor) return toCSS(fill.fgColor, theme)
  return 'none'
}

/** SVG stroke attributes */
export function strokeToSVGAttrs(stroke: Stroke | undefined, theme: ThemeColors): {
  stroke: string
  strokeWidth: string
  strokeDasharray?: string
} {
  if (!stroke || stroke.fill.type === 'none') {
    return { stroke: 'none', strokeWidth: '0' }
  }

  const color = fillToSVG(stroke.fill, theme)
  const width = stroke.width != null ? `${stroke.width}pt` : '1pt'

  let strokeDasharray: string | undefined
  switch (stroke.dashStyle) {
    case 'dot':
      strokeDasharray = '2 2'
      break
    case 'dash':
      strokeDasharray = '6 2'
      break
    case 'dashDot':
      strokeDasharray = '6 2 2 2'
      break
    case 'lgDash':
      strokeDasharray = '12 4'
      break
    case 'lgDashDot':
      strokeDasharray = '12 4 2 4'
      break
  }

  return { stroke: color, strokeWidth: width, strokeDasharray }
}
