import type { BaseElement, Transform } from '@pptx/parser'
import type React from 'react'

/**
 * Returns a CSS `transform` string for a slide element's transform descriptor.
 *
 * Rotation origin is the element center (set via transform-origin: center in
 * the parent style). For flip + rotation, the order matters: scale then rotate.
 */
export function toCSSTransform(transform: Transform | undefined): string | undefined {
  if (!transform) return undefined
  const parts: string[] = []
  if (transform.flipH) parts.push('scaleX(-1)')
  if (transform.flipV) parts.push('scaleY(-1)')
  if (transform.rotation) parts.push(`rotate(${transform.rotation}deg)`)
  return parts.length ? parts.join(' ') : undefined
}

/**
 * Returns the absolute-positioned inline style for any slide element.
 * All sizes/positions are in CSS `pt` units, matching the parser output.
 */
export function elementStyle(el: BaseElement): React.CSSProperties {
  const transform = toCSSTransform(el.transform)
  return {
    position: 'absolute',
    left: `${el.position.x}pt`,
    top: `${el.position.y}pt`,
    width: `${el.size.width}pt`,
    height: `${el.size.height}pt`,
    ...(transform ? { transform, transformOrigin: 'center center' } : {}),
    ...(el.hidden ? { visibility: 'hidden' } : {}),
  }
}
