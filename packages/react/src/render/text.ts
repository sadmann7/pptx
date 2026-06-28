import type {
  BodyProperties,
  ParagraphStyle,
  RunStyle,
  ThemeColors,
} from '@pptx/parser'
import type React from 'react'
import { toCSS } from './color'

// ─── Body / container ────────────────────────────────────────────────────────

export function bodyStyle(
  props: BodyProperties,
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  _theme: ThemeColors,
): React.CSSProperties {
  const css: React.CSSProperties = {
    overflow: 'hidden',
    width: '100%',
    height: '100%',
    boxSizing: 'border-box',
    display: 'flex',
    flexDirection: 'column',
  }

  if (props.insetLeft != null) css.paddingLeft = `${props.insetLeft}pt`
  if (props.insetRight != null) css.paddingRight = `${props.insetRight}pt`
  if (props.insetTop != null) css.paddingTop = `${props.insetTop}pt`
  if (props.insetBottom != null) css.paddingBottom = `${props.insetBottom}pt`

  // Vertical alignment via flexbox
  switch (props.verticalAlignment) {
    case 'mid':
      css.justifyContent = 'center'
      break
    case 'bottom':
    case 'just':
      css.justifyContent = 'flex-end'
      break
    default:
      css.justifyContent = 'flex-start'
      break
  }

  if (props.wrap === 'none') css.whiteSpace = 'nowrap'

  if (props.columns && props.columns > 1) {
    css.columnCount = props.columns
    if (props.columnSpacing != null) {
      css.columnGap = `${props.columnSpacing}pt`
    }
  }

  return css
}

// ─── Paragraph ────────────────────────────────────────────────────────────────

export function paragraphStyle(
  style: ParagraphStyle,
  _theme: ThemeColors,
): React.CSSProperties {
  const css: React.CSSProperties = {
    margin: 0,
    padding: 0,
    lineHeight: 'inherit',
  }

  if (style.alignment) {
    css.textAlign = style.alignment as React.CSSProperties['textAlign']
  }

  if (style.marginLeft != null) css.marginLeft = `${style.marginLeft}pt`
  if (style.marginRight != null) css.marginRight = `${style.marginRight}pt`
  if (style.indent != null) css.textIndent = `${style.indent}pt`

  if (style.spaceBefore) {
    css.marginTop =
      style.spaceBefore.unit === 'pt'
        ? `${style.spaceBefore.value}pt`
        : `${style.spaceBefore.value}%`
  }

  if (style.spaceAfter) {
    css.marginBottom =
      style.spaceAfter.unit === 'pt'
        ? `${style.spaceAfter.value}pt`
        : `${style.spaceAfter.value}%`
  }

  if (style.lineSpacing) {
    css.lineHeight =
      style.lineSpacing.unit === 'pct'
        ? `${style.lineSpacing.value}%`
        : `${style.lineSpacing.value}pt`
  }

  return css
}

// ─── Run ─────────────────────────────────────────────────────────────────────

export function runStyle(
  style: RunStyle,
  theme: ThemeColors,
  parentFontSize?: number,
): React.CSSProperties {
  const css: React.CSSProperties = {}

  if (style.bold) css.fontWeight = 'bold'
  if (style.italic) css.fontStyle = 'italic'
  if (style.fontSize != null) css.fontSize = `${style.fontSize}pt`

  if (style.color) css.color = toCSS(style.color, theme)

  if (style.fontFamily) {
    css.fontFamily = `"${style.fontFamily}", sans-serif`
  }

  // Text decorations
  if (style.underline && style.underline !== 'none') {
    css.textDecoration = style.strikethrough ? 'underline line-through' : 'underline'
  } else if (style.strikethrough) {
    css.textDecoration = 'line-through'
  }

  // Superscript / subscript via baseline
  if (style.baseline != null) {
    if (style.baseline > 0) {
      css.verticalAlign = 'super'
      css.fontSize = `${((style.fontSize ?? parentFontSize ?? 12) * 0.65).toFixed(1)}pt`
    } else if (style.baseline < 0) {
      css.verticalAlign = 'sub'
      css.fontSize = `${((style.fontSize ?? parentFontSize ?? 12) * 0.65).toFixed(1)}pt`
    }
  }

  return css
}
