import type { Paragraph, ParagraphContent, ThemeColors } from '@pptx/parser'
import { toCSS } from '../../render/color'
import { paragraphStyle, runStyle } from '../../render/text'

interface ParagraphElementProps {
  paragraph: Paragraph
  theme: ThemeColors
}

export function ParagraphElement({ paragraph, theme }: ParagraphElementProps) {
  const style = paragraphStyle(paragraph.style, theme)
  const hasBullet =
    paragraph.style.bullet &&
    paragraph.style.bullet.type !== 'none' &&
    paragraph.runs.some((r) => r.type === 'run' && r.text !== '')

  if (!hasBullet) {
    return (
      <p style={style}>
        {paragraph.runs.map((run, i) => (
          <RunElement key={i} run={run} theme={theme} />
        ))}
      </p>
    )
  }

  const bullet = paragraph.style.bullet!
  const bulletChar =
    bullet.type === 'char'
      ? bullet.char
      : bullet.type === 'auto'
        ? (bullet.char ?? '•')
        : '•'
  const bulletColor =
    bullet.type !== 'none' && 'color' in bullet && bullet.color
      ? toCSS(bullet.color, theme)
      : undefined

  return (
    <p style={{ ...style, display: 'flex', gap: '4pt' }}>
      <span style={{ flexShrink: 0, color: bulletColor }}>{bulletChar}</span>
      <span style={{ flex: 1 }}>
        {paragraph.runs.map((run, i) => (
          <RunElement key={i} run={run} theme={theme} />
        ))}
      </span>
    </p>
  )
}

function RunElement({
  run,
  theme,
}: {
  run: ParagraphContent
  theme: ThemeColors
}) {
  if (run.type === 'lineBreak') return <br />

  const style = runStyle(run.style, theme)

  if (run.type === 'field') {
    return <span style={style}>{run.text}</span>
  }

  if (!run.text) return null

  if (run.style.link) {
    return (
      <a href={run.style.link} style={style} target="_blank" rel="noopener noreferrer">
        {run.text}
      </a>
    )
  }

  return <span style={style}>{run.text}</span>
}
