import React from 'react'
import type { Slide, ThemeColors } from '@pptx/parser'
import { usePresentation, useSlide } from '../context'
import { fillToCSS, toCSS } from '../render/color'
import { SlideElementRenderer } from '../elements/index'

export interface ThumbnailsProps {
  className?: string
  style?: React.CSSProperties
  /** CSS thumbnail width. Height is computed from the slide aspect ratio. @default '120px' */
  thumbnailWidth?: number | string
  /** Called when the user clicks a thumbnail */
  onSlideClick?: (index: number) => void
  /** Render a custom selected indicator */
  renderSelected?: (index: number) => React.ReactNode
}

/**
 * <Presentation.Thumbnails>
 *
 * A vertical filmstrip of slide thumbnails. Each thumbnail is a scaled-down
 * version of the actual slide — no canvas, just CSS scale().
 *
 * Clicking a thumbnail calls `goTo(index)` on the store.
 */
export function Thumbnails({
  className,
  style,
  thumbnailWidth = 120,
  onSlideClick,
  renderSelected,
}: ThumbnailsProps) {
  const { presentation, status } = usePresentation()
  const { index: currentIndex, goTo } = useSlide()

  if (status === 'loading') {
    return (
      <div className={className} style={{ padding: 8, fontFamily: 'sans-serif', fontSize: 12, color: '#9ca3af', ...style }}>
        Loading…
      </div>
    )
  }

  if (status !== 'ready' || !presentation) return null

  const { width: slideW, height: slideH } = presentation.slideSize
  const tw = typeof thumbnailWidth === 'number' ? thumbnailWidth : Number.parseInt(thumbnailWidth)
  const th = Math.round(tw * (slideH / slideW))
  const scale = tw / slideW

  const containerStyle: React.CSSProperties = {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    overflowY: 'auto',
    padding: '8px',
    ...style,
  }

  const handleClick = (idx: number) => {
    goTo(idx)
    onSlideClick?.(idx)
  }

  return (
    <div className={className} style={containerStyle} role="tablist" aria-label="Slide thumbnails">
      {presentation.slides.map((slide) => (
        <ThumbnailItem
          key={slide.index}
          slide={slide}
          themeColors={presentation.theme.colors}
          isActive={slide.index === currentIndex}
          width={tw}
          height={th}
          scale={scale}
          slideWidth={slideW}
          slideHeight={slideH}
          onClick={() => handleClick(slide.index)}
          renderSelected={renderSelected}
        />
      ))}
    </div>
  )
}

interface ThumbnailItemProps {
  slide: Slide
  themeColors: ThemeColors
  isActive: boolean
  width: number
  height: number
  scale: number
  slideWidth: number
  slideHeight: number
  onClick: () => void
  renderSelected?: (index: number) => React.ReactNode
}

function ThumbnailItem({
  slide,
  themeColors,
  isActive,
  width,
  height,
  scale,
  slideWidth,
  slideHeight,
  onClick,
  renderSelected,
}: ThumbnailItemProps) {
  const bg = slide.background ? fillToCSS(slide.background.fill, themeColors) : '#ffffff'

  const wrapperStyle: React.CSSProperties = {
    position: 'relative',
    width: `${width}px`,
    height: `${height}px`,
    flexShrink: 0,
    cursor: 'pointer',
    outline: isActive ? '2px solid #3b82f6' : '1px solid #e5e7eb',
    outlineOffset: '1px',
    overflow: 'hidden',
    background: bg,
  }

  const defaultTextColor = toCSS({ type: 'scheme', token: 'dk1' }, themeColors)

  const canvasStyle: React.CSSProperties = {
    position: 'absolute',
    top: 0,
    left: 0,
    width: `${slideWidth}pt`,
    height: `${slideHeight}pt`,
    transformOrigin: 'top left',
    transform: `scale(${scale})`,
    color: defaultTextColor,
    fontFamily: 'sans-serif',
    pointerEvents: 'none',
    userSelect: 'none',
  }

  return (
    <button
      type="button"
      role="tab"
      aria-selected={isActive}
      aria-label={`Slide ${slide.index + 1}`}
      style={wrapperStyle}
      onClick={onClick}
    >
      <div style={canvasStyle}>
        {slide.elements.map((el) => (
          <SlideElementRenderer key={el.id} element={el} theme={themeColors} />
        ))}
      </div>
      {renderSelected && isActive && renderSelected(slide.index)}
    </button>
  )
}
