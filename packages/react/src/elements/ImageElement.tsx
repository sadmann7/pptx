import React from 'react'
import type { ImageShape, ThemeColors } from '@pptx/parser'
import { elementStyle } from '../render/transform'

export function ImageElement({ element }: { element: ImageShape; theme: ThemeColors }) {
  if (!element.src) return null

  const outer: React.CSSProperties = {
    ...elementStyle(element),
    overflow: 'hidden',
  }

  const imgStyle: React.CSSProperties = {
    width: '100%',
    height: '100%',
    objectFit: 'cover',
    display: 'block',
  }

  // Crop via object-position / clip-path if cropRect is set
  if (element.cropRect) {
    const { top, right, bottom, left } = element.cropRect
    imgStyle.objectPosition = `${(-left / (1 - left - right)) * 100}% ${(-top / (1 - top - bottom)) * 100}%`
    imgStyle.objectFit = 'none'
    // Scale the image to fill the cropped area
    const scaleX = 1 / (1 - left - right)
    const scaleY = 1 / (1 - top - bottom)
    imgStyle.width = `${scaleX * 100}%`
    imgStyle.height = `${scaleY * 100}%`
    imgStyle.marginLeft = `${-left * scaleX * 100}%`
    imgStyle.marginTop = `${-top * scaleY * 100}%`
  }

  return (
    <div style={outer} data-element-type="image" data-element-id={element.id}>
      <img
        src={element.src}
        alt={element.name ?? ''}
        style={imgStyle}
        loading="lazy"
        decoding="async"
      />
    </div>
  )
}
