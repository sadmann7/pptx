import React from 'react'
import type { SlideElement, ThemeColors } from '@pptx/parser'
import { TextElement } from './TextElement'
import { ImageElement } from './ImageElement'
import { ShapeElement } from './ShapeElement'
import { TableElement } from './TableElement'
import { ConnectorElement } from './ConnectorElement'
import { GroupElement } from './GroupElement'
import { ChartElement } from './ChartElement'

export type ElementRendererFn = (element: SlideElement, theme: ThemeColors) => React.ReactNode

interface SlideElementRendererProps {
  element: SlideElement
  theme: ThemeColors
  /** Override the default renderer for a given element type or id */
  renderElement?: ElementRendererFn
}

export function SlideElementRenderer({
  element,
  theme,
  renderElement,
}: SlideElementRendererProps) {
  // Allow consumers to override any element's rendering
  if (renderElement) {
    const custom = renderElement(element, theme)
    if (custom !== undefined) return <>{custom}</>
  }

  switch (element.type) {
    case 'text':
      return <TextElement element={element} theme={theme} />
    case 'image':
      return <ImageElement element={element} theme={theme} />
    case 'shape':
      return <ShapeElement element={element} theme={theme} />
    case 'table':
      return <TableElement element={element} theme={theme} />
    case 'connector':
      return <ConnectorElement element={element} theme={theme} />
    case 'group':
      return <GroupElement element={element} theme={theme} />
    case 'chart':
      return <ChartElement element={element} theme={theme} />
  }
}

export { TextElement, ImageElement, ShapeElement, TableElement, ConnectorElement, GroupElement, ChartElement }
