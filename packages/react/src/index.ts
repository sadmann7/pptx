// ─── Compound component namespace ────────────────────────────────────────────
import { Root } from './components/Root'
import { Viewport } from './components/Viewport'
import { Slide } from './components/Slide'
import { Thumbnails } from './components/Thumbnails'
import { Notes } from './components/Notes'
import {
  PreviousSlide,
  NextSlide,
  SlideCounter,
  ZoomIn,
  ZoomOut,
  ZoomReset,
} from './components/Controls'

export const Presentation = {
  Root,
  Viewport,
  Slide,
  Thumbnails,
  Notes,
  PreviousSlide,
  NextSlide,
  SlideCounter,
  ZoomIn,
  ZoomOut,
  ZoomReset,
} as const

// ─── Hooks ────────────────────────────────────────────────────────────────────
export {
  usePresentation,
  useSlide,
  useZoom,
  useThemeColors,
  usePresentationState,
  usePresentationStoreRef,
} from './context'

// ─── Store (for advanced integrations) ───────────────────────────────────────
export { PresentationStore } from './store'
export type { PresentationState, PresentationStatus } from './store'

// ─── Individual components (for tree-shaking / custom composition) ────────────
export { Root } from './components/Root'
export { Viewport } from './components/Viewport'
export { Slide } from './components/Slide'
export { Thumbnails } from './components/Thumbnails'
export { Notes } from './components/Notes'
export {
  PreviousSlide,
  NextSlide,
  SlideCounter,
  ZoomIn,
  ZoomOut,
  ZoomReset,
} from './components/Controls'

// ─── Element renderers (for custom element overrides) ────────────────────────
export {
  SlideElementRenderer,
  TextElement,
  ImageElement,
  ShapeElement,
  TableElement,
  ConnectorElement,
  GroupElement,
  ChartElement,
} from './elements/index'
export type { ElementRendererFn } from './elements/index'

// ─── Prop types ───────────────────────────────────────────────────────────────
export type { RootProps } from './components/Root'
export type { SlideProps } from './components/Slide'
export type { ViewportProps } from './components/Viewport'
export type { ThumbnailsProps } from './components/Thumbnails'
export type { NotesProps } from './components/Notes'
export type {
  PreviousSlideProps,
  NextSlideProps,
  SlideCounterProps,
  ZoomInProps,
  ZoomOutProps,
  ZoomResetProps,
} from './components/Controls'
