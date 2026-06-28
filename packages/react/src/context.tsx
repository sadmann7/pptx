import React from 'react'
import type { Presentation, Slide, ThemeColors } from '@pptx/parser'
import type { PresentationState } from './store'
import { PresentationStore } from './store'

// ─── Context ─────────────────────────────────────────────────────────────────

export const PresentationContext = React.createContext<PresentationStore | null>(null)

function usePresentationStore(): PresentationStore {
  const store = React.useContext(PresentationContext)
  if (!store) {
    throw new Error(
      '[pptx/react] Presentation hooks must be used inside <Presentation.Root>',
    )
  }
  return store
}

// ─── Server snapshot ─────────────────────────────────────────────────────────

const SERVER_SNAPSHOT: PresentationState = {
  status: 'idle',
  presentation: null,
  currentIndex: 0,
  zoom: 1,
  progress: 0,
  error: null,
}

// ─── Public hooks ─────────────────────────────────────────────────────────────

export interface UsePresentationResult {
  presentation: Presentation | null
  status: PresentationState['status']
  error: Error | null
  progress: number
}

export function usePresentation(): UsePresentationResult {
  const store = usePresentationStore()

  /**
   * Return store.getState() directly — it's a stable object reference that
   * is only replaced (new reference) when setState is called. This means
   * Object.is comparisons in useSyncExternalStore work correctly:
   *   - Same state → same reference → no re-render
   *   - New state → new reference → re-render
   *
   * Returning a destructured object literal here would create a new object
   * on every getSnapshot call, causing an infinite update loop.
   */
  const state = React.useSyncExternalStore(
    store.subscribe.bind(store),
    store.getState.bind(store),
    () => SERVER_SNAPSHOT,
  )

  return {
    presentation: state.presentation,
    status: state.status,
    error: state.error,
    progress: state.progress,
  }
}

export interface UseSlideResult {
  slide: Slide | null
  index: number
  total: number
  isFirst: boolean
  isLast: boolean
  goTo: (index: number) => void
  next: () => void
  prev: () => void
}

export function useSlide(): UseSlideResult {
  const store = usePresentationStore()

  // Use primitive selectors to avoid returning new objects from getSnapshot
  const currentIndex = React.useSyncExternalStore(
    store.subscribe.bind(store),
    () => store.getState().currentIndex,
    () => 0,
  )

  const total = React.useSyncExternalStore(
    store.subscribe.bind(store),
    () => store.getState().presentation?.slides.length ?? 0,
    () => 0,
  )

  /**
   * Slide is a reference into the parsed presentation array.
   * The same Slide object is returned as long as currentIndex and
   * the presentation haven't changed — Object.is handles this correctly.
   */
  const slide = React.useSyncExternalStore(
    store.subscribe.bind(store),
    () => {
      const { presentation, currentIndex: idx } = store.getState()
      return presentation?.slides[idx] ?? null
    },
    () => null,
  )

  const goTo = React.useCallback((i: number) => store.goTo(i), [store])
  const next = React.useCallback(() => store.next(), [store])
  const prev = React.useCallback(() => store.prev(), [store])

  return {
    slide,
    index: currentIndex,
    total,
    isFirst: currentIndex === 0,
    isLast: total > 0 && currentIndex === total - 1,
    goTo,
    next,
    prev,
  }
}

export interface UseZoomResult {
  zoom: number
  setZoom: (zoom: number) => void
  zoomIn: (step?: number) => void
  zoomOut: (step?: number) => void
  fitTo: (containerWidth: number, containerHeight: number, padding?: number) => void
}

export function useZoom(): UseZoomResult {
  const store = usePresentationStore()

  // Primitive — Object.is comparison on a number is always correct
  const zoom = React.useSyncExternalStore(
    store.subscribe.bind(store),
    () => store.getState().zoom,
    () => 1,
  )

  return {
    zoom,
    setZoom: React.useCallback((z) => store.setZoom(z), [store]),
    zoomIn: React.useCallback((step) => store.zoomIn(step), [store]),
    zoomOut: React.useCallback((step) => store.zoomOut(step), [store]),
    fitTo: React.useCallback((w, h, p) => store.fitTo(w, h, p), [store]),
  }
}

/** Access the current slide's theme colors for resolving Color values. */
export function useThemeColors(): ThemeColors | null {
  const store = usePresentationStore()
  return React.useSyncExternalStore(
    store.subscribe.bind(store),
    () => store.getState().presentation?.theme.colors ?? null,
    () => null,
  )
}

/** Low-level store access for custom integrations. */
export function usePresentationStoreRef(): PresentationStore {
  return usePresentationStore()
}

/** Full raw state — prefer the specific hooks above. */
export function usePresentationState(): PresentationState {
  const store = usePresentationStore()
  return React.useSyncExternalStore(
    store.subscribe.bind(store),
    store.getState.bind(store),
    () => SERVER_SNAPSHOT,
  )
}
