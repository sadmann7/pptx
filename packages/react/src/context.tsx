import * as React from "react";

import type { PresentationData, SlideData } from "@diceui/pptx-parser";

import type { PresentationState, PresentationStore } from "./store";
import { createPresentationStore } from "./store";

export const PresentationContext = React.createContext<PresentationStore | null>(null);

/**
 * Creates a stable `PresentationStore` instance for use in controlled mode.
 *
 * ```tsx
 * const store = useCreatePresentationStore();
 *
 * // Load manually: e.g. after a fetch/upload
 * await store.load(buffer, { defaultSlideIndex: 2 });
 *
 * // Pass the store to Root; `file` prop is no longer needed
 * <Presentation.Root store={store}>…</Presentation.Root>
 * ```
 */
export function useCreatePresentationStore(): PresentationStore {
  const ref = React.useRef<PresentationStore | null>(null);
  if (ref.current === null) {
    ref.current = createPresentationStore();
  }
  return ref.current;
}

export function usePresentationStore(consumerName: string): PresentationStore {
  const store = React.useContext(PresentationContext);
  if (!store) {
    throw new Error(`\`${consumerName}\` must be used inside \`Presentation\``);
  }
  return store;
}

const SERVER_SNAPSHOT: PresentationState = {
  status: "idle",
  presentation: null,
  activeSlideId: null,
  zoom: 1,
  progress: 0,
  error: null,
};

// ---------------------------------------------------------------------------
// usePresentation
// ---------------------------------------------------------------------------

export interface UsePresentationResult {
  presentation: PresentationData | null;
  status: PresentationState["status"];
  error: Error | null;
  progress: number;
}

export function usePresentation(): UsePresentationResult {
  const store = usePresentationStore("usePresentation");
  const state = React.useSyncExternalStore(store.subscribe, store.getState, () => SERVER_SNAPSHOT);
  return {
    presentation: state.presentation,
    status: state.status,
    error: state.error,
    progress: state.progress,
  };
}

// ---------------------------------------------------------------------------
// useSlide
// ---------------------------------------------------------------------------

export interface UseSlideResult {
  slide: SlideData | null;
  /**
   * Stable identity of the active slide (`SlideData.id`).
   * Use this: not `index`: as the source of truth for navigation,
   * keys, and any future editing operations.
   */
  slideId: string | null;
  /**
   * Current display position (0-based). Derived from `slideId` via
   * `findIndex`: safe to use for display but do NOT store it as identity.
   */
  index: number;
  total: number;
  isFirst: boolean;
  isLast: boolean;
  /** Navigate by stable slide ID. */
  goTo: (slideId: string) => void;
  /** Navigate by current position. Clamps to valid range. */
  goToIndex: (index: number) => void;
  next: () => void;
  prev: () => void;
}

export function useSlide(): UseSlideResult {
  const store = usePresentationStore("useSlide");

  // Subscribe to primitives / stable references only.
  // useSyncExternalStore compares via Object.is: returning a new object
  // literal on every call would cause an infinite re-render loop.

  // activeSlideId: string | null: primitive, safe to compare directly.
  const slideId = React.useSyncExternalStore(
    store.subscribe,
    () => store.getState().activeSlideId,
    () => null,
  );

  // presentation: object reference: only changes when a new file is loaded,
  // not on slide navigation (the store spreads state but keeps the same ref).
  const presentation = React.useSyncExternalStore(
    store.subscribe,
    () => store.getState().presentation,
    () => null,
  );

  // Derive the rest synchronously from the two stable subscribed values.
  const index =
    presentation && slideId ? presentation.slides.findIndex((s) => s.id === slideId) : -1;
  const slide = index >= 0 ? (presentation?.slides[index] ?? null) : null;
  const total = presentation?.slides.length ?? 0;

  return {
    slide,
    slideId,
    index,
    total,
    isFirst: index === 0,
    isLast: total > 0 && index === total - 1,
    goTo: React.useCallback((id: string) => store.goTo(id), [store]),
    goToIndex: React.useCallback((i: number) => store.goToIndex(i), [store]),
    next: React.useCallback(() => store.next(), [store]),
    prev: React.useCallback(() => store.prev(), [store]),
  };
}

// ---------------------------------------------------------------------------
// useZoom
// ---------------------------------------------------------------------------

export interface UseZoomResult {
  zoom: number;
  setZoom: (zoom: number) => void;
  zoomIn: (step?: number) => void;
  zoomOut: (step?: number) => void;
  fitTo: (containerWidth: number, containerHeight: number, padding?: number) => void;
}

export function useZoom(): UseZoomResult {
  const store = usePresentationStore("useZoom");
  const zoom = React.useSyncExternalStore(
    store.subscribe,
    () => store.getState().zoom,
    () => 1,
  );
  return {
    zoom,
    setZoom: React.useCallback((z: number) => store.setZoom(z), [store]),
    zoomIn: React.useCallback((step?: number) => store.zoomIn(step), [store]),
    zoomOut: React.useCallback((step?: number) => store.zoomOut(step), [store]),
    fitTo: React.useCallback((w: number, h: number, p?: number) => store.fitTo(w, h, p), [store]),
  };
}
