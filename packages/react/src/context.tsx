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

/**
 * Returns the `PresentationStore` from the nearest `<Presentation.Root>`.
 * Throws if called outside a `Presentation` tree.
 *
 * @param consumerName - Component name included in the error message for easier debugging.
 */
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

export interface UsePresentationResult {
  /** Parsed presentation data, or `null` before the first successful load. */
  presentation: PresentationData | null;

  /** Current lifecycle status of the store. */
  status: PresentationState["status"];

  /** Error thrown during the last failed parse, or `null` otherwise. */
  error: Error | null;

  /** Parse progress reported by the store (0-100). */
  progress: number;
}

/**
 * Subscribes to top-level presentation state: parse status, progress, and errors.
 *
 * Must be called inside a `<Presentation.Root>` tree.
 */
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

export interface UseSlideResult {
  /** Full parsed data for the active slide, or `null` before load. */
  slide: SlideData | null;

  /**
   * Stable identity of the active slide (`SlideData.id`).
   * Use this (not `index`) as the source of truth for navigation,
   * keys, and any future editing operations.
   */
  slideId: string | null;

  /**
   * Current display position (0-based). Derived from `slideId` via
   * `findIndex`. Safe to use for display but do NOT store it as identity.
   */
  index: number;

  /** Total number of slides in the loaded presentation. `0` before load. */
  total: number;

  /** `true` when the active slide is the first in the deck. */
  isFirst: boolean;

  /** `true` when the active slide is the last in the deck. */
  isLast: boolean;

  /** Navigate to a slide by its stable id. */
  goTo: (slideId: string) => void;

  /** Navigate to a slide by its 0-based index. Clamps to a valid range. */
  goToIndex: (index: number) => void;

  /** Advance to the next slide. No-ops on the last slide. */
  next: () => void;

  /** Go back to the previous slide. No-ops on the first slide. */
  prev: () => void;
}

/**
 * Subscribes to slide navigation state and exposes navigation actions.
 *
 * Must be called inside a `<Presentation.Root>` tree.
 */
export function useSlide(): UseSlideResult {
  const store = usePresentationStore("useSlide");

  // Subscribe to primitives/stable refs only; new object literals on every call cause infinite loops.
  const slideId = React.useSyncExternalStore(
    store.subscribe,
    () => store.getState().activeSlideId,
    () => null,
  );

  // Subscribe to a stable object ref that only changes on new file load, not on slide navigation.
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

export interface UseZoomResult {
  /** Current zoom level (1 = 100%, 0.5 = 50%). */
  zoom: number;

  /** Set an explicit zoom level. */
  setZoom: (zoom: number) => void;

  /**
   * Increase zoom by `step`.
   *
   * @default step 0.1
   */
  zoomIn: (step?: number) => void;

  /**
   * Decrease zoom by `step`.
   *
   * @default step 0.1
   */
  zoomOut: (step?: number) => void;

  /**
   * Compute and apply a zoom that fits the slide inside the given container
   * dimensions, respecting the optional padding.
   *
   * @param containerWidth - Available width in pixels.
   * @param containerHeight - Available height in pixels.
   * @param padding - Padding on all sides in pixels. Defaults to `24`.
   */
  fitTo: (containerWidth: number, containerHeight: number, padding?: number) => void;
}

/**
 * Subscribes to zoom state and exposes zoom actions.
 *
 * Must be called inside a `<Presentation.Root>` tree. For automatic fitting,
 * prefer `<Presentation.Viewport autoFit>` which calls `fitTo` internally.
 */
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
