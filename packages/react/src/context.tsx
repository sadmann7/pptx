import * as React from "react";

import type { PresentationData, SlideData } from "@diceui/pptx-core";

import { DEFAULT_STORE_STATE } from "./constant";
import { useLatestRef, useLazyRef } from "./hook";
import type { AutoFitPadding, Store, StoreEventMap, StoreState, ZoomLevel } from "./store";
import { createStore } from "./store";

export const Context = React.createContext<Store | null>(null);

export interface ProviderProps {
  /** The `PresentationStore` to make available to descendants. */
  store: Store;

  /** The children to render inside the provider. */
  children?: React.ReactNode;
}

/**
 * Provides a `PresentationStore` to descendants without rendering any DOM.
 *
 * Use this when a component needs `usePresentation`/`useSlide`/`useZoom`
 * but must live outside `<Presentation.Root>`'s DOM tree (e.g. a debug bar
 * or toolbar positioned as a sibling so it doesn't join `Root`'s layout).
 *
 * ```tsx
 * const store = useCreatePresentationStore();
 *
 * <Presentation.Provider store={store}>
 *   <DebugBar />
 *   <Presentation.Root>…</Presentation.Root> // inherits the store from Provider
 * </Presentation.Provider>
 * ```
 */
export function Provider({ store, children }: ProviderProps) {
  return <Context.Provider value={store}>{children}</Context.Provider>;
}

export namespace Provider {
  export type Props = ProviderProps;
}

/**
 * Creates a stable `PresentationStore` instance for use in controlled mode.
 *
 * ```tsx
 * const store = useCreatePresentationStore();
 *
 * // Load manually: e.g. after a fetch/upload
 * await store.load(buffer, { defaultSlideIndex: 2 });
 *
 * // Provide the store; `Root` inherits it and the `file` prop is not needed
 * <Presentation.Provider store={store}>
 *   <Presentation.Root>…</Presentation.Root>
 * </Presentation.Provider>
 * ```
 */
export function useCreateStore(): Store {
  return useLazyRef(createStore).current;
}

/**
 * Returns the `PresentationStore` from the nearest `Presentation.Root`
 * or `Presentation.Provider`. Throws if called outside either.
 */
export function useStoreContext(consumerName: string): Store {
  const store = React.useContext(Context);
  if (!store) {
    throw new Error(
      `\`${consumerName}\` must be used within \`Presentation.Root\` or \`Presentation.Provider\``,
    );
  }
  return store;
}

/**
 * Bridges an optional event-handler prop to a store event for the lifetime of
 * the component. Internal: components expose the handler as a prop rather than
 * making consumers manage a subscription.
 *
 * The handler is read through a ref, so passing an inline function does not
 * resubscribe on every render.
 */
export function useStoreEvent<E extends keyof StoreEventMap>(
  store: Store,
  event: E,
  handler: ((payload: StoreEventMap[E]) => void) | undefined,
): void {
  const handlerRef = useLatestRef(handler);

  // Subscribing to an external store: an effect is the right tool here.
  React.useEffect(
    () => store.on(event, (payload) => handlerRef.current?.(payload)),
    [store, event, handlerRef],
  );
}

/**
 * Subscribes to a single derived value from the presentation store.
 *
 * Re-renders the caller only when `selector(state)` returns a different value
 * (compared with `Object.is`). Use this instead of calling `store.getState()`
 * directly so unrelated state changes don't cause unnecessary re-renders.
 */
export function useStoreSelector<T>(
  store: Store,
  selector: (state: StoreState) => T,
  serverSnapshot: T,
): T {
  return React.useSyncExternalStore(
    store.subscribe,
    () => selector(store.getState()),
    () => serverSnapshot,
  );
}

/**
 * Subscribes to a slide's edit revision (`store.getSlideRevision`).
 */
export function useSlideRevision(store: Store, slideId: string | null | undefined): number {
  return React.useSyncExternalStore(
    store.subscribe,
    () => (slideId != null ? store.getSlideRevision(slideId) : 0),
    () => 0,
  );
}

/**
 * Subscribes to a slide's content revision (`store.getSlideContentRevision`):
 * bumped by content edits but not by transform-only moves.
 */
export function useSlideContentRevision(store: Store, slideId: string | null | undefined): number {
  return React.useSyncExternalStore(
    store.subscribe,
    () => (slideId != null ? store.getSlideContentRevision(slideId) : 0),
    () => 0,
  );
}

/**
 * Subscribes to a slide's 0-based index (`store.getSlideIndex`).
 */
export function useSlideIndex(store: Store, slideId: string): number {
  return React.useSyncExternalStore(
    store.subscribe,
    () => store.getSlideIndex(slideId),
    () => -1,
  );
}

export interface UsePresentationResult {
  /** Parsed presentation data, or `null` before the first successful load. */
  presentation: PresentationData | null;

  /** Current lifecycle status of the store. */
  status: StoreState["status"];

  /** Error thrown during the last failed parse, or `null` otherwise. */
  error: Error | null;

  /** Parse progress reported by the store (0-100). */
  progress: number;

  /**
   * Edit revision counter; bumps on every `store.edit()`, `undo()`, or
   * `redo()`. The `presentation` object is mutated in place by edits, so
   * derive from this value (not object identity) to react to content changes.
   */
  revision: number;
}

/**
 * Subscribes to top-level presentation state: parse status, progress, and errors.
 *
 * Must be called inside a `<Presentation.Root>` tree.
 *
 * Each field is subscribed independently so that unrelated store updates
 * (zoom changes, slide navigation) do not cause consumers to re-render.
 */
export function usePresentation(): UsePresentationResult {
  const store = useStoreContext("usePresentation");
  const presentation = useStoreSelector(
    store,
    (s) => s.presentation,
    DEFAULT_STORE_STATE.presentation,
  );
  const status = useStoreSelector(store, (s) => s.status, DEFAULT_STORE_STATE.status);
  const error = useStoreSelector(store, (s) => s.error, DEFAULT_STORE_STATE.error);
  const progress = useStoreSelector(store, (s) => s.progress, DEFAULT_STORE_STATE.progress);
  const revision = useStoreSelector(store, (s) => s.revision, DEFAULT_STORE_STATE.revision);

  return { presentation, status, error, progress, revision };
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
  const store = useStoreContext("useSlide");

  // Subscribe to primitives/stable refs only; new object literals on every call cause infinite loops.
  const slideId = useStoreSelector(store, (s) => s.activeSlideId, null);
  // Stable object ref that only changes on new file load, not on slide navigation.
  const presentation = useStoreSelector(store, (s) => s.presentation, null);

  const index = slideId ? store.getSlideIndex(slideId) : -1;
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
  /** Zoom the slide is rendered at (1 = 100%, 0.5 = 50%). */
  zoom: number;

  /**
   * The zoom that was asked for: a level, or `"fit"` while the viewport keeps
   * the slide fitted. A zoom control reads this as its value.
   */
  zoomLevel: ZoomLevel;

  /**
   * Set an explicit zoom level, or `"fit"` to hand the zoom back to the
   * viewport. A level stays put across container resizes.
   *
   * ```tsx
   * setZoom(1.5);
   * setZoom("fit");
   * ```
   */
  setZoom: (zoom: ZoomLevel) => void;

  /**
   * Increase zoom by `step`.
   *
   * @default step 0.25
   */
  zoomIn: (step?: number) => void;

  /**
   * Decrease zoom by `step`.
   *
   * @default step 0.25
   */
  zoomOut: (step?: number) => void;

  /**
   * Compute and apply a zoom that fits the slide inside the given container
   * dimensions, respecting the optional padding.
   *
   * @param containerWidth - Available width in pixels.
   * @param containerHeight - Available height in pixels.
   * @param padding - Uniform padding or per-side values in pixels.
   */
  fitTo: (containerWidth: number, containerHeight: number, padding?: AutoFitPadding) => void;
}

/**
 * Subscribes to zoom state and exposes zoom actions.
 *
 * Must be called inside a `<Presentation.Root>` tree. For automatic fitting,
 * prefer `<Presentation.Viewport autoFit>` which calls `fitTo` internally.
 */
export function useZoom(): UseZoomResult {
  const store = useStoreContext("useZoom");
  const zoom = useStoreSelector(store, (s) => s.zoom, 1);
  const zoomLevel = useStoreSelector<ZoomLevel>(store, (s) => s.zoomLevel, "fit");

  return {
    zoom,
    zoomLevel,
    setZoom: React.useCallback((level: ZoomLevel) => store.setZoom(level), [store]),
    zoomIn: React.useCallback((step?: number) => store.zoomIn(step), [store]),
    zoomOut: React.useCallback((step?: number) => store.zoomOut(step), [store]),
    fitTo: React.useCallback(
      (w: number, h: number, padding?: AutoFitPadding) => store.fitTo(w, h, padding),
      [store],
    ),
  };
}
