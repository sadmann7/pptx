import * as React from "react";

import type { PresentationData, SlideData } from "@diceui/pptx-core";

import { DEFAULT_STORE_STATE } from "./constant";
import { useLatestRef, useLazyRef } from "./hooks";
import type { AutoFitPadding, Store, StoreEventMap, StoreState } from "./store";
import { createStore } from "./store";

export const Context = React.createContext<Store | null>(null);

// The deck's focus boundary, thumbnails included. A ref, since only event handlers read it.
export const RootContext =
  React.createContext<React.RefObject<HTMLElement | null> | null>(null);

export interface ProviderProps {
  /** The `PresentationStore` to make available to descendants. */
  store: Store;

  children?: React.ReactNode;
}

export function Provider({ store, children }: ProviderProps) {
  return <Context.Provider value={store}>{children}</Context.Provider>;
}

export namespace Provider {
  export type Props = ProviderProps;
}

/**
 * Creates a stable `PresentationStore` for controlled mode, reset when the calling component unmounts.
 *
 * ```tsx
 * const store = useCreatePresentationStore();
 * await store.load(buffer, { defaultSlideIndex: 2 });
 *
 * <Presentation.Provider store={store}>
 *   <Presentation.Root>…</Presentation.Root>
 * </Presentation.Provider>
 * ```
 */
export function useCreateStore(): Store {
  const store = useLazyRef(createStore).current;
  React.useEffect(() => () => store.reset(), [store]);
  return store;
}

export function useStoreContext(consumerName: string): Store {
  const store = React.useContext(Context);
  if (!store) {
    throw new Error(
      `\`${consumerName}\` must be used within \`Presentation.Root\` or \`Presentation.Provider\`.`,
    );
  }
  return store;
}

/**
 * Returns the stable `PresentationStore` from `Presentation.Provider` or `Presentation.Root`
 * without subscribing to it; read state with `usePresentation`, `useSlide` or `useZoom`.
 *
 * ```tsx
 * function SaveButton() {
 *   const store = usePresentationStore();
 *   return <button onClick={() => store.save()}>Save</button>;
 * }
 *
 * <Presentation.Root file={file} readOnly={false}>
 *   <SaveButton />
 * </Presentation.Root>
 * ```
 */
export function useStore(): Store {
  return useStoreContext("usePresentationStore");
}

export function useStoreEvent<E extends keyof StoreEventMap>(
  store: Store,
  event: E,
  callback: ((payload: StoreEventMap[E]) => void) | undefined,
): void {
  const callbackRef = useLatestRef(callback);

  React.useEffect(
    () => store.on(event, (payload) => callbackRef.current?.(payload)),
    [store, event, callbackRef],
  );
}

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

/** Subscribes to `store.getSlideRevision`. */
export function useSlideRevision(
  store: Store,
  slideId: string | null | undefined,
): number {
  return React.useSyncExternalStore(
    store.subscribe,
    () => (slideId != null ? store.getSlideRevision(slideId) : 0),
    () => 0,
  );
}

export function useSlideContentRevision(
  store: Store,
  slideId: string | null | undefined,
): number {
  return React.useSyncExternalStore(
    store.subscribe,
    () => (slideId != null ? store.getSlideContentRevision(slideId) : 0),
    () => 0,
  );
}

/** Subscribes to `store.getSlideIndex`. */
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

  /** Parse progress from `0` to `100`. */
  progress: number;

  /**
   * Counter bumped on every edit, undo, or redo. Edits mutate `presentation` in place,
   * so derive from `revision` to react to content changes.
   */
  revision: number;
}

/** Subscribes to the loaded presentation, its parse status, progress, error and revision. */
export function usePresentation(): UsePresentationResult {
  const store = useStoreContext("usePresentation");
  const presentation = useStoreSelector(
    store,
    (s) => s.presentation,
    DEFAULT_STORE_STATE.presentation,
  );
  const status = useStoreSelector(
    store,
    (s) => s.status,
    DEFAULT_STORE_STATE.status,
  );
  const error = useStoreSelector(
    store,
    (s) => s.error,
    DEFAULT_STORE_STATE.error,
  );
  const progress = useStoreSelector(
    store,
    (s) => s.progress,
    DEFAULT_STORE_STATE.progress,
  );
  const revision = useStoreSelector(
    store,
    (s) => s.revision,
    DEFAULT_STORE_STATE.revision,
  );

  return { presentation, status, error, progress, revision };
}

export interface UseSlideResult {
  /** Full parsed data for the active slide, or `null` before load. */
  slide: SlideData | null;

  /** Stable id of the active slide (`SlideData.id`), to use over `index` for keys and navigation. */
  slideId: string | null;

  /** 0-based position of the active slide, for display rather than identity. */
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

/** Subscribes to the active slide and exposes navigation actions. */
export function useSlide(): UseSlideResult {
  const store = useStoreContext("useSlide");

  // Selectors must return stable values; a fresh object each call loops forever.
  const slideId = useStoreSelector(store, (s) => s.activeSlideId, null);
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
  /** Current zoom level, where `1` equals 100%. */
  zoom: number;

  /** Whether zoom tracks the viewport size, so a zoom control can show "Fit". */
  isAutoFit: boolean;

  /** Set an explicit zoom level. Turns auto-fit off. */
  setZoom: (zoom: number) => void;

  /**
   * Increases zoom by `step`. Turns auto-fit off.
   *
   * @default step 0.25
   */
  zoomIn: (step?: number) => void;

  /**
   * Decreases zoom by `step`. Turns auto-fit off.
   *
   * @default step 0.25
   */
  zoomOut: (step?: number) => void;

  /** Turns auto-fit on or off; while on, the slide refits to the viewport on every resize. */
  setAutoFit: (isAutoFit: boolean) => void;

  /** Sets zoom once to fit the slide in a container of the given pixel size, minus padding. */
  fitTo: (
    containerWidth: number,
    containerHeight: number,
    padding?: AutoFitPadding,
  ) => void;
}

/** Subscribes to zoom state and exposes zoom actions; `Presentation.Viewport autoFit` does the fitting. */
export function useZoom(): UseZoomResult {
  const store = useStoreContext("useZoom");
  const zoom = useStoreSelector(store, (s) => s.zoom, 1);
  const isAutoFit = useStoreSelector(
    store,
    (s) => s.isAutoFit,
    DEFAULT_STORE_STATE.isAutoFit,
  );

  return {
    zoom,
    isAutoFit,
    setZoom: React.useCallback((z: number) => store.setZoom(z), [store]),
    zoomIn: React.useCallback((step?: number) => store.zoomIn(step), [store]),
    zoomOut: React.useCallback((step?: number) => store.zoomOut(step), [store]),
    setAutoFit: React.useCallback(
      (next: boolean) => store.setAutoFit(next),
      [store],
    ),
    fitTo: React.useCallback(
      (w: number, h: number, padding?: AutoFitPadding) =>
        store.fitTo(w, h, padding),
      [store],
    ),
  };
}

export interface UseHistoryResult {
  /** Whether there is an edit to undo. */
  canUndo: boolean;

  /** Whether there is an edit to redo. */
  canRedo: boolean;

  /** Whether the deck has unsaved changes, cleared again by undoing back to the last save. */
  isDirty: boolean;

  /** Revert the most recent edit. Resolves `false` when the undo stack is empty. */
  undo: () => Promise<boolean>;

  /** Re-apply the most recently undone edit. Resolves `false` when the redo stack is empty. */
  redo: () => Promise<boolean>;
}

/**
 * Subscribes to undo, redo, and whether the deck has unsaved edits.
 *
 * ```tsx
 * const { canUndo, undo } = useHistory();
 * <button disabled={!canUndo} onClick={() => void undo()}>Undo</button>
 * ```
 */
export function useHistory(): UseHistoryResult {
  const store = useStoreContext("useHistory");

  const subscribe = React.useCallback(
    (onStoreChange: () => void) => store.on("historyChange", onStoreChange),
    [store],
  );

  const canUndo = React.useSyncExternalStore(
    subscribe,
    store.canUndo,
    () => false,
  );
  const canRedo = React.useSyncExternalStore(
    subscribe,
    store.canRedo,
    () => false,
  );
  const isDirty = React.useSyncExternalStore(
    subscribe,
    store.isDirty,
    () => false,
  );

  return {
    canUndo,
    canRedo,
    isDirty,
    undo: React.useCallback(() => store.undo(), [store]),
    redo: React.useCallback(() => store.redo(), [store]),
  };
}
