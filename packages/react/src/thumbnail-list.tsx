/**
 * Roving focus adapted from Radix UI's `@radix-ui/react-roving-focus`.
 * @see https://github.com/radix-ui/primitives/tree/main/packages/react/roving-focus
 */

import * as React from "react";

import type { SlideData, SlideHandle } from "@diceui/pptx-core";
import { applySlideScale, renderSlide } from "@diceui/pptx-core";

import { VISUALLY_HIDDEN_STYLE } from "./constant";
import {
  usePresentation,
  useSlideIndex,
  useSlideRevision,
  useStoreContext,
  useStoreSelector,
} from "./context";
import { useLatestRef, useLazyRef } from "./hooks";
import type { PrimitiveProps } from "./render";
import { renderElement } from "./render";

const THUMBNAIL_LIST_NAME = "Presentation.ThumbnailList";
const THUMBNAIL_ITEM_NAME = "Presentation.ThumbnailItem";
const THUMBNAIL_ITEM_PREVIEW_NAME = "Presentation.ThumbnailItemPreview";
const THUMBNAIL_ITEM_NUMBER_NAME = "Presentation.ThumbnailItemNumber";

// Renders thumbnails before they scroll into view so scrolling never reveals a placeholder.
const INTERSECTION_OBSERVER_ROOT_MARGIN = "200px 0px";

export interface CachedThumbnail {
  slideHandle: SlideHandle;
  revision: number;
}

type FocusIntent = "first" | "last" | "prev" | "next";

const MAP_KEY_TO_INTENT: Record<string, FocusIntent> = {
  ArrowUp: "prev",
  ArrowDown: "next",
  Home: "first",
  End: "last",
};

function focusFirst(candidates: HTMLElement[], preventScroll = false) {
  const prev = document.activeElement;
  for (const candidate of candidates) {
    if (candidate === prev) return;
    candidate.focus({ preventScroll });
    if (document.activeElement !== prev) return;
  }
}

export function getItemsInPageOrder(
  items: Map<string, HTMLButtonElement>,
): HTMLButtonElement[] {
  return Array.from(items.values()).sort((a, b) =>
    a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1,
  );
}

function wrapArray<T>(array: T[], startIndex: number): T[] {
  return array.map((_, i) => array[(startIndex + i) % array.length] as T);
}

// Past this, the least recently shown off-screen miniatures are released.
export const THUMBNAIL_CACHE_LIMIT = 50;

// A Map iterates oldest first, so re-inserting marks the entry as most recent.
export function markThumbnailRecent(
  slideHandleCache: Map<string, CachedThumbnail>,
  slideId: string,
) {
  const cached = slideHandleCache.get(slideId);
  if (!cached) return;
  slideHandleCache.delete(slideId);
  slideHandleCache.set(slideId, cached);
}

export function releaseOldThumbnails(
  slideHandleCache: Map<string, CachedThumbnail>,
  keepId: string,
) {
  for (const [slideId, cached] of slideHandleCache) {
    if (slideHandleCache.size <= THUMBNAIL_CACHE_LIMIT) return;
    if (slideId === keepId || cached.slideHandle.element.parentElement)
      continue;
    cached.slideHandle.dispose();
    slideHandleCache.delete(slideId);
  }
}

// False when a live entry for this revision holds the slot, as when a drag overlay and its
// item mount the same slide; the caller then disposes its private copy.
function getIsThumbnailCached(
  slideHandleCache: Map<string, CachedThumbnail>,
  slideId: string,
  slideHandle: SlideHandle,
  revision: number,
): boolean {
  const cached = slideHandleCache.get(slideId);
  if (cached) {
    if (cached.slideHandle !== slideHandle && cached.revision === revision)
      return false;
    cached.slideHandle.dispose();
    slideHandleCache.delete(slideId);
  }
  slideHandleCache.set(slideId, { slideHandle, revision });
  releaseOldThumbnails(slideHandleCache, slideId);
  return true;
}

interface ThumbnailRovingContextValue {
  // An external store, not context state, so a focus change re-renders only the old and new tab stop.
  subscribeTabStop: (callback: () => void) => () => void;
  getEffectiveTabStopId: () => string | null;
  onItemFocus: (slideId: string) => void;
  onItemRegister: (slideId: string, el: HTMLButtonElement) => void;
  onItemUnregister: (slideId: string) => void;
  itemsRef: React.RefObject<Map<string, HTMLButtonElement>>;
  loop: boolean;
  mediaUrlCache: Map<string, string>;
  // Entries go stale once the slide's edit revision moves past the cached one.
  slideHandleCache: Map<string, CachedThumbnail>;
  observeResize: (element: Element, cb: (width: number) => void) => () => void;
  scheduleRender: (fn: () => void) => () => void;
}

const ThumbnailRovingContext =
  React.createContext<ThumbnailRovingContextValue | null>(null);

function useThumbnailRovingContext(consumerName: string) {
  const context = React.useContext(ThumbnailRovingContext);
  if (!context) {
    throw new Error(
      `\`${consumerName}\` must be used within \`${THUMBNAIL_LIST_NAME}\`.`,
    );
  }
  return context;
}

interface ThumbnailItemContextValue {
  slideId: string;
  displayIndex: number;
  isActive: boolean;
}

const ThumbnailItemContext =
  React.createContext<ThumbnailItemContextValue | null>(null);

function useThumbnailItemContext(consumerName: string) {
  const context = React.useContext(ThumbnailItemContext);
  if (!context) {
    throw new Error(
      `\`${consumerName}\` must be used within \`${THUMBNAIL_ITEM_NAME}\`.`,
    );
  }
  return context;
}

export interface ThumbnailListState {
  /** Total number of slides in the loaded presentation. */
  total: number;

  /** Stable id of the currently active slide, or `null` before load. */
  activeSlideId: string | null;

  /** 0-based position of the active slide. */
  activeIndex: number;
}

export interface ThumbnailListRenderState {
  /** All slides in the loaded presentation, in order. */
  slides: SlideData[];

  /** Stable id of the currently active slide, or `null` before load. */
  activeSlideId: string | null;

  /** 0-based position of the active slide. */
  activeIndex: number;

  /** Navigate to a slide by its stable id. */
  goTo: (slideId: string) => void;

  /** Navigate to a slide by its 0-based index. */
  goToIndex: (index: number) => void;
}

export interface ThumbnailListProps extends Omit<
  PrimitiveProps<"div", ThumbnailListState>,
  "children"
> {
  /** Items to render, one `ThumbnailItem` per slide by default; a function receives the slide state. */
  children?:
    | React.ReactNode
    | ((state: ThumbnailListRenderState) => React.ReactNode);

  /**
   * Whether keyboard navigation wraps between the last and first items.
   *
   * @default false
   */
  loop?: boolean;

  /**
   * What takes focus once a presentation finishes loading, evaluated once per load.
   *
   * - `false` → nothing moves.
   * - `true` → the active thumbnail, or the first one.
   * - A ref → that element, or the default target if it has not attached yet.
   * - A function → called when the load settles; return an element, `true` or
   *   `null` for the default target, or `false`/nothing to leave focus alone.
   *
   * @default false
   */
  initialFocus?:
    | boolean
    | React.RefObject<HTMLElement | null>
    | (() => boolean | void | HTMLElement | null);
}

export const ThumbnailList = React.forwardRef<
  HTMLDivElement,
  ThumbnailListProps
>(function ThumbnailList(
  {
    render,
    children,
    loop = false,
    initialFocus = false,
    ...thumbnailListProps
  },
  forwardedRef,
) {
  const { presentation, status } = usePresentation();
  const store = useStoreContext(THUMBNAIL_LIST_NAME);

  const itemsRef = useLazyRef(() => new Map<string, HTMLButtonElement>());
  const isClickFocusRef = React.useRef(false);
  const autoFocusedPresentationRef = React.useRef<object | null>(null);

  const currentTabStopIdRef = React.useRef<string | null>(null);
  const tabStopListenersRef = useLazyRef(() => new Set<() => void>());

  const getEffectiveTabStopId = React.useCallback(
    () => currentTabStopIdRef.current ?? store.getState().activeSlideId,
    [store],
  );

  const subscribeTabStop = React.useCallback(
    (callback: () => void) => {
      tabStopListenersRef.current.add(callback);
      const unsubscribeStore = store.subscribe(callback);
      return () => {
        tabStopListenersRef.current.delete(callback);
        unsubscribeStore();
      };
    },
    [store, tabStopListenersRef],
  );

  const setCurrentTabStopId = React.useCallback(
    (slideId: string | null) => {
      if (currentTabStopIdRef.current === slideId) return;
      currentTabStopIdRef.current = slideId;
      for (const callback of tabStopListenersRef.current) callback();
    },
    [tabStopListenersRef],
  );

  // oxlint-disable-next-line react-hooks/exhaustive-deps -- presentation is the cache key
  const mediaUrlCache = React.useMemo(
    () => new Map<string, string>(),
    [presentation],
  );
  const slideHandleCache = React.useMemo(
    () => new Map<string, CachedThumbnail>(),
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- presentation is the cache key
    [presentation],
  );

  React.useEffect(() => {
    const slideHandles = slideHandleCache;
    const media = mediaUrlCache;
    return () => {
      for (const entry of slideHandles.values()) entry.slideHandle.dispose();
      slideHandles.clear();
      for (const url of media.values()) URL.revokeObjectURL(url);
      media.clear();
    };
  }, [slideHandleCache, mediaUrlCache]);

  const resizeCallbacksRef = useLazyRef(
    () => new Map<Element, (width: number) => void>(),
  );
  const sharedResizeObserverRef = useLazyRef(() =>
    typeof ResizeObserver !== "undefined"
      ? new ResizeObserver((entries) => {
          for (const entry of entries) {
            resizeCallbacksRef.current.get(entry.target)?.(
              entry.contentRect.width,
            );
          }
        })
      : null,
  );

  const observeResize = React.useCallback(
    (element: Element, cb: (width: number) => void) => {
      resizeCallbacksRef.current.set(element, cb);
      sharedResizeObserverRef.current?.observe(element);
      return () => {
        resizeCallbacksRef.current.delete(element);
        sharedResizeObserverRef.current?.unobserve(element);
      };
    },
    [resizeCallbacksRef, sharedResizeObserverRef],
  );

  const renderQueueRef = useLazyRef<Array<() => void>>(() => []);
  const drainRafRef = React.useRef<number | null>(null);

  const drainRenderQueue = React.useCallback(
    function drain() {
      drainRafRef.current = null;
      const queue = renderQueueRef.current;
      if (queue.length === 0) return;
      const budgetMs = 8;
      const frameStart = performance.now();
      do {
        queue.shift()?.();
      } while (queue.length > 0 && performance.now() - frameStart < budgetMs);
      if (queue.length > 0) drainRafRef.current = requestAnimationFrame(drain);
    },
    [renderQueueRef],
  );

  const scheduleRender = React.useCallback(
    (fn: () => void): (() => void) => {
      renderQueueRef.current.push(fn);
      if (drainRafRef.current === null)
        drainRafRef.current = requestAnimationFrame(drainRenderQueue);
      return () => {
        const idx = renderQueueRef.current.indexOf(fn);
        if (idx !== -1) renderQueueRef.current.splice(idx, 1);
      };
    },
    [drainRenderQueue, renderQueueRef],
  );

  React.useEffect(() => {
    return () => {
      sharedResizeObserverRef.current?.disconnect();
      sharedResizeObserverRef.current = null;
      if (drainRafRef.current !== null) {
        cancelAnimationFrame(drainRafRef.current);
        drainRafRef.current = null;
      }
      renderQueueRef.current = [];
    };
  }, [sharedResizeObserverRef, renderQueueRef]);

  const activeSlideId = useStoreSelector(store, (s) => s.activeSlideId, null);

  const initialFocusRef = useLatestRef(initialFocus);

  React.useEffect(() => {
    if (!presentation || autoFocusedPresentationRef.current === presentation)
      return;
    // Marked first so a load that wants no focus is not revisited when the active slide changes.
    autoFocusedPresentationRef.current = presentation;

    const target = initialFocusRef.current;
    const resolved = typeof target === "function" ? target() : target;
    if (resolved === false || resolved === undefined) return;

    let element: HTMLElement | null | undefined;
    if (resolved !== true && resolved !== null) {
      element = "current" in resolved ? resolved.current : resolved;
    }

    // A named element that is not there yet (a ref before it attaches) falls
    // back to the default target rather than dropping the focus request.
    if (!element) {
      const items = itemsRef.current;
      const activeItem = activeSlideId ? items.get(activeSlideId) : undefined;
      element = activeItem ?? getItemsInPageOrder(items)[0];
    }
    element?.focus({ preventScroll: true });
  }, [presentation, activeSlideId, itemsRef, initialFocusRef]);

  React.useEffect(() => {
    return store.on("slideChange", ({ slideId, reason }) => {
      if (reason !== "edit" || !slideId) return;
      setCurrentTabStopId(slideId);
      const items = itemsRef.current;
      const nextItem = items.get(slideId);
      if (!nextItem) return;
      const activeElement = document.activeElement;
      if (activeElement === nextItem) return;
      // Only take focus if the list already had it: an undo triggered from a
      // toolbar button should move the tab stop without stealing the caret.
      for (const item of items.values()) {
        if (item === activeElement) {
          nextItem.focus({ preventScroll: true });
          return;
        }
      }
    });
  }, [store, setCurrentTabStopId, itemsRef]);

  const hasTabStop = React.useSyncExternalStore(
    subscribeTabStop,
    () => getEffectiveTabStopId() != null,
    () => false,
  );

  const rovingContextValue = React.useMemo<ThumbnailRovingContextValue>(
    () => ({
      subscribeTabStop,
      getEffectiveTabStopId,
      onItemFocus: setCurrentTabStopId,
      onItemRegister: (slideId, element) =>
        itemsRef.current.set(slideId, element),
      onItemUnregister: (slideId) => itemsRef.current.delete(slideId),
      itemsRef,
      loop,
      mediaUrlCache,
      slideHandleCache,
      observeResize,
      scheduleRender,
    }),
    [
      subscribeTabStop,
      getEffectiveTabStopId,
      setCurrentTabStopId,
      itemsRef,
      loop,
      mediaUrlCache,
      slideHandleCache,
      observeResize,
      scheduleRender,
    ],
  );

  if (status !== "ready" || !presentation) return null;

  const total = presentation.slides.length;
  const activeIndex = activeSlideId ? store.getSlideIndex(activeSlideId) : -1;

  const state: ThumbnailListState = { total, activeSlideId, activeIndex };

  let resolvedChildren: React.ReactNode;
  if (typeof children === "function") {
    resolvedChildren = children({
      slides: presentation.slides,
      activeSlideId,
      activeIndex,
      goTo: (id) => store.goTo(id),
      goToIndex: (i) => store.goToIndex(i),
    });
  } else if (children != null) {
    resolvedChildren = children;
  } else {
    resolvedChildren = presentation.slides.map((slide) => (
      <ThumbnailItem key={slide.id} slideId={slide.id}>
        <ThumbnailItemPreview />
        <ThumbnailItemNumber style={VISUALLY_HIDDEN_STYLE} />
      </ThumbnailItem>
    ));
  }

  return (
    <ThumbnailRovingContext.Provider value={rovingContextValue}>
      {renderElement(
        "div",
        { render },
        {
          state,
          ref: forwardedRef,
          props: [
            {
              role: "listbox",
              "aria-label": "Slides",
              "aria-orientation": "vertical",
              tabIndex: hasTabStop ? -1 : 0,
              onMouseDown: (event: React.MouseEvent<HTMLDivElement>) => {
                if (event.target === event.currentTarget)
                  isClickFocusRef.current = true;
              },
              onFocus: (event: React.FocusEvent<HTMLDivElement>) => {
                if (event.target !== event.currentTarget) return;
                if (isClickFocusRef.current) {
                  isClickFocusRef.current = false;
                  return;
                }
                focusFirst(getItemsInPageOrder(itemsRef.current), true);
              },
              children: resolvedChildren,
            },
            thumbnailListProps,
          ],
        },
      )}
    </ThumbnailRovingContext.Provider>
  );
});

export namespace ThumbnailList {
  export type State = ThumbnailListState;
  export type RenderState = ThumbnailListRenderState;
  export type Props = ThumbnailListProps;
}

export interface ThumbnailItemState {
  /** Stable id of the slide this item represents (`SlideData.id`). */
  slideId: string;

  /** `true` when this item's slide is the currently active slide. */
  isActive: boolean;

  /** 0-based position of this slide in the presentation. */
  displayIndex: number;
}

/** Triggered when a thumbnail is about to become the active slide. */
export interface ThumbnailSelectEvent {
  /** Stable id of the slide the list is about to navigate to. */
  slideId: string;

  /** Call to keep the navigation from happening. */
  preventDefault: () => void;
}

export interface ThumbnailItemProps extends Omit<
  PrimitiveProps<"button", ThumbnailItemState>,
  "onSelect"
> {
  slideId: string;

  /**
   * Runs before this item becomes the active slide, on click or keyboard focus; call
   * `preventDefault()` to stop it, since `onClick` fires too late. Keep it stable to keep the item memoized.
   *
   * ```tsx
   * <Presentation.ThumbnailItem
   *   slideId={slideId}
   *   onSelect={(event) => {
   *     if (store.isDirty() && !confirmDiscard()) event.preventDefault();
   *   }}
   * />
   * ```
   */
  onSelect?: (event: ThumbnailSelectEvent) => void;

  /**
   * Renders only the item's visuals, with no listbox role, focus or navigation, for a second
   * copy of an item such as a drag overlay's thumbnail.
   *
   * @default false
   */
  decorative?: boolean;
}

export const ThumbnailItem = React.memo(
  React.forwardRef<HTMLButtonElement, ThumbnailItemProps>(
    function ThumbnailItem(
      {
        slideId,
        children,
        render,
        onSelect,
        decorative = false,
        ...thumbnailItemProps
      },
      forwardedRef,
    ) {
      const store = useStoreContext(THUMBNAIL_ITEM_NAME);
      const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_NAME);

      const isActive = useStoreSelector(
        store,
        (s) => s.activeSlideId === slideId,
        false,
      );

      const displayIndex = useSlideIndex(store, slideId);

      const setSize = useStoreSelector(
        store,
        (s) => s.presentation?.slides.length ?? 0,
        0,
      );

      const { onItemRegister, onItemUnregister } = rovingContext;
      const registerRef = React.useCallback(
        (element: HTMLButtonElement | null) => {
          if (decorative) return;
          if (element) onItemRegister(slideId, element);
          else onItemUnregister(slideId);
        },
        [slideId, decorative, onItemRegister, onItemUnregister],
      );

      const isCurrentTabStop = React.useSyncExternalStore(
        rovingContext.subscribeTabStop,
        () => rovingContext.getEffectiveTabStopId() === slideId,
        () => false,
      );

      const state: ThumbnailItemState = { slideId, isActive, displayIndex };

      // Click and focus share this path so a veto applies whichever fires first.
      function selectSlide(): void {
        if (onSelect) {
          let isPrevented = false;
          onSelect({
            slideId,
            preventDefault: () => {
              isPrevented = true;
            },
          });
          if (isPrevented) return;
        }
        store.goTo(slideId);
      }

      const optionProps: React.ComponentProps<"button"> = decorative
        ? { "aria-hidden": true, inert: true, tabIndex: -1 }
        : {
            role: "option",
            "aria-label": `Slide ${displayIndex + 1}`,
            "aria-selected": isActive,
            "aria-posinset": displayIndex + 1,
            "aria-setsize": setSize,
            tabIndex: isCurrentTabStop ? 0 : -1,
            onClick: () => selectSlide(),
            onFocus: () => {
              rovingContext.onItemFocus(slideId);
              selectSlide();
            },
            onMouseDown: () => {
              rovingContext.onItemFocus(slideId);
            },
            onKeyDown: (event: React.KeyboardEvent<HTMLButtonElement>) => {
              if (event.target !== event.currentTarget) return;
              const focusIntent = MAP_KEY_TO_INTENT[event.key];
              if (
                !focusIntent ||
                event.metaKey ||
                event.ctrlKey ||
                event.altKey ||
                event.shiftKey
              )
                return;
              event.preventDefault();

              let candidates = getItemsInPageOrder(
                rovingContext.itemsRef.current,
              );
              if (focusIntent === "last") {
                candidates = candidates.reverse();
              } else if (focusIntent === "prev" || focusIntent === "next") {
                if (focusIntent === "prev") candidates = candidates.reverse();
                const idx = candidates.indexOf(event.currentTarget);
                candidates = rovingContext.loop
                  ? wrapArray(candidates, idx + 1)
                  : candidates.slice(idx + 1);
              }
              // Deferred so the browser finishes processing this keydown
              // before focus moves; synchronous focus can be swallowed.
              setTimeout(() => focusFirst(candidates));
            },
          };

      const itemContextValue = React.useMemo<ThumbnailItemContextValue>(
        () => ({ slideId, displayIndex, isActive }),
        [slideId, displayIndex, isActive],
      );

      return (
        <ThumbnailItemContext.Provider value={itemContextValue}>
          {renderElement(
            "button",
            { render },
            {
              state,
              ref: [registerRef, forwardedRef],
              props: [
                {
                  type: "button",
                  "data-active": isActive || undefined,
                  "data-slide-id": slideId,
                  style: { width: "100%" },
                  ...optionProps,
                  children,
                },
                thumbnailItemProps,
              ],
            },
          )}
        </ThumbnailItemContext.Provider>
      );
    },
  ),
);

export namespace ThumbnailItem {
  export type State = ThumbnailItemState;
  export type Props = ThumbnailItemProps;
}

export interface ThumbnailItemPreviewState {
  /** Stable id of the slide being rendered. */
  slideId: string;

  /** CSS scale factor from presentation width to container width, `0` until measured. */
  scale: number;
}

export interface ThumbnailItemPreviewProps extends PrimitiveProps<
  "div",
  ThumbnailItemPreviewState
> {}

export const ThumbnailItemPreview = React.forwardRef<
  HTMLDivElement,
  ThumbnailItemPreviewProps
>(function ThumbnailItemPreview(
  { render, ...thumbnailItemPreviewProps },
  forwardedRef,
) {
  const itemContext = useThumbnailItemContext(THUMBNAIL_ITEM_PREVIEW_NAME);
  const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_PREVIEW_NAME);
  const { presentation } = usePresentation();
  const store = useStoreContext(THUMBNAIL_ITEM_PREVIEW_NAME);

  const itemPreviewRef = React.useRef<HTMLDivElement>(null);
  const slideHandleRef = React.useRef<SlideHandle | null>(null);
  // False when another mount held the cache slot, so this preview owns a private copy.
  const ownsCachedHandleRef = React.useRef(false);
  const hasRenderedRef = React.useRef(false);
  const attachedSlideIdRef = React.useRef<string | null>(null);
  const attachedRevisionRef = React.useRef<number | null>(null);
  // A teardown that sees a newer run knows it was a re-run rather than an unmount.
  const observerRunRef = React.useRef(0);
  const { mediaUrlCache, slideHandleCache, observeResize, scheduleRender } =
    rovingContext;

  const slideIndex = store.getSlideIndex(itemContext.slideId);
  const slide = presentation?.slides[slideIndex] ?? null;
  const presentationWidth = presentation?.width ?? 1;
  const presentationHeight = presentation?.height ?? 1;

  // In a ref so an edit does not tear down the observer effect, whose detach and re-attach flashes.
  const revision = useSlideRevision(store, itemContext.slideId);
  const revisionRef = useLatestRef(revision);

  const widthRef = React.useRef(0);
  const [containerWidth, setContainerWidth] = React.useState(0);
  const hasRenderPropRef = useLatestRef(render != null);
  const presentationWidthRef = useLatestRef(presentationWidth);
  const scale = containerWidth > 0 ? containerWidth / presentationWidth : 0;

  React.useLayoutEffect(() => {
    const itemPreviewElement = itemPreviewRef.current;
    if (itemPreviewElement && itemPreviewElement.offsetWidth > 0) {
      widthRef.current = itemPreviewElement.offsetWidth;
      if (hasRenderPropRef.current)
        setContainerWidth(itemPreviewElement.offsetWidth);
    }
  }, [hasRenderPropRef]);

  React.useEffect(() => {
    const itemPreviewElement = itemPreviewRef.current;
    if (!itemPreviewElement) return;
    return observeResize(itemPreviewElement, (width) => {
      widthRef.current = width;
      const nextScale = width > 0 ? width / presentationWidthRef.current : 0;
      if (slideHandleRef.current && nextScale > 0)
        applySlideScale(slideHandleRef.current.element, nextScale);
      if (hasRenderPropRef.current) setContainerWidth(width);
    });
  }, [observeResize, hasRenderPropRef, presentationWidthRef]);

  React.useEffect(() => {
    const itemPreviewElement = itemPreviewRef.current;
    const run = ++observerRunRef.current;
    if (!itemPreviewElement || !presentation || !slide) return;
    if (typeof IntersectionObserver === "undefined") return;

    const attach = (
      element: HTMLDivElement,
      slideHandle: SlideHandle,
      isCached: boolean,
    ) => {
      if (slideHandleRef.current === slideHandle) return;
      const currentScale =
        widthRef.current > 0
          ? widthRef.current / presentationWidthRef.current
          : 0;
      if (currentScale > 0) applySlideScale(slideHandle.element, currentScale);
      element.appendChild(slideHandle.element);
      slideHandleRef.current = slideHandle;
      attachedSlideIdRef.current = slide.id;
      attachedRevisionRef.current = revisionRef.current;
      ownsCachedHandleRef.current = isCached;
      hasRenderedRef.current = true;
      delete element.dataset.pending;
    };

    const detach = (element: HTMLDivElement) => {
      const slideHandle = slideHandleRef.current;
      slideHandleRef.current = null;
      attachedSlideIdRef.current = null;
      slideHandle?.element.remove();
      // The cached handle is kept alive for the next mount; a private copy
      // has no other owner, so it is released here.
      if (slideHandle && !ownsCachedHandleRef.current) slideHandle.dispose();
      ownsCachedHandleRef.current = false;
      hasRenderedRef.current = false;
      element.dataset.pending = "";
    };

    // A previous run can leave a miniature attached; it stays only for the same slide.
    if (slideHandleRef.current && attachedSlideIdRef.current !== slide.id) {
      detach(itemPreviewElement);
    }

    let cancelRender: (() => void) | null = null;

    const intersectionObserver = new IntersectionObserver(
      (entries) => {
        const entry = entries.at(-1);
        if (!entry) return;
        const element = itemPreviewRef.current;
        if (!element) return;

        // A render queued by a previous cycle must not outlive re-attaching or
        // detaching.
        cancelRender?.();
        cancelRender = null;

        if (entry.isIntersecting) {
          const cached = slideHandleCache.get(slide.id);
          const isCurrent =
            cached !== undefined && cached.revision === revisionRef.current;
          // A cached node that is already inside another preview belongs to
          // that mount; moving it here would blank it out.
          const parent = cached?.slideHandle.element.parentElement ?? null;
          const isAvailable = parent === null || parent === element;

          if (isCurrent && isAvailable) {
            markThumbnailRecent(slideHandleCache, slide.id);
            attach(element, cached.slideHandle, true);
          } else {
            // Anything still attached here is stale, and it has to go before
            // its cache entry is retired or the re-render below is skipped.
            if (slideHandleRef.current) detach(element);
            if (cached && !isCurrent && isAvailable) {
              cached.slideHandle.dispose();
              slideHandleCache.delete(slide.id);
            }
            cancelRender = scheduleRender(() => {
              cancelRender = null;
              const mountedElement = itemPreviewRef.current;
              if (!mountedElement || slideHandleRef.current) return;
              const slideHandle = renderSlide(presentation, slide, {
                mediaUrlCache,
              });
              const isCached = getIsThumbnailCached(
                slideHandleCache,
                slide.id,
                slideHandle,
                revisionRef.current,
              );
              attach(mountedElement, slideHandle, isCached);
            });
          }
        } else {
          detach(element);
        }
      },
      { rootMargin: INTERSECTION_OBSERVER_ROOT_MARGIN },
    );

    intersectionObserver.observe(itemPreviewElement);

    return () => {
      intersectionObserver.disconnect();
      cancelRender?.();
      cancelRender = null;
      // An unmount and a re-run look the same here, and detaching on a re-run flashes the
      // miniature, so the decision waits until a re-run has started its next run.
      queueMicrotask(() => {
        // oxlint-disable-next-line react-hooks/exhaustive-deps -- reading the latest value is the point: it tells a re-run from an unmount
        if (observerRunRef.current !== run) return;
        detach(itemPreviewElement);
      });
    };
  }, [
    presentation,
    slide,
    mediaUrlCache,
    slideHandleCache,
    scheduleRender,
    revisionRef,
    presentationWidthRef,
  ]);

  // Off-screen slides re-render when the observer next reports them.
  React.useEffect(() => {
    const element = itemPreviewRef.current;
    if (!element || !presentation || !slide) return;
    if (!slideHandleRef.current) return;
    // Effects are re-created whenever React restores this subtree, so an
    // unchanged revision means the attached miniature is already current.
    if (attachedRevisionRef.current === revision) return;

    const oldHandle = slideHandleRef.current;
    oldHandle.element.remove();
    oldHandle.dispose();
    slideHandleRef.current = null;
    // Only the mount that published the entry may retire it; a private copy
    // would otherwise remove a node still attached to another preview.
    if (ownsCachedHandleRef.current) slideHandleCache.delete(slide.id);
    ownsCachedHandleRef.current = false;

    const slideHandle = renderSlide(presentation, slide, { mediaUrlCache });
    const currentScale =
      widthRef.current > 0
        ? widthRef.current / presentationWidthRef.current
        : 0;
    if (currentScale > 0) applySlideScale(slideHandle.element, currentScale);
    element.appendChild(slideHandle.element);
    slideHandleRef.current = slideHandle;
    attachedRevisionRef.current = revision;
    ownsCachedHandleRef.current = getIsThumbnailCached(
      slideHandleCache,
      slide.id,
      slideHandle,
      revision,
    );
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- revision-only: the IO effect above handles presentation and slide changes, and every visual edit bumps revision
  }, [revision]);

  return renderElement(
    "div",
    { render },
    {
      state: { slideId: itemContext.slideId, scale },
      ref: [itemPreviewRef, forwardedRef],
      props: [
        {
          "aria-hidden": "true",
          "data-active": itemContext.isActive || undefined,
          "data-pending": hasRenderedRef.current ? undefined : "",
          inert: true,
          style: {
            width: "100%",
            aspectRatio: `${presentationWidth} / ${presentationHeight}`,
            overflow: "hidden",
            pointerEvents: "none",
            contain: "layout style paint",
          },
        },
        thumbnailItemPreviewProps,
      ],
    },
  );
});

export namespace ThumbnailItemPreview {
  export type State = ThumbnailItemPreviewState;
  export type Props = ThumbnailItemPreviewProps;
}

export interface ThumbnailItemNumberProps extends PrimitiveProps<
  "span",
  { isActive: boolean; displayIndex: number; slideId: string }
> {}

export const ThumbnailItemNumber = React.forwardRef<
  HTMLSpanElement,
  ThumbnailItemNumberProps
>(function ThumbnailItemNumber(
  { children, render, ...thumbnailItemNumberProps },
  forwardedRef,
) {
  const itemContext = useThumbnailItemContext(THUMBNAIL_ITEM_NUMBER_NAME);

  return renderElement(
    "span",
    { render },
    // oxlint-disable-next-line react/refs -- renderElement attaches the ref without reading it
    {
      state: {
        isActive: itemContext.isActive,
        displayIndex: itemContext.displayIndex,
        slideId: itemContext.slideId,
      },
      ref: forwardedRef,
      props: [
        {
          "data-active": itemContext.isActive || undefined,
          style: { userSelect: "none" },
          children: children ?? itemContext.displayIndex + 1,
        },
        thumbnailItemNumberProps,
      ],
    },
  );
});

export namespace ThumbnailItemNumber {
  export type Props = ThumbnailItemNumberProps;
}
