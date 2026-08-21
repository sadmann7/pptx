/**
 * Roving focus implementation adapted from Radix UI's `@radix-ui/react-roving-focus`.
 *
 * The keyboard-navigation model (tab-stop management, focus-intent mapping,
 * `focusFirst` helper, and click-vs-keyboard distinction) is directly inspired
 * by that package.
 *
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
import { useLatestRef, useLazyRef } from "./hook";
import type { RenderProp } from "./render";
import { renderElement } from "./render";

const THUMBNAIL_LIST_NAME = "Presentation.ThumbnailList";
const THUMBNAIL_ITEM_NAME = "Presentation.ThumbnailItem";
const THUMBNAIL_ITEM_PREVIEW_NAME = "Presentation.ThumbnailItemPreview";
const THUMBNAIL_ITEM_NUMBER_NAME = "Presentation.ThumbnailItemNumber";

/**
 * IntersectionObserver rootMargin applied when observing each preview frame.
 * Pre-renders thumbnails this many px before they scroll into view so normal
 * scrolling never reveals a pending placeholder.
 */
const INTERSECTION_OBSERVER_ROOT_MARGIN = "200px 0px";

interface CachedThumbnail {
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

function wrapArray<T>(array: T[], startIndex: number): T[] {
  return array.map((_, i) => array[(startIndex + i) % array.length] as T);
}

/**
 * Publish `slideHandle` as the shared miniature for `slideId`.
 *
 * Returns `false` when a live entry for the current revision already holds the
 * slot, which happens while the same slide is mounted twice (a drag overlay
 * next to its list item). The caller then holds a private copy that it must
 * dispose itself, so the two previews never move one DOM node between them.
 */
function getIsThumbnailCached(
  slideHandleCache: Map<string, CachedThumbnail>,
  slideId: string,
  slideHandle: SlideHandle,
  revision: number,
): boolean {
  const cached = slideHandleCache.get(slideId);
  if (cached) {
    if (cached.slideHandle !== slideHandle && cached.revision === revision) return false;
    cached.slideHandle.dispose();
  }
  slideHandleCache.set(slideId, { slideHandle, revision });
  return true;
}

interface ThumbnailRovingContextValue {
  /**
   * Tab-stop tracking lives in a mini external store instead of context
   * state: putting the roving id in the context value would re-render every
   * memoized item on each focus change, when only two items (old and new tab
   * stop) actually need to update.
   */
  subscribeTabStop: (callback: () => void) => () => void;
  getEffectiveTabStopId: () => string | null;
  onItemFocus: (slideId: string) => void;
  onItemRegister: (slideId: string, el: HTMLButtonElement) => void;
  onItemUnregister: (slideId: string) => void;
  itemsRef: React.RefObject<Map<string, HTMLButtonElement>>;
  loop: boolean;
  /** Shared object-URL cache so each image is decoded once across all previews. */
  mediaUrlCache: Map<string, string>;
  /**
   * Rendered slide DOM cache, keyed by slide id. Scrolling back re-attaches
   * the existing element instantly instead of re-rendering. Entries are
   * invalidated when the slide's edit revision moves past the cached one.
   */
  slideHandleCache: Map<string, CachedThumbnail>;
  /**
   * Register with the list-level shared ResizeObserver.
   * Returns a cleanup function that unregisters the element.
   */
  observeResize: (element: Element, cb: (width: number) => void) => () => void;
  /**
   * Enqueue a `renderSlide()` call, drained FIFO per animation frame
   * within an ~8ms budget so no single frame blocks the main thread.
   */
  scheduleRender: (fn: () => void) => () => void;
}

const ThumbnailRovingContext = React.createContext<ThumbnailRovingContextValue | null>(null);

function useThumbnailRovingContext(consumerName: string) {
  const context = React.useContext(ThumbnailRovingContext);
  if (!context) {
    throw new Error(`\`${consumerName}\` must be used within \`${THUMBNAIL_LIST_NAME}\``);
  }
  return context;
}

interface ThumbnailItemContextValue {
  slideId: string;
  displayIndex: number;
  isActive: boolean;
}

const ThumbnailItemContext = React.createContext<ThumbnailItemContextValue | null>(null);

function useThumbnailItemContext(consumerName: string) {
  const context = React.useContext(ThumbnailItemContext);
  if (!context) {
    throw new Error(`\`${consumerName}\` must be used within \`${THUMBNAIL_ITEM_NAME}\``);
  }
  return context;
}

export interface ThumbnailListState {
  /** Total number of slides in the loaded presentation. */
  total: number;

  /** Stable id of the currently active slide, or `null` before load. */
  activeSlideId: string | null;

  /** 0-based position of the active slide. Derived from `activeSlideId`. */
  activeIndex: number;
}

export interface ThumbnailListRenderState {
  /** All slides in the loaded presentation, in order. */
  slides: SlideData[];

  /** Stable id of the currently active slide, or `null` before load. */
  activeSlideId: string | null;

  /** 0-based position of the active slide. Derived from `activeSlideId`. */
  activeIndex: number;

  /** Navigate to a slide by its stable id. */
  goTo: (slideId: string) => void;

  /** Navigate to a slide by its 0-based index. */
  goToIndex: (index: number) => void;
}

export interface ThumbnailListProps extends Omit<React.ComponentProps<"div">, "children"> {
  /**
   * Replace the list container element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<ThumbnailListState>;

  /**
   * - Absent → default `ThumbnailItem` list (one per slide)
   * - ReactNode → rendered as-is inside the container
   * - Function → called with slide state when ready
   */
  children?: React.ReactNode | ((state: ThumbnailListRenderState) => React.ReactNode);

  /**
   * When `true`, keyboard navigation wraps from the last item back to the
   * first (and vice versa).
   *
   * @default false
   */
  loop?: boolean;
}

/**
 * Scrollable `listbox` container listing all slide thumbnails.
 * Renders nothing until the presentation is `"ready"`.
 *
 * All slide buttons are mounted immediately (cheap empty containers).
 * Thumbnail content is rendered lazily; each preview uses an
 * IntersectionObserver with a generous `rootMargin` so content fills in
 * before the element scrolls into view. For normal scrolling the transition
 * is invisible; rapid drag may briefly reveal a pending placeholder.
 *
 * Handles keyboard navigation (↑↓ / Home / End) with roving focus so only
 * the active item lives in the tab order at any time.
 */
export const ThumbnailList = React.forwardRef<HTMLDivElement, ThumbnailListProps>(
  function ThumbnailList({ render, children, loop = false, ...thumbnailListProps }, forwardedRef) {
    const { presentation, status } = usePresentation();
    const store = useStoreContext(THUMBNAIL_LIST_NAME);

    const itemsRef = useLazyRef(() => new Map<string, HTMLButtonElement>());
    const isClickFocusRef = React.useRef(false);
    const autoFocusedPresentationRef = React.useRef<object | null>(null);

    // A ref and listener set, not react state: state would flow through the
    // context value and re-render every memoized item on each focus change,
    // when only two items (the old and new tab stop) actually need to update.
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

    const mediaCacheRef = React.useRef<{ key: object; cache: Map<string, string> } | null>(null);
    if (!mediaCacheRef.current || mediaCacheRef.current.key !== presentation) {
      mediaCacheRef.current = { key: presentation ?? {}, cache: new Map() };
    }
    const mediaUrlCache = mediaCacheRef.current.cache;

    const slideHandleCacheRef = React.useRef<{
      key: object;
      cache: Map<string, CachedThumbnail>;
    } | null>(null);
    if (!slideHandleCacheRef.current || slideHandleCacheRef.current.key !== presentation) {
      slideHandleCacheRef.current = { key: presentation ?? {}, cache: new Map() };
    }
    const slideHandleCache = slideHandleCacheRef.current.cache;

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

    const resizeCallbacksRef = useLazyRef(() => new Map<Element, (width: number) => void>());
    const sharedResizeObserverRef = useLazyRef(() =>
      typeof ResizeObserver !== "undefined"
        ? new ResizeObserver((entries) => {
            for (const entry of entries) {
              resizeCallbacksRef.current.get(entry.target)?.(entry.contentRect.width);
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

    // Auto-focus the active (or first) thumbnail once per presentation load.
    React.useEffect(() => {
      if (!presentation || autoFocusedPresentationRef.current === presentation) return;
      autoFocusedPresentationRef.current = presentation;
      const items = itemsRef.current;
      const activeItem = activeSlideId ? items.get(activeSlideId) : undefined;
      const firstItem = activeItem ?? items.values().next().value;
      firstItem?.focus({ preventScroll: true });
    }, [presentation, activeSlideId, itemsRef]);

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
        onItemRegister: (slideId, element) => itemsRef.current.set(slideId, element),
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
                "aria-orientation": "vertical",
                tabIndex: hasTabStop ? -1 : 0,
                onMouseDown: (event: React.MouseEvent<HTMLDivElement>) => {
                  if (event.target === event.currentTarget) isClickFocusRef.current = true;
                },
                onFocus: (event: React.FocusEvent<HTMLDivElement>) => {
                  if (event.target !== event.currentTarget) return;
                  if (isClickFocusRef.current) {
                    isClickFocusRef.current = false;
                    return;
                  }
                  focusFirst(Array.from(itemsRef.current.values()), true);
                },
                children: resolvedChildren,
              },
              thumbnailListProps,
            ],
          },
        )}
      </ThumbnailRovingContext.Provider>
    );
  },
);

export namespace ThumbnailList {
  export type State = ThumbnailListState;
  export type RenderState = ThumbnailListRenderState;
  export type Props = ThumbnailListProps;
}

export interface ThumbnailItemState {
  /**
   * Stable id of the slide this item represents (`SlideData.id`).
   */
  slideId: string;

  /**
   * `true` when this item's slide is the currently active slide.
   */
  isActive: boolean;

  /**
   * 0-based position of this slide in the presentation.
   */
  displayIndex: number;
}

/** Fired when a thumbnail is about to become the active slide. */
export interface ThumbnailSelectEvent {
  /** Stable id of the slide the list is about to navigate to. */
  slideId: string;

  /** Call to keep the navigation from happening. */
  preventDefault: () => void;
}

export interface ThumbnailItemProps extends Omit<React.ComponentProps<"button">, "onSelect"> {
  slideId: string;
  render?: RenderProp<ThumbnailItemState>;

  /**
   * Runs just before this item becomes the active slide, whether from a click
   * or from keyboard roving (arrowing through the list changes slides, so it
   * triggers there too). Call `preventDefault()` to stop the navigation.
   *
   * This exists because the list navigates on focus, which the browser
   * delivers before `click`: a veto expressed in `onClick` would arrive after
   * the slide had already changed. Use `onClick` for side effects that do not
   * need to block navigation, such as analytics.
   *
   * Note that `ThumbnailItem` is memoized: passing a handler whose identity
   * changes every render opts the item out of that memoization.
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
   * Render the item's visuals only: no listbox role, no roving focus
   * registration, no navigation on click or focus, and no presence in the
   * accessibility tree.
   *
   * Use it for a second copy of an item that is already in the list, such as
   * the floating thumbnail a drag overlay renders while reordering. Without
   * it the copy claims the original's slot in the roving focus map and
   * unregisters it on unmount, and the deck is announced with a duplicate
   * option.
   *
   * @default false
   */
  decorative?: boolean;
}

/**
 * Clickable `option` button for a single slide in a `ThumbnailList`.
 */
export const ThumbnailItem = React.memo(
  React.forwardRef<HTMLButtonElement, ThumbnailItemProps>(function ThumbnailItem(
    { slideId, children, render, onSelect, decorative = false, ...thumbnailItemProps },
    forwardedRef,
  ) {
    const store = useStoreContext(THUMBNAIL_ITEM_NAME);
    const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_NAME);

    const isActive = useStoreSelector(store, (s) => s.activeSlideId === slideId, false);

    const displayIndex = useSlideIndex(store, slideId);

    const setSize = useStoreSelector(store, (s) => s.presentation?.slides.length ?? 0, 0);

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

    /**
     * Single navigation path for both click and focus, so a consumer's veto
     * applies no matter which one got there first.
     */
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

    /**
     * Listbox behaviour, dropped for a decorative copy so it stays out of
     * roving focus, the accessibility tree, and navigation.
     */
    const optionProps: React.ComponentProps<"button"> = decorative
      ? { "aria-hidden": true, inert: true, tabIndex: -1 }
      : {
          role: "option",
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
            if (!focusIntent || event.metaKey || event.ctrlKey || event.altKey || event.shiftKey)
              return;
            event.preventDefault();

            let candidates = Array.from(rovingContext.itemsRef.current.values());
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
  }),
);

export namespace ThumbnailItem {
  export type State = ThumbnailItemState;
  export type Props = ThumbnailItemProps;
}

export interface ThumbnailItemPreviewState {
  /**
   * Stable id of the slide being rendered.
   */
  slideId: string;

  /**
   * CSS scale factor (container width / presentation width). `0` before measured.
   */
  scale: number;
}

export interface ThumbnailItemPreviewProps extends React.ComponentProps<"div"> {
  render?: RenderProp<ThumbnailItemPreviewState>;
}

/**
 * Renders the slide miniature for the enclosing `ThumbnailItem`.
 *
 * Uses an IntersectionObserver with a 200 px vertical rootMargin so
 * `renderSlide()` is called slightly before the element scrolls into view.
 * For normal scrolling the thumbnail is always ready before it's visible.
 * Rapid scrollbar drag may briefly show a pending placeholder; the same
 * behaviour as the reference vanilla implementation.
 *
 * Rendered DOM is kept in the list's handle cache: scrolling back re-attaches
 * the existing element instantly. The cache is cleared when the presentation
 * changes or the list unmounts.
 */
export const ThumbnailItemPreview = React.forwardRef<HTMLDivElement, ThumbnailItemPreviewProps>(
  function ThumbnailItemPreview({ render, ...thumbnailItemPreviewProps }, forwardedRef) {
    const itemContext = useThumbnailItemContext(THUMBNAIL_ITEM_PREVIEW_NAME);
    const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_PREVIEW_NAME);
    const { presentation } = usePresentation();
    const store = useStoreContext(THUMBNAIL_ITEM_PREVIEW_NAME);

    const itemPreviewRef = React.useRef<HTMLDivElement>(null);
    const slideHandleRef = React.useRef<SlideHandle | null>(null);
    /**
     * Whether the attached handle is the shared cached one. `false` means this
     * preview rendered a private copy because another mount already held the
     * cache slot, and owns its disposal.
     */
    const ownsCachedHandleRef = React.useRef(false);
    const hasRenderedRef = React.useRef(false);
    const attachedSlideIdRef = React.useRef<string | null>(null);
    const attachedRevisionRef = React.useRef<number | null>(null);
    /**
     * Incremented by every run of the observer effect below. A teardown that
     * sees a newer run knows it was a re-run rather than an unmount.
     */
    const observerRunRef = React.useRef(0);
    const { mediaUrlCache, slideHandleCache, observeResize, scheduleRender } = rovingContext;

    const slideIndex = store.getSlideIndex(itemContext.slideId);
    const slide = presentation?.slides[slideIndex] ?? null;
    const presentationWidth = presentation?.width ?? 1;
    const presentationHeight = presentation?.height ?? 1;

    // Edit revision of this slide; a bump means the cached miniature is
    // stale. Kept in a ref so the IO effect doesn't tear down on every edit
    // (that detach→re-attach gap is what causes the thumbnail flash).
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
        if (hasRenderPropRef.current) setContainerWidth(itemPreviewElement.offsetWidth);
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

      const attach = (element: HTMLDivElement, slideHandle: SlideHandle, isCached: boolean) => {
        if (slideHandleRef.current === slideHandle) return;
        const currentScale =
          widthRef.current > 0 ? widthRef.current / presentationWidthRef.current : 0;
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

      // A previous run of this effect can leave a miniature attached (see the
      // teardown below). Keep it only when it is still the right slide.
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

          // Cancel any pending render queued by a previous cycle before either
          // re-attaching or detaching.
          cancelRender?.();
          cancelRender = null;

          if (entry.isIntersecting) {
            const cached = slideHandleCache.get(slide.id);
            const isCurrent = cached !== undefined && cached.revision === revisionRef.current;
            // A cached node that is already inside another preview belongs to
            // that mount; moving it here would blank it out.
            const parent = cached?.slideHandle.element.parentElement ?? null;
            const isAvailable = parent === null || parent === element;

            if (isCurrent && isAvailable) {
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
                const slideHandle = renderSlide(presentation, slide, { mediaUrlCache });
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
        // An unmount and a re-run of this effect are indistinguishable here,
        // and detaching on a re-run blanks the miniature for the frame it takes
        // a fresh observer to report, which reads as a flash. Defer the
        // decision: a re-run has already started the next run by now.
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

    // Re-render the miniature in place when an edit bumps the revision,
    // without tearing down the IntersectionObserver (the detach→re-attach
    // gap was the cause of the thumbnail flash on every edit). Only triggers
    // when the slide is currently visible/attached; off-screen slides are
    // re-rendered on demand by the IntersectionObserver callback using revisionRef.
    React.useEffect(() => {
      const element = itemPreviewRef.current;
      if (!element || !presentation || !slide) return;
      if (!slideHandleRef.current) return; // not visible; IO handles it
      // Effects are re-created whenever React restores this subtree, so an
      // unchanged revision means the attached miniature is already current.
      if (attachedRevisionRef.current === revision) return;

      const oldHandle = slideHandleRef.current;
      oldHandle.element.remove();
      oldHandle.dispose();
      slideHandleRef.current = null;
      // Only the mount that published the entry may retire it; a private copy
      // would otherwise evict a node still attached to another preview.
      if (ownsCachedHandleRef.current) slideHandleCache.delete(slide.id);
      ownsCachedHandleRef.current = false;

      const slideHandle = renderSlide(presentation, slide, { mediaUrlCache });
      const currentScale =
        widthRef.current > 0 ? widthRef.current / presentationWidthRef.current : 0;
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
  },
);

export namespace ThumbnailItemPreview {
  export type State = ThumbnailItemPreviewState;
  export type Props = ThumbnailItemPreviewProps;
}

export interface ThumbnailItemNumberProps extends React.ComponentProps<"span"> {
  render?: RenderProp<{ isActive: boolean; displayIndex: number; slideId: string }>;
}

/**
 * Renders the 1-based slide number for the enclosing `ThumbnailItem`.
 */
export const ThumbnailItemNumber = React.forwardRef<HTMLSpanElement, ThumbnailItemNumberProps>(
  function ThumbnailItemNumber({ children, render, ...thumbnailItemNumberProps }, forwardedRef) {
    const itemContext = useThumbnailItemContext(THUMBNAIL_ITEM_NUMBER_NAME);

    return renderElement(
      "span",
      { render },
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
  },
);

export namespace ThumbnailItemNumber {
  export type Props = ThumbnailItemNumberProps;
}
