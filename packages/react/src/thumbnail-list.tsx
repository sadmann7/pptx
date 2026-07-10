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

import type { SlideData, SlideHandle } from "@diceui/pptx-parser";
import { renderSlide } from "@diceui/pptx-parser";

import { TYPOGRAPHY_RESET_STYLE, VISUALLY_HIDDEN_STYLE } from "./constant";
import { usePresentation, usePresentationStore } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";

const THUMBNAIL_LIST_NAME = "PresentationThumbnailList";
const THUMBNAIL_ITEM_NAME = "PresentationThumbnailItem";
const THUMBNAIL_ITEM_PREVIEW_NAME = "PresentationThumbnailItemPreview";
const THUMBNAIL_ITEM_NUMBER_NAME = "PresentationThumbnailItemNumber";

/**
 * IntersectionObserver rootMargin applied when observing each preview frame.
 * Pre-renders thumbnails this many px before they scroll into view so normal
 * scrolling never reveals a pending placeholder.
 */
const IO_ROOT_MARGIN = "200px 0px";

type FocusIntent = "first" | "last" | "prev" | "next";

/** Cached rendered thumbnail plus the edit revision it was rendered at. */
interface CachedThumbnail {
  handle: SlideHandle;
  revision: number;
}

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
  observeResize: (el: Element, cb: (width: number) => void) => () => void;
  /**
   * Enqueue a `renderThumbnail()` call, drained FIFO per animation frame
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
 * Thumbnail content is rendered lazily — each preview uses an
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
    const store = usePresentationStore(THUMBNAIL_LIST_NAME);

    const itemsRef = React.useRef<Map<string, HTMLButtonElement>>(new Map());
    const isClickFocusRef = React.useRef(false);
    const loopRef = React.useRef(loop);
    loopRef.current = loop;

    // --- Roving tab stop (mini external store) ---
    // A ref + listener set, NOT React state: state would flow through the
    // context value and re-render every memoized item on each focus change,
    // when only two items (the old and new tab stop) actually need to update.
    const currentTabStopIdRef = React.useRef<string | null>(null);
    const tabStopListenersRef = React.useRef(new Set<() => void>());

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
      [store],
    );

    const setCurrentTabStopId = React.useCallback((slideId: string | null) => {
      if (currentTabStopIdRef.current === slideId) return;
      currentTabStopIdRef.current = slideId;
      for (const callback of tabStopListenersRef.current) callback();
    }, []);

    // --- Shared caches, keyed by presentation identity ---
    const mediaCacheRef = React.useRef<{ key: object; cache: Map<string, string> } | null>(null);
    if (!mediaCacheRef.current || mediaCacheRef.current.key !== presentation) {
      mediaCacheRef.current = { key: presentation ?? {}, cache: new Map() };
    }
    const mediaUrlCache = mediaCacheRef.current.cache;

    const handleCacheRef = React.useRef<{
      key: object;
      cache: Map<string, CachedThumbnail>;
    } | null>(null);
    if (!handleCacheRef.current || handleCacheRef.current.key !== presentation) {
      handleCacheRef.current = { key: presentation ?? {}, cache: new Map() };
    }
    const slideHandleCache = handleCacheRef.current.cache;

    React.useEffect(() => {
      const handles = slideHandleCache;
      const media = mediaUrlCache;
      return () => {
        for (const entry of handles.values()) entry.handle.dispose();
        handles.clear();
        for (const url of media.values()) URL.revokeObjectURL(url);
        media.clear();
      };
    }, [slideHandleCache, mediaUrlCache]);

    // --- Shared ResizeObserver ---
    const resizeCallbacksRef = React.useRef(new Map<Element, (width: number) => void>());
    const sharedRORef = React.useRef<ResizeObserver | null>(null);
    if (!sharedRORef.current && typeof ResizeObserver !== "undefined") {
      sharedRORef.current = new ResizeObserver((entries) => {
        for (const entry of entries) {
          resizeCallbacksRef.current.get(entry.target)?.(entry.contentRect.width);
        }
      });
    }

    const observeResize = React.useCallback((el: Element, cb: (width: number) => void) => {
      resizeCallbacksRef.current.set(el, cb);
      sharedRORef.current?.observe(el);
      return () => {
        resizeCallbacksRef.current.delete(el);
        sharedRORef.current?.unobserve(el);
      };
    }, []);

    // --- Central render queue ---
    // Batches renderThumbnail() calls that arrive simultaneously (e.g. initial
    // viewport fills, rapid scroll) and drains them within an ~8ms per-frame
    // budget so no single commit blocks the main thread.
    const renderQueueRef = React.useRef<Array<() => void>>([]);
    const drainRafRef = React.useRef<number | null>(null);

    const drainRenderQueue = React.useCallback(function drain() {
      drainRafRef.current = null;
      const queue = renderQueueRef.current;
      if (queue.length === 0) return;
      const budgetMs = 8;
      const frameStart = performance.now();
      do {
        queue.shift()?.();
      } while (queue.length > 0 && performance.now() - frameStart < budgetMs);
      if (queue.length > 0) drainRafRef.current = requestAnimationFrame(drain);
    }, []);

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
      [drainRenderQueue],
    );

    React.useEffect(() => {
      return () => {
        sharedRORef.current?.disconnect();
        sharedRORef.current = null;
        if (drainRafRef.current !== null) {
          cancelAnimationFrame(drainRafRef.current);
          drainRafRef.current = null;
        }
        renderQueueRef.current = [];
      };
    }, []);

    const activeSlideId = React.useSyncExternalStore(
      store.subscribe,
      () => store.getState().activeSlideId,
      () => null,
    );

    // Auto-focus the active (or first) thumbnail ONCE per presentation load.
    const autoFocusedPresentationRef = React.useRef<object | null>(null);
    React.useEffect(() => {
      if (!presentation || autoFocusedPresentationRef.current === presentation) return;
      autoFocusedPresentationRef.current = presentation;
      const items = itemsRef.current;
      const activeItem = activeSlideId ? items.get(activeSlideId) : undefined;
      const firstItem = activeItem ?? items.values().next().value;
      firstItem?.focus({ preventScroll: true });
    }, [presentation, activeSlideId]);

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
        onItemRegister: (slideId, el) => itemsRef.current.set(slideId, el),
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
        loop,
        mediaUrlCache,
        slideHandleCache,
        observeResize,
        scheduleRender,
      ],
    );

    if (status !== "ready" || !presentation) return null;

    const total = presentation.slides.length;
    const activeIndex = activeSlideId
      ? presentation.slides.findIndex((s) => s.id === activeSlideId)
      : -1;

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
                "aria-label": "Slide thumbnails",
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
  /** Stable id of the slide this item represents (`SlideData.id`). */
  slideId: string;

  /** `true` when this item's slide is the currently active slide. */
  isActive: boolean;

  /** 0-based position of this slide in the presentation. */
  displayIndex: number;
}

export interface ThumbnailItemProps extends Omit<React.ComponentProps<"button">, "onClick"> {
  slideId: string;
  render?: RenderProp<ThumbnailItemState>;
}

/**
 * Clickable `option` button for a single slide in a `ThumbnailList`.
 */
export const ThumbnailItem = React.memo(
  React.forwardRef<HTMLButtonElement, ThumbnailItemProps>(function ThumbnailItem(
    { slideId, children, render, ...thumbnailItemProps },
    forwardedRef,
  ) {
    const store = usePresentationStore(THUMBNAIL_ITEM_NAME);
    const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_NAME);

    const isActive = React.useSyncExternalStore(
      store.subscribe,
      () => store.getState().activeSlideId === slideId,
      () => false,
    );

    const displayIndex = React.useSyncExternalStore(
      store.subscribe,
      () => store.getSlideIndex(slideId),
      () => -1,
    );

    const setSize = React.useSyncExternalStore(
      store.subscribe,
      () => store.getState().presentation?.slides.length ?? 0,
      () => 0,
    );

    const { onItemRegister, onItemUnregister } = rovingContext;
    const registerRef = React.useCallback(
      (element: HTMLButtonElement | null) => {
        if (element) onItemRegister(slideId, element);
        else onItemUnregister(slideId);
      },
      [slideId, onItemRegister, onItemUnregister],
    );

    const isCurrentTabStop = React.useSyncExternalStore(
      rovingContext.subscribeTabStop,
      () => rovingContext.getEffectiveTabStopId() === slideId,
      () => false,
    );

    const state: ThumbnailItemState = { slideId, isActive, displayIndex };

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
                role: "option",
                "aria-selected": isActive,
                "aria-label": `Slide ${displayIndex + 1}`,
                "aria-posinset": displayIndex + 1,
                "aria-setsize": setSize,
                "data-active": isActive || undefined,
                "data-slide-id": slideId,
                tabIndex: isCurrentTabStop ? 0 : -1,
                style: { width: "100%" },
                onClick: () => store.goTo(slideId),
                onFocus: () => {
                  rovingContext.onItemFocus(slideId);
                  store.goTo(slideId);
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
  /** Stable id of the slide being rendered. */
  slideId: string;
  /** CSS scale factor (container width / presentation width). `0` before measured. */
  scale: number;
}

export interface ThumbnailItemPreviewProps extends React.ComponentProps<"div"> {
  render?: RenderProp<ThumbnailItemPreviewState>;
}

/**
 * Renders the slide miniature for the enclosing `ThumbnailItem`.
 *
 * Uses an IntersectionObserver with a 200 px rootMargin so `renderThumbnail()`
 * is called slightly before the element scrolls into view. For normal
 * scrolling the thumbnail is always ready before it's visible. Rapid
 * scrollbar drag may briefly show a pending placeholder — the same
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
    const store = usePresentationStore(THUMBNAIL_ITEM_PREVIEW_NAME);

    const itemPreviewRef = React.useRef<HTMLDivElement>(null);
    const slideHandleRef = React.useRef<SlideHandle | null>(null);
    const hasRenderedRef = React.useRef(false);
    const { mediaUrlCache, slideHandleCache, observeResize, scheduleRender } = rovingContext;

    const slideIndex = store.getSlideIndex(itemContext.slideId);
    const slide = presentation?.slides[slideIndex] ?? null;
    const pWidth = presentation?.width ?? 1;
    const pHeight = presentation?.height ?? 1;

    // Edit revision of this slide; a bump means the cached miniature is
    // stale. Kept in a ref so the IO effect doesn't tear down on every edit
    // (that detach→re-attach gap is what causes the thumbnail flash).
    const revision = React.useSyncExternalStore(
      store.subscribe,
      () => store.getSlideRevision(itemContext.slideId),
      () => 0,
    );
    const revisionRef = React.useRef(revision);
    revisionRef.current = revision;

    const widthRef = React.useRef(0);
    const [containerWidth, setContainerWidth] = React.useState(0);
    const hasRenderPropRef = React.useRef(false);
    hasRenderPropRef.current = render != null;
    const pWidthRef = React.useRef(pWidth);
    pWidthRef.current = pWidth;
    const scale = containerWidth > 0 ? containerWidth / pWidth : 0;

    React.useLayoutEffect(() => {
      const el = itemPreviewRef.current;
      if (el && el.offsetWidth > 0) {
        widthRef.current = el.offsetWidth;
        if (hasRenderPropRef.current) setContainerWidth(el.offsetWidth);
      }
    }, []);

    React.useEffect(() => {
      const el = itemPreviewRef.current;
      if (!el) return;
      return observeResize(el, (width) => {
        widthRef.current = width;
        const nextScale = width > 0 ? width / pWidthRef.current : 0;
        if (slideHandleRef.current && nextScale > 0)
          slideHandleRef.current.element.style.transform = `scale(${nextScale})`;
        if (hasRenderPropRef.current) setContainerWidth(width);
      });
    }, [observeResize]);

    // IO-GATED RENDER MODEL.
    //
    // The IntersectionObserver fires when this preview enters/leaves the
    // rootMargin zone (200px around the scroll container). On entry:
    //   - Cache hit  → re-attach synchronously (zero pending flash)
    //   - Cache miss → enqueue renderThumbnail() on the budgeted queue
    // On exit: detach DOM, keep handle in cache for instant re-attach.
    //
    // Unlike the old IO approach this does NOT use entry.boundingClientRect
    // (which was 0 in React 18 Strict Mode during the simulated remount).
    // We only use entry.isIntersecting, which is set correctly immediately.
    React.useEffect(() => {
      const el = itemPreviewRef.current;
      if (!el || !presentation || !slide) return;
      if (typeof IntersectionObserver === "undefined") return;

      const attach = (element: HTMLDivElement, handle: SlideHandle) => {
        if (slideHandleRef.current === handle) return; // already attached
        const currentScale = widthRef.current > 0 ? widthRef.current / pWidthRef.current : 0;
        if (currentScale > 0) handle.element.style.transform = `scale(${currentScale})`;
        element.appendChild(handle.element);
        slideHandleRef.current = handle;
        hasRenderedRef.current = true;
        delete element.dataset.pending;
      };

      const detach = (element: HTMLDivElement) => {
        const handle = slideHandleRef.current;
        slideHandleRef.current = null;
        handle?.element.remove();
        hasRenderedRef.current = false;
        element.dataset.pending = "";
      };

      let cancelRender: (() => void) | null = null;

      const io = new IntersectionObserver(
        (entries) => {
          const entry = entries[entries.length - 1]; // latest state wins
          if (!entry) return;
          const element = itemPreviewRef.current;
          if (!element) return;

          if (entry.isIntersecting) {
            // Cancel any pending detach-queued render from a previous cycle.
            cancelRender?.();
            cancelRender = null;

            const cached = slideHandleCache.get(slide.id);
            if (cached && cached.revision === revisionRef.current) {
              attach(element, cached.handle);
            } else {
              if (cached) {
                // Rendered under an older edit revision — discard.
                cached.handle.dispose();
                slideHandleCache.delete(slide.id);
              }
              cancelRender = scheduleRender(() => {
                cancelRender = null;
                const el2 = itemPreviewRef.current;
                if (!el2 || slideHandleRef.current) return;
                const handle = renderSlide(presentation, slide, { mediaUrlCache });
                handle.element.style.transformOrigin = "top left";
                slideHandleCache.set(slide.id, { handle, revision: revisionRef.current });
                attach(el2, handle);
              });
            }
          } else {
            cancelRender?.();
            cancelRender = null;
            detach(element);
          }
        },
        { rootMargin: IO_ROOT_MARGIN },
      );

      io.observe(el);

      return () => {
        io.disconnect();
        cancelRender?.();
        cancelRender = null;
        const element = itemPreviewRef.current;
        if (element) detach(element);
      };
    }, [presentation, slide, mediaUrlCache, slideHandleCache, scheduleRender]);

    // Re-render the miniature in place when an edit bumps the revision,
    // without tearing down the IntersectionObserver (the detach→re-attach
    // gap was the cause of the thumbnail flash on every edit). Only fires
    // when the slide is currently visible/attached; off-screen slides are
    // re-rendered on demand by the IO callback using revisionRef.
    React.useEffect(() => {
      const element = itemPreviewRef.current;
      if (!element || !presentation || !slide) return;
      if (!slideHandleRef.current) return; // not visible — IO handles it

      const oldHandle = slideHandleRef.current;
      oldHandle.element.remove();
      oldHandle.dispose();
      slideHandleRef.current = null;
      slideHandleCache.delete(slide.id);

      const handle = renderSlide(presentation, slide, { mediaUrlCache });
      handle.element.style.transformOrigin = "top left";
      const currentScale = widthRef.current > 0 ? widthRef.current / pWidthRef.current : 0;
      if (currentScale > 0) handle.element.style.transform = `scale(${currentScale})`;
      element.appendChild(handle.element);
      slideHandleRef.current = handle;
      slideHandleCache.set(slide.id, { handle, revision });
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
              aspectRatio: `${pWidth} / ${pHeight}`,
              overflow: "hidden",
              pointerEvents: "none",
              contain: "layout style paint",
              ...TYPOGRAPHY_RESET_STYLE,
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
            "aria-hidden": "true",
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
