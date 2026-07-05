/**
 * Roving focus implementation adapted from Radix UI's `@radix-ui/react-roving-focus`.
 *
 * The keyboard-navigation model (tab-stop management, focus-intent mapping,
 * and click-vs-keyboard distinction) is directly inspired by that package,
 * reworked to be index-based so it stays correct under virtualization
 * (where most item DOM is unmounted at any given time).
 *
 * @see https://github.com/radix-ui/primitives/tree/main/packages/react/roving-focus
 */

import * as React from "react";

import type { SlideData, SlideHandle } from "@diceui/pptx-parser";
import { materializeSlideNodes, renderSlide } from "@diceui/pptx-parser";

import { usePresentation, usePresentationStore } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";

const THUMBNAIL_LIST_NAME = "PresentationThumbnailList";
const THUMBNAIL_ITEM_NAME = "PresentationThumbnailItem";
const THUMBNAIL_ITEM_PREVIEW_NAME = "PresentationThumbnailItemPreview";
const THUMBNAIL_ITEM_NUMBER_NAME = "PresentationThumbnailItemNumber";

/**
 * Items mounted above/below the visible range so scrolling never reveals an
 * unmounted row before React can commit the next window.
 */
const DEFAULT_OVERSCAN = 5;

/**
 * Number of items mounted before the first layout measurement resolves the
 * real item stride (and in environments without layout, e.g. jsdom).
 */
const FALLBACK_WINDOW_SIZE = 20;

const VISUALLY_HIDDEN_STYLE: React.CSSProperties = {
  position: "absolute",
  width: "1px",
  height: "1px",
  padding: "0",
  margin: "-1px",
  overflow: "hidden",
  clipPath: "inset(50%)",
  whiteSpace: "nowrap",
  borderWidth: 0,
};

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

/**
 * Nearest scrollable ancestor (`overflow-y: auto | scroll | overlay`), or
 * `null` when the element scrolls with the window.
 */
function findScrollContainer(element: Element): HTMLElement | null {
  let ancestor = element.parentElement;
  while (ancestor) {
    const { overflowY } = getComputedStyle(ancestor);
    if (overflowY === "auto" || overflowY === "scroll" || overflowY === "overlay") {
      return ancestor;
    }
    ancestor = ancestor.parentElement;
  }
  return null;
}

/**
 * The element whose scroll position drives the virtualization window: the
 * list itself when it scrolls, otherwise the nearest scrollable ancestor,
 * otherwise `null` (the window scrolls).
 */
function resolveScroller(listEl: HTMLElement): HTMLElement | null {
  const { overflowY } = getComputedStyle(listEl);
  if (overflowY === "auto" || overflowY === "scroll" || overflowY === "overlay") {
    return listEl;
  }
  return findScrollContainer(listEl);
}

/**
 * Dev-only render-queue instrumentation.
 *
 * Each drain frame reports how many slide renders ran, how long the frame's
 * render work took, and the remaining backlog. Aggregates accumulate on
 * `window.__pptxThumbnailPerf` and a `pptx:thumbnail-perf` CustomEvent is
 * dispatched so a debug UI (e.g. the docs playground) can display them live.
 * Compiled out of production bundles via the NODE_ENV guard.
 */
interface ThumbnailPerfAggregate {
  frames: number;
  renders: number;
  totalMs: number;
  maxFrameMs: number;
  backlog: number;
}

function recordThumbnailPerf(renderCount: number, frameMs: number, backlog: number) {
  if (process.env.NODE_ENV === "production") return;
  if (typeof window === "undefined") return;
  const target = window as unknown as { __pptxThumbnailPerf?: ThumbnailPerfAggregate };
  const agg = (target.__pptxThumbnailPerf ??= {
    frames: 0,
    renders: 0,
    totalMs: 0,
    maxFrameMs: 0,
    backlog: 0,
  });
  agg.frames += 1;
  agg.renders += renderCount;
  agg.totalMs += frameMs;
  agg.maxFrameMs = Math.max(agg.maxFrameMs, frameMs);
  agg.backlog = backlog;
  window.dispatchEvent(new CustomEvent("pptx:thumbnail-perf", { detail: { ...agg } }));
}

interface ThumbnailRovingContextValue {
  /**
   * Tab-stop tracking lives in a mini external store instead of context
   * state: putting the roving id in the context value would invalidate
   * every item on each focus/slide change, re-rendering all mounted
   * thumbnails when only two (old and new tab stop) actually changed.
   */
  subscribeTabStop: (callback: () => void) => () => void;
  getEffectiveTabStopId: () => string | null;
  onItemFocus: (slideId: string) => void;
  onItemRegister: (slideId: string, el: HTMLButtonElement) => void;
  onItemUnregister: (slideId: string) => void;
  /**
   * Index-based keyboard navigation. Resolves the target slide from the
   * store's ordering, scrolls it into the virtualization window if its
   * button is currently unmounted, and focuses it once available.
   */
  focusByIntent: (slideId: string, intent: FocusIntent) => void;
  /** Shared object-URL cache so each image is decoded once across all previews. */
  mediaUrlCache: Map<string, string>;
  /**
   * Rendered slide DOM cache, keyed by slide id. Previews unmounted by the
   * virtualization window keep their handle here; scrolling back re-attaches
   * the existing element instead of re-parsing and re-rendering the slide.
   */
  slideHandleCache: Map<string, SlideHandle>;
  /**
   * Register with the list-level shared ResizeObserver.
   * Returns a cleanup function that unregisters the element.
   */
  observeResize: (el: Element, cb: (width: number) => void) => () => void;
  /**
   * Enqueue a slide render, drained FIFO per animation frame within a time
   * budget (at least one per frame). With virtualization only near-viewport
   * items are ever queued, so FIFO order matches visual order.
   * Returns a cancel function that removes the entry from the queue.
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
  /**
   * The slides currently mounted by the virtualization window, in order.
   * Only slides near the visible scrollport are included; the list manages
   * scrollbar geometry with spacer elements so mapping this slice to
   * `ThumbnailItem`s produces a correctly sized, scrollable rail.
   */
  slides: SlideData[];

  /** Stable id of the currently active slide, or `null` before load. */
  activeSlideId: string | null;

  /** 0-based position of the active slide within ALL slides. */
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
   * - Absent → default `ThumbnailItem` list (virtualized)
   * - ReactNode → rendered as-is inside the container (NOT virtualized)
   * - Function → called with the virtualized window of slides
   */
  children?: React.ReactNode | ((state: ThumbnailListRenderState) => React.ReactNode);

  /**
   * When `true`, keyboard navigation wraps from the last item back to the
   * first (and vice versa).
   *
   * @default false
   */
  loop?: boolean;

  /**
   * Number of extra items to keep mounted above and below the visible
   * scrollport so scrolling never reveals an unmounted row.
   *
   * @default 5
   */
  overscan?: number;
}

interface VirtualRange {
  start: number;
  end: number;
}

interface VirtualMetrics {
  /** Item height + row gap. `0` until the first item has been measured. */
  stride: number;
  /** Row gap of the list container. */
  gap: number;
}

/**
 * Scrollable `listbox` container listing all slide thumbnails.
 * Renders nothing until the presentation is `"ready"`.
 *
 * VIRTUALIZED: only slides near the visible scrollport are mounted (plus an
 * `overscan` buffer on each side); spacer elements keep the scrollbar sized
 * as if every slide were rendered. Item height is measured from the first
 * mounted item, so all items are assumed to share one height (true for the
 * default width + aspect-ratio sizing).
 *
 * Handles keyboard navigation (↑↓ / Home / End) with roving focus so only
 * the active item lives in the tab order at any time. Navigation is
 * index-based: targets outside the window are scrolled into it, mounted,
 * then focused.
 */
export const ThumbnailList = React.forwardRef<HTMLDivElement, ThumbnailListProps>(
  function ThumbnailList(
    { render, children, loop = false, overscan = DEFAULT_OVERSCAN, ...thumbnailListProps },
    forwardedRef,
  ) {
    const { presentation, status } = usePresentation();
    const store = usePresentationStore(THUMBNAIL_LIST_NAME);

    const listRef = React.useRef<HTMLDivElement>(null);
    const itemsRef = React.useRef<Map<string, HTMLButtonElement>>(new Map());
    const isClickFocusRef = React.useRef(false);

    // Props read from stable callbacks without invalidating them.
    const loopRef = React.useRef(loop);
    loopRef.current = loop;
    const overscanRef = React.useRef(overscan);
    overscanRef.current = overscan;

    // --- Roving tab stop (mini external store) ---
    // A ref + listener set, NOT React state: state would flow through the
    // context value and re-render every memoized item on each focus change,
    // when only two items (the old and new tab stop) actually need to
    // update. Falls back to the store's active slide before any explicit
    // focus, so subscribers listen to both sources.
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
    // One object-URL cache and one rendered-slide-handle cache for all
    // previews in this list. A new presentation always starts fresh; the
    // effect below disposes the previous presentation's resources.
    const mediaCacheRef = React.useRef<{ key: object; cache: Map<string, string> } | null>(null);
    if (!mediaCacheRef.current || mediaCacheRef.current.key !== presentation) {
      mediaCacheRef.current = { key: presentation ?? {}, cache: new Map() };
    }
    const mediaUrlCache = mediaCacheRef.current.cache;

    const handleCacheRef = React.useRef<{
      key: object;
      cache: Map<string, SlideHandle>;
    } | null>(null);
    if (!handleCacheRef.current || handleCacheRef.current.key !== presentation) {
      handleCacheRef.current = { key: presentation ?? {}, cache: new Map() };
    }
    const slideHandleCache = handleCacheRef.current.cache;

    // Dispose the caches created for a presentation when it is replaced or
    // the list unmounts. The cleanup closes over the maps belonging to the
    // presentation that was current when the effect ran.
    React.useEffect(() => {
      const handles = slideHandleCache;
      const media = mediaUrlCache;
      return () => {
        for (const handle of handles.values()) handle.dispose();
        handles.clear();
        for (const url of media.values()) URL.revokeObjectURL(url);
        media.clear();
      };
    }, [slideHandleCache, mediaUrlCache]);

    // --- Shared ResizeObserver ---
    // One ResizeObserver serves every preview in the list; per-item
    // observers add meaningful setup and bookkeeping overhead.
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
    // All renderSlide() work is serialised here and drained FIFO within a
    // fixed per-frame time budget. With virtualization the queue only ever
    // holds near-viewport items (mount order == visual order), so no
    // visibility partitioning or prioritisation is needed.
    const renderQueueRef = React.useRef<Array<() => void>>([]);
    const drainRafRef = React.useRef<number | null>(null);

    const drainRenderQueue = React.useCallback(function drain() {
      drainRafRef.current = null;
      const queue = renderQueueRef.current;
      if (queue.length === 0) return;

      const frameStart = performance.now();
      let rendered = 0;

      // ~8ms budget: leaves headroom in a 16.7ms frame for style/layout/
      // paint of the slides just built. Simple slides take 1-3ms, so
      // several render per frame; complex ones degrade to ~1 per frame.
      // Always renders at least one entry so progress is guaranteed.
      const budgetMs = 8;
      do {
        queue.shift()?.();
        rendered += 1;
      } while (queue.length > 0 && performance.now() - frameStart < budgetMs);

      recordThumbnailPerf(rendered, performance.now() - frameStart, queue.length);

      if (queue.length > 0) {
        drainRafRef.current = requestAnimationFrame(drain);
      }
    }, []);

    const scheduleRender = React.useCallback(
      (fn: () => void): (() => void) => {
        renderQueueRef.current.push(fn);
        if (drainRafRef.current === null) {
          drainRafRef.current = requestAnimationFrame(drainRenderQueue);
        }
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

    // --- Virtualization window ---
    // Range of slide indices currently mounted. Starts with a fallback
    // window until the first item's height has been measured; from then on
    // it is derived from the scroll position.
    const [range, setRange] = React.useState<VirtualRange>({
      start: 0,
      end: FALLBACK_WINDOW_SIZE - 1,
    });
    const rangeRef = React.useRef(range);
    rangeRef.current = range;

    const [metrics, setMetrics] = React.useState<VirtualMetrics>({ stride: 0, gap: 0 });
    const metricsRef = React.useRef(metrics);
    metricsRef.current = metrics;

    // Top of item 0 in the scroller's scroll coordinates (accounts for list
    // padding and any content above the list inside the scroller).
    const listTopRef = React.useRef(0);
    // The resolved scroll driver; `null` means the window scrolls.
    const scrollerRef = React.useRef<HTMLElement | null>(null);

    // Reset the window when a new presentation loads (render-phase state
    // adjustment, guarded by identity).
    const lastPresentationRef = React.useRef<object | null>(null);
    if (lastPresentationRef.current !== presentation) {
      lastPresentationRef.current = presentation;
      setRange({ start: 0, end: FALLBACK_WINDOW_SIZE - 1 });
      setMetrics({ stride: 0, gap: 0 });
    }

    const computeRange = React.useCallback(() => {
      const { stride } = metricsRef.current;
      if (stride <= 0) return;
      const total = store.getState().presentation?.slides.length ?? 0;
      if (total === 0) return;

      const scroller = scrollerRef.current;
      const scrollTop = scroller ? scroller.scrollTop : window.scrollY;
      const viewportHeight = scroller ? scroller.clientHeight : window.innerHeight;
      const relativeTop = scrollTop - listTopRef.current;

      const visibleStart = Math.floor(Math.max(relativeTop, 0) / stride);
      const visibleEnd = Math.floor(Math.max(relativeTop + viewportHeight, 0) / stride);
      const start = Math.max(0, Math.min(visibleStart - overscanRef.current, total - 1));
      const end = Math.max(0, Math.min(visibleEnd + overscanRef.current, total - 1));

      setRange((prev) => (prev.start === start && prev.end === end ? prev : { start, end }));
    }, [store]);

    /**
     * Measures the item stride from the first mounted item and the list's
     * row gap, and anchors item 0's position in scroller coordinates.
     * Returns `false` when nothing is measurable yet (no item, or no layout
     * as in jsdom) — the fallback window stays in effect.
     */
    const measure = React.useCallback((): boolean => {
      const listEl = listRef.current;
      if (!listEl) return false;
      const firstItem = listEl.querySelector<HTMLElement>("[data-slide-id]");
      if (!firstItem || firstItem.offsetHeight === 0) return false;

      const gap = Number.parseFloat(getComputedStyle(listEl).rowGap) || 0;
      const stride = firstItem.offsetHeight + gap;

      const scroller = scrollerRef.current;
      const itemRect = firstItem.getBoundingClientRect();
      const itemTopInScroller = scroller
        ? itemRect.top - scroller.getBoundingClientRect().top + scroller.scrollTop
        : itemRect.top + window.scrollY;
      // The first mounted item is `range.start`, not necessarily item 0.
      listTopRef.current = itemTopInScroller - rangeRef.current.start * stride;

      setMetrics((prev) => (prev.stride === stride && prev.gap === gap ? prev : { stride, gap }));
      return true;
    }, []);

    // Wire up the scroll driver: resolve the scroller, listen for scrolls,
    // and re-measure when the list resizes (a width change alters item
    // heights through the aspect-ratio sizing).
    React.useLayoutEffect(() => {
      const listEl = listRef.current;
      if (!listEl || !presentation) return;

      scrollerRef.current = resolveScroller(listEl);
      const target: EventTarget = scrollerRef.current ?? window;

      const measureAndCompute = () => {
        if (measure()) computeRange();
      };

      const onScroll = () => computeRange();
      target.addEventListener("scroll", onScroll, { passive: true });

      measureAndCompute();

      let resizeObserver: ResizeObserver | null = null;
      if (typeof ResizeObserver !== "undefined") {
        resizeObserver = new ResizeObserver(measureAndCompute);
        resizeObserver.observe(listEl);
      }

      return () => {
        target.removeEventListener("scroll", onScroll);
        resizeObserver?.disconnect();
        scrollerRef.current = null;
      };
    }, [presentation, measure, computeRange]);

    /** Instantly scrolls so the item at `index` is fully visible. */
    const scrollToIndex = React.useCallback(
      (index: number) => {
        const { stride, gap } = metricsRef.current;
        if (stride <= 0) return;
        const scroller = scrollerRef.current;

        const itemTop = listTopRef.current + index * stride;
        const itemBottom = itemTop + stride - gap;
        const viewTop = scroller ? scroller.scrollTop : window.scrollY;
        const viewportHeight = scroller ? scroller.clientHeight : window.innerHeight;

        let nextTop: number | null = null;
        if (itemTop < viewTop) nextTop = itemTop;
        else if (itemBottom > viewTop + viewportHeight) nextTop = itemBottom - viewportHeight;
        if (nextTop === null) return;

        if (scroller) scroller.scrollTo({ top: nextTop });
        else window.scrollTo({ top: nextTop });
        // Update the window synchronously so the target mounts this commit
        // instead of waiting for the async scroll event.
        computeRange();
      },
      [computeRange],
    );

    // Focus target for keyboard navigation whose button wasn't mounted at
    // intent time; fulfilled by onItemRegister when the button appears.
    const pendingFocusIdRef = React.useRef<string | null>(null);

    const focusByIntent = React.useCallback(
      (slideId: string, intent: FocusIntent) => {
        const slides = store.getState().presentation?.slides;
        if (!slides || slides.length === 0) return;
        const total = slides.length;
        const index = store.getSlideIndex(slideId);
        if (index < 0) return;

        let target: number;
        if (intent === "first") target = 0;
        else if (intent === "last") target = total - 1;
        else {
          target = intent === "next" ? index + 1 : index - 1;
          if (loopRef.current) target = (target + total) % total;
          else if (target < 0 || target >= total) return;
        }

        const targetSlide = slides[target];
        if (!targetSlide) return;
        const button = itemsRef.current.get(targetSlide.id);
        if (button) {
          // Deferred so the browser finishes processing this keydown (and
          // any focus/scroll side effects) before focus moves; synchronous
          // focus here can be swallowed. Same technique as Radix UI's
          // roving-focus implementation.
          setTimeout(() => button.focus());
        } else {
          pendingFocusIdRef.current = targetSlide.id;
          scrollToIndex(target);
        }
      },
      [store, scrollToIndex],
    );

    const activeSlideId = React.useSyncExternalStore(
      store.subscribe,
      () => store.getState().activeSlideId,
      () => null,
    );

    // Auto-focus the active (or first) thumbnail ONCE per presentation load.
    // This lets arrow-key navigation work immediately without requiring the
    // user to Tab into the list first. Guarded by presentation identity:
    // re-running on every activeSlideId change would steal focus from
    // whatever the user is interacting with (e.g. a next-slide button in the
    // main view) each time the slide changes.
    const autoFocusedPresentationRef = React.useRef<object | null>(null);
    React.useEffect(() => {
      if (!presentation || autoFocusedPresentationRef.current === presentation) return;
      autoFocusedPresentationRef.current = presentation;
      const items = itemsRef.current;
      const activeItem = activeSlideId ? items.get(activeSlideId) : undefined;
      const firstItem = activeItem ?? items.values().next().value;
      firstItem?.focus({ preventScroll: true });
    }, [presentation, activeSlideId]);

    // Container tabIndex only cares whether ANY button owns the tab stop,
    // so subscribe to the boolean — the container re-renders on the
    // null↔non-null transition, not on every focus move between items.
    const hasTabStop = React.useSyncExternalStore(
      subscribeTabStop,
      () => getEffectiveTabStopId() != null,
      () => false,
    );

    const onItemRegister = React.useCallback((slideId: string, el: HTMLButtonElement) => {
      itemsRef.current.set(slideId, el);
      if (pendingFocusIdRef.current === slideId) {
        pendingFocusIdRef.current = null;
        // Deferred for the same keydown-swallowing reason as focusByIntent.
        setTimeout(() => el.focus());
      }
    }, []);

    const onItemUnregister = React.useCallback((slideId: string) => {
      itemsRef.current.delete(slideId);
    }, []);

    const rovingContextValue = React.useMemo<ThumbnailRovingContextValue>(
      () => ({
        subscribeTabStop,
        getEffectiveTabStopId,
        onItemFocus: setCurrentTabStopId,
        onItemRegister,
        onItemUnregister,
        focusByIntent,
        mediaUrlCache,
        slideHandleCache,
        observeResize,
        scheduleRender,
      }),
      // Every dependency is stable per-presentation (the caches are guarded
      // by identity refs above; the callbacks are stable useCallbacks), so
      // this context value never invalidates the memoized items on focus or
      // slide changes.
      [
        subscribeTabStop,
        getEffectiveTabStopId,
        setCurrentTabStopId,
        onItemRegister,
        onItemUnregister,
        focusByIntent,
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

    // Clamp the window to the current deck and slice the mounted slides.
    const start = Math.max(0, Math.min(range.start, total - 1));
    const end = Math.max(start, Math.min(range.end, total - 1));
    const windowSlides = presentation.slides.slice(start, end + 1);

    // Spacers replace the unmounted items above/below the window so the
    // scrollbar geometry matches a fully rendered list. As flex children
    // they also participate in the container's row gap, hence the `- gap`
    // (a spacer contributes its height PLUS one gap).
    const { stride, gap } = metrics;
    const topSpacerHeight = stride > 0 && start > 0 ? start * stride - gap : 0;
    const bottomSpacerHeight = stride > 0 && end < total - 1 ? (total - 1 - end) * stride - gap : 0;

    let resolvedChildren: React.ReactNode;
    if (typeof children === "function") {
      resolvedChildren = children({
        slides: windowSlides,
        activeSlideId,
        activeIndex,
        goTo: (id) => store.goTo(id),
        goToIndex: (i) => store.goToIndex(i),
      });
    } else if (children != null) {
      resolvedChildren = children;
    } else {
      resolvedChildren = windowSlides.map((slide) => (
        <ThumbnailItem key={slide.id} slideId={slide.id}>
          <ThumbnailItemPreview />
          <ThumbnailItemNumber style={VISUALLY_HIDDEN_STYLE} />
        </ThumbnailItem>
      ));
    }

    // Static ReactNode children opt out of virtualization (no spacers).
    const isVirtualized = typeof children === "function" || children == null;
    const virtualizedChildren = isVirtualized ? (
      <>
        {topSpacerHeight > 0 && (
          <div aria-hidden style={{ height: topSpacerHeight, flexShrink: 0 }} />
        )}
        {resolvedChildren}
        {bottomSpacerHeight > 0 && (
          <div aria-hidden style={{ height: bottomSpacerHeight, flexShrink: 0 }} />
        )}
      </>
    ) : (
      resolvedChildren
    );

    return (
      <ThumbnailRovingContext.Provider value={rovingContextValue}>
        {renderElement(
          "div",
          { render },
          {
            state,
            ref: [listRef, forwardedRef],
            props: [
              {
                role: "listbox",
                "aria-label": "Slide thumbnails",
                "aria-orientation": "vertical",
                // When a button owns tabIndex=0 the container steps out of the tab
                // order: the list has exactly ONE external tab stop (the active
                // button). Shift+Tab from the button then skips the container and
                // exits the list in a single key press.
                // When no button has a tab stop yet (e.g. before auto-focus fires),
                // the container acts as the entry point and redirects focus.
                tabIndex: hasTabStop ? -1 : 0,
                onMouseDown: (event: React.MouseEvent<HTMLDivElement>) => {
                  if (event.target === event.currentTarget) isClickFocusRef.current = true;
                },
                onFocus: (event: React.FocusEvent<HTMLDivElement>) => {
                  // Container only receives keyboard focus when no button owns
                  // tabIndex=0 yet. Redirect to the first mounted button in
                  // slide order (registration order is mount order, which
                  // under virtualization is not slide order).
                  if (event.target !== event.currentTarget) return;
                  if (isClickFocusRef.current) {
                    isClickFocusRef.current = false;
                    return;
                  }
                  const candidates = Array.from(itemsRef.current.entries())
                    .sort((a, b) => store.getSlideIndex(a[0]) - store.getSlideIndex(b[0]))
                    .map(([, el]) => el);
                  focusFirst(candidates, true);
                },
                children: virtualizedChildren,
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
  /**
   * Stable identifier for the slide this item represents (`SlideData.id`).
   * Correct across reorders, insertions, and deletions.
   */
  slideId: string;

  /**
   * Replace the item button element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<ThumbnailItemState>;
}

/**
 * Clickable `option` button for a single slide in a `ThumbnailList`.
 *
 * - Identity is derived from `slide.id`: stable across list mutations.
 * - Auto-wires `onClick → goTo`, `data-active`, `aria-selected`, and roving
 *   `tabIndex` (0 when active, -1 otherwise).
 * - Exposes `aria-posinset`/`aria-setsize` so assistive tech announces the
 *   full deck size even though the list is virtualized.
 * - Provides context so nested `ThumbnailItemPreview` and `ThumbnailItemNumber`
 *   need no explicit props.
 * - Defaults to rendering `<ThumbnailItemPreview />` + `<ThumbnailItemNumber />`
 *   when no children are provided.
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

    // O(1) map lookup — the store maintains a slideIndexById map that is
    // rebuilt on load; no linear scan per emit.
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

    // Register this button in the roving context's map so the list can
    // focus targets without querySelectorAll. The callback ref fires when
    // the DOM element is attached/detached.
    const { onItemRegister, onItemUnregister } = rovingContext;
    const registerRef = React.useCallback(
      (element: HTMLButtonElement | null) => {
        if (element) {
          onItemRegister(slideId, element);
        } else {
          onItemUnregister(slideId);
        }
      },
      [slideId, onItemRegister, onItemUnregister],
    );

    // This item owns the single tab stop inside the list when it matches the
    // roving tab-stop store: all other items get tabIndex=-1. Subscribing to
    // the boolean slice means a focus move re-renders exactly two items (the
    // old and new tab stop) instead of the whole list.
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
                style: {
                  width: "100%",
                },
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

                  rovingContext.focusByIntent(slideId, focusIntent);
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

  /**
   * Css scale factor applied to the slide element
   * (container width / presentation width). `0` before the container is
   * measured. Only tracked reactively when a `render` prop is provided;
   * default usage applies the scale imperatively without re-rendering.
   */
  scale: number;
}

export interface ThumbnailItemPreviewProps extends React.ComponentProps<"div"> {
  /**
   * Replace the preview element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   *
   * The rendered element is the clipping container: the parsed slide DOM is
   * appended to it imperatively. Preserve `overflow: hidden` and dimensions.
   */
  render?: RenderProp<ThumbnailItemPreviewState>;
}

/**
 * Renders the slide miniature for the enclosing `ThumbnailItem`.
 *
 * Must be a descendant of `<Presentation.ThumbnailItem>`: reads slide ID
 * from context. Width is measured automatically via `ResizeObserver` so no
 * sizing props are required.
 *
 * The rendered element is the clipping container: parsed slide DOM is appended
 * to it and css-scaled to fit. Marked `aria-hidden` since the enclosing
 * button's `aria-label` already identifies the slide.
 *
 * The slide DOM itself is built once per slide and kept in the list's
 * handle cache: when the virtualization window unmounts this preview the
 * element is detached (not disposed), and re-attached instantly if the
 * slide scrolls back into view.
 */
export const ThumbnailItemPreview = React.forwardRef<HTMLDivElement, ThumbnailItemPreviewProps>(
  function ThumbnailItemPreview({ render, ...thumbnailItemPreviewProps }, forwardedRef) {
    const itemContext = useThumbnailItemContext(THUMBNAIL_ITEM_PREVIEW_NAME);
    const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_PREVIEW_NAME);
    const { presentation } = usePresentation();
    const store = usePresentationStore(THUMBNAIL_ITEM_PREVIEW_NAME);

    const itemPreviewRef = React.useRef<HTMLDivElement>(null);
    const slideHandleRef = React.useRef<SlideHandle | null>(null);
    // Shared list-level infrastructure: one media cache, one rendered-slide
    // cache, one ResizeObserver, one render queue.
    const { mediaUrlCache, slideHandleCache, observeResize, scheduleRender } = rovingContext;
    // Tracks whether the slide DOM has actually been appended, used imperatively
    // to remove data-pending without triggering a React re-render per slide.
    const hasRenderedRef = React.useRef(false);

    // IMPERATIVE SIZING MODEL — width measurements never re-render by default.
    //
    // The measured container width lives in `widthRef`, and resizes apply the
    // scale transform directly to the slide element. React state
    // (`containerWidth`) is only kept in sync when a `render` prop is
    // provided, because that is the sole consumer of the reactive `scale`
    // value. Default usage therefore pays zero re-renders on mount
    // measurement and on every subsequent rail/window resize.
    const widthRef = React.useRef(0);
    const [containerWidth, setContainerWidth] = React.useState(0);
    const hasRenderPropRef = React.useRef(false);
    hasRenderPropRef.current = render != null;

    // O(1) via the store's id→index map; avoids an O(N) slides.find() on
    // every render of every preview.
    const slideIndex = store.getSlideIndex(itemContext.slideId);
    const slide = presentation?.slides[slideIndex] ?? null;
    const pWidth = presentation?.width ?? 1;
    const pHeight = presentation?.height ?? 1;
    const scale = containerWidth > 0 ? containerWidth / pWidth : 0;

    // Ref so the resize callback can read the current presentation width
    // without re-subscribing to the shared ResizeObserver.
    const pWidthRef = React.useRef(pWidth);
    pWidthRef.current = pWidth;

    // Measure container width synchronously before first paint so the queued
    // slide render picks up the correct transform immediately.
    React.useLayoutEffect(() => {
      const el = itemPreviewRef.current;
      if (el && el.offsetWidth > 0) {
        widthRef.current = el.offsetWidth;
        if (hasRenderPropRef.current) setContainerWidth(el.offsetWidth);
      }
    }, []);

    // Wire the shared ResizeObserver. Applies the transform imperatively:
    // a resize touches only style on already-rendered slide elements, it
    // never re-renders the React tree or re-runs the slide render effect.
    React.useEffect(() => {
      const el = itemPreviewRef.current;
      if (!el) return;
      return observeResize(el, (width) => {
        widthRef.current = width;
        const nextScale = width > 0 ? width / pWidthRef.current : 0;
        if (slideHandleRef.current && nextScale > 0) {
          slideHandleRef.current.element.style.transform = `scale(${nextScale})`;
        }
        if (hasRenderPropRef.current) setContainerWidth(width);
      });
    }, [observeResize]);

    // RENDER-ONCE, CACHE-FOREVER MODEL (per presentation).
    //
    // The virtualization window mounts/unmounts this preview as the user
    // scrolls. On mount: a cached handle re-attaches synchronously (layout
    // effect, so no pending flash when scrolling back); a cache miss
    // schedules the expensive parse+render on the list's budgeted queue.
    // On unmount: the slide element is detached but NOT disposed — the
    // handle stays in the list-level cache for instant re-attachment. The
    // list disposes the whole cache when the presentation changes or the
    // list unmounts.
    React.useLayoutEffect(() => {
      const el = itemPreviewRef.current;
      if (!el || !presentation || !slide) return;

      const attach = (element: HTMLDivElement, handle: SlideHandle) => {
        // Read the latest measured width at attach time (a queued render
        // may land many frames after it was scheduled).
        const currentScale = widthRef.current > 0 ? widthRef.current / pWidthRef.current : 0;
        if (currentScale > 0) {
          handle.element.style.transform = `scale(${currentScale})`;
        }
        element.appendChild(handle.element);
        slideHandleRef.current = handle;
        // Flip data-pending imperatively: avoids triggering N React
        // re-renders while the queue drains.
        hasRenderedRef.current = true;
        delete element.dataset.pending;
      };

      let cancel: (() => void) | null = null;
      const cached = slideHandleCache.get(slide.id);
      if (cached) {
        attach(el, cached);
      } else {
        cancel = scheduleRender(() => {
          const element = itemPreviewRef.current;
          // Guard against a second render being queued before the first
          // cleanup runs (e.g. Strict Mode remount, effect identity change).
          if (!element || slideHandleRef.current) return;

          if (!slide.nodesMaterialized) materializeSlideNodes(presentation, slide);
          const handle = renderSlide(presentation, slide, { mediaUrlCache });
          handle.element.style.transformOrigin = "top left";
          slideHandleCache.set(slide.id, handle);
          attach(element, handle);
        });
      }

      return () => {
        cancel?.();
        const handle = slideHandleRef.current;
        slideHandleRef.current = null;
        // Detach only — the handle lives on in the cache.
        handle?.element.remove();
        const element = itemPreviewRef.current;
        if (element) {
          hasRenderedRef.current = false;
          element.dataset.pending = "";
        }
      };
    }, [presentation, slide, mediaUrlCache, slideHandleCache, scheduleRender]);

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
            // Present while the slide DOM hasn't landed; removed imperatively
            // once the queued render completes. Style with [data-pending].
            "data-pending": hasRenderedRef.current ? undefined : "",
            // Prevent Tab from entering focusable PPTX content (links, forms, etc.)
            inert: true,
            style: {
              width: "100%",
              aspectRatio: `${pWidth} / ${pHeight}`,
              overflow: "hidden",
              pointerEvents: "none",
              // `layout style paint` gives layout isolation (mutations
              // inside don't reflow the page) and paint isolation WITHOUT
              // `size` containment, which would suppress the
              // `aspect-ratio`-derived height and collapse the box.
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
  /**
   *  Replace the number span element.
   * - ReactElement: cloned with composed props
   * - Function: (props, state) => ReactElement
   */
  render?: RenderProp<{ isActive: boolean; displayIndex: number; slideId: string }>;
}

/**
 * Renders the 1-based slide number for the enclosing `ThumbnailItem`.
 *
 * Must be a descendant of `<Presentation.ThumbnailItem>`: reads the display
 * number from context. Marked `aria-hidden` since the enclosing button's
 * `aria-label` already announces the slide number.
 *
 * Completely unstyled: add `className` / `style` for visual treatment.
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
