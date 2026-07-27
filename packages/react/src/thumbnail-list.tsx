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
import { applySlideScale, materializeSlide, renderSlide } from "@diceui/pptx-core";

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
 *
 * Runway for the previews the background pass has not reached yet: observer
 * notifications are delivered after the frame is painted, so a preview that
 * first hears about itself when it is already on screen has necessarily been
 * painted as a skeleton at least once. ~2000px is a dozen thumbnails, which at
 * a hard flick buys several frames of lead time.
 */
const INTERSECTION_OBSERVER_ROOT_MARGIN = "2000px 0px";

/**
 * How many rendered miniatures the list keeps alive at once.
 *
 * Sized to hold an ordinary deck whole, because the background pass renders
 * ahead of the scroll and a cap below the deck length would spend that work
 * only to throw it away. Each miniature is on the order of a hundred nodes, so
 * a full window is tens of thousands: `content-visibility` means the browser
 * skips layout and paint for the offscreen ones, but they are not free.
 */
const MAX_RETAINED_PREVIEWS = 200;

/**
 * How many retained miniatures may be slides carrying a chart.
 *
 * A chart is a live ECharts instance backed by its own canvas rather than
 * plain DOM, so these cost far more than the cap above accounts for and are
 * held to a window around the viewport instead. They are also the one thing
 * the background pass skips, for the same reason.
 */
const MAX_RETAINED_CHART_PREVIEWS = 48;

/** Per-frame slice of the main thread spent filling previews on demand. */
const RENDER_BUDGET_MS = 8;

/**
 * First guess at what rendering one slide costs, before the background pass
 * has measured any. Only used to decide whether an idle period has room for
 * another slide, so being wrong costs at most one overrun.
 */
const PRERENDER_COST_GUESS_MS = 2;

/**
 * Floor under that estimate. An idle period this short is not worth taking,
 * and the estimate must not be able to collapse to nothing on a fast machine
 * and start a render with no room left.
 */
const PRERENDER_MIN_SLICE_MS = 2;

/**
 * Allowance for a background slice taken on a frame callback rather than an
 * idle one. Small enough to leave the frame its own budget, since unlike an
 * idle callback this one is not being offered spare time, it is taking it.
 */
const IDLE_FRAME_BUDGET_MS = 4;

/**
 * How long the scroller has to be still before the background pass resumes.
 *
 * A dropped frame mid-scroll is more noticeable than a thumbnail arriving a
 * moment later, and the browser is willing to report idle time during a scroll
 * that is driven by the compositor, so the pass stands down on its own.
 */
const SCROLL_IDLE_MS = 200;

/**
 * Cached rendered thumbnail plus the edit revision it was rendered at.
 */
interface CachedThumbnail {
  slideHandle: SlideHandle;
  revision: number;
}

/**
 * A rendered miniature the list may reclaim when it runs past its cap.
 *
 * Registered both by previews, when they mount a miniature, and by the
 * background pass, for miniatures rendered before any preview asked for one.
 */
interface RetainedPreview {
  /** Position in the deck: retention is a window around what is on screen. */
  index: number;
  hasChart: boolean;
  /** Whether a preview has this miniature mounted right now. */
  isAttached: boolean;
  /** Previews the observer currently reports as on screen are never reclaimed. */
  isVisible: () => boolean;
  /** Detaches, disposes and uncaches the miniature, back to a skeleton. */
  evict: () => void;
}

interface QueuedRender {
  element: HTMLElement;
  render: () => void;
}

/**
 * True when the slide has a chart node.
 *
 * Only meaningful once the slide has been materialized: until then its nodes
 * are still raw XML and every slide reads as chartless. Only top-level nodes
 * are checked, so a grouped chart also reads as chartless, which costs nothing
 * but a place in the wrong retention budget.
 */
export function slideHasChart(slide: SlideData): boolean {
  return slide.nodes.some((node) => node.nodeType === "chart");
}

interface IdleHandle {
  cancel: () => void;
}

/**
 * Runs `task` when the browser has time to spare, handing it the milliseconds
 * it may use.
 *
 * An idle callback is raced against a frame callback, and whichever arrives
 * first wins. `requestIdleCallback` offers by far the better slice, up to 50ms
 * against a few, but it only offers it when the browser has decided it is
 * genuinely idle, which on a page doing anything at all is a few times a
 * second. On its own that is too slow to render a deck ahead of a reader. The
 * frame callback puts a floor under it of one small slice per frame, and is
 * also the whole story where `requestIdleCallback` is missing, which is Safari
 * before 17.4 and most test environments.
 */
function scheduleIdle(task: (timeRemaining: () => number) => void): IdleHandle {
  let done = false;
  const runOnce = (timeRemaining: () => number) => {
    if (done) return;
    done = true;
    cancel();
    task(timeRemaining);
  };

  const frame = requestAnimationFrame(() => {
    // Timed from when the frame callback runs, not from when it was asked for:
    // the wait in between is the browser's, not the task's.
    const startedAt = performance.now();
    runOnce(() => Math.max(0, IDLE_FRAME_BUDGET_MS - (performance.now() - startedAt)));
  });
  const idle =
    typeof requestIdleCallback === "function"
      ? requestIdleCallback((deadline) => runOnce(() => deadline.timeRemaining()))
      : null;

  function cancel() {
    cancelAnimationFrame(frame);
    if (idle !== null) cancelIdleCallback(idle);
  }

  return {
    cancel: () => {
      done = true;
      cancel();
    },
  };
}

/**
 * Nearest scrollable ancestor of `element`, or `null` when that is the document.
 *
 * This is the IntersectionObserver root. `rootMargin` only inflates the root's
 * own rect: the clip rects of scrollers in between are applied as they are, so
 * observing a list that scrolls inside a container against the document gives no
 * runway whatsoever, however large the margin.
 */
function findScrollRoot(element: Element | null): Element | null {
  for (let node = element; node instanceof HTMLElement; node = node.parentElement) {
    if (node === document.body || node === document.documentElement) break;
    const { overflowY } = getComputedStyle(node);
    if (overflowY === "auto" || overflowY === "scroll") return node;
  }
  return null;
}

/**
 * Distance in px between `element` and the visible viewport; `0` while any part
 * of it is on screen.
 *
 * Used to order queued renders. Measuring against the viewport rather than a
 * scroll container keeps this independent of which ancestor actually scrolls.
 */
export function distanceFromViewport(element: HTMLElement): number {
  const viewportHeight = element.ownerDocument.defaultView?.innerHeight ?? 0;
  const rect = element.getBoundingClientRect();
  if (rect.bottom >= 0 && rect.top <= viewportHeight) return 0;
  return rect.top > viewportHeight ? rect.top - viewportHeight : -rect.bottom;
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
   * Rendered slide DOM cache, keyed by slide id. Survives a preview remount
   * (slide reorder) so the element is re-attached instead of rendered again.
   * Entries are invalidated when the slide's edit revision moves past the
   * cached one.
   */
  slideHandleCache: Map<string, CachedThumbnail>;
  /**
   * Register with the list-level shared ResizeObserver.
   * Returns a cleanup function that unregisters the element.
   */
  observeResize: (element: Element, cb: (width: number) => void) => () => void;
  /**
   * Enqueue a `renderSlide()` call for `element`, drained per animation frame
   * within a fixed budget, nearest to the viewport first.
   */
  scheduleRender: (element: HTMLElement, render: () => void) => () => void;
  /**
   * Register a rendered miniature as reclaimable. Returns a release function
   * that takes it back out of the retention set.
   */
  retainPreview: (slideId: string, preview: RetainedPreview) => () => void;
  /**
   * Report whether a preview is inside the observer's runway. This is what the
   * list centres its retention window and its background pass on.
   */
  setPreviewVisible: (index: number, isVisible: boolean) => void;
  /** Scroller the previews live in, used as the IntersectionObserver root. */
  scrollRootRef: React.RefObject<Element | null>;
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
 * All slide buttons are mounted immediately (cheap empty containers). The
 * miniatures inside them are rendered by a background pass that works outward
 * from the viewport during the browser's idle time, so by the time a preview
 * is scrolled to there is usually nothing left to do but attach it. A preview
 * the pass has not reached yet falls back to rendering on demand, through an
 * IntersectionObserver with a generous `rootMargin` and a per-frame budget
 * drained nearest to the viewport first. Rendered miniatures then stay in the
 * DOM, so a given slide can only ever show its placeholder once.
 *
 * Handles keyboard navigation (↑↓ / Home / End) with roving focus so only
 * the active item lives in the tab order at any time.
 */
export const ThumbnailList = React.forwardRef<HTMLDivElement, ThumbnailListProps>(
  function ThumbnailList({ render, children, loop = false, ...thumbnailListProps }, forwardedRef) {
    const { presentation, status } = usePresentation();
    const store = useStoreContext(THUMBNAIL_LIST_NAME);

    const itemsRef = useLazyRef(() => new Map<string, HTMLButtonElement>());
    const listRef = React.useRef<HTMLDivElement | null>(null);
    const scrollRootRef = React.useRef<Element | null>(null);
    const lastScrollAtRef = React.useRef(0);
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

    // Shared caches, keyed by presentation identity
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

    // Shared ResizeObserver used to observe the size of the thumbnail list
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

    // Batch renderSlide() calls that arrive simultaneously (initial viewport
    // fill, or a scroll that outruns the background pass) and drain them within
    // a per-frame budget so no single commit blocks the main thread.
    const renderQueueRef = useLazyRef<QueuedRender[]>(() => []);
    const drainRafRef = React.useRef<number | null>(null);

    const drainRenderQueue = React.useCallback(
      function drain() {
        drainRafRef.current = null;
        const queue = renderQueueRef.current;
        if (queue.length === 0) return;

        // Nearest to the viewport first. During a fast scroll the queue fills
        // with previews the user has already passed, and spending the budget on
        // those is exactly what leaves the ones now on screen as skeletons.
        //
        // The order is computed once per frame: re-measuring after each render
        // would force a layout flush per entry, which costs more than acting on
        // an order that is at most one frame stale.
        const ordered = queue
          .map((entry) => ({ entry, distance: distanceFromViewport(entry.element) }))
          .sort((a, b) => a.distance - b.distance);

        const frameStart = performance.now();
        for (const { entry } of ordered) {
          const index = queue.indexOf(entry);
          if (index !== -1) queue.splice(index, 1);
          entry.render();
          if (performance.now() - frameStart >= RENDER_BUDGET_MS) break;
        }

        if (queue.length > 0) drainRafRef.current = requestAnimationFrame(drain);
      },
      [renderQueueRef],
    );

    const scheduleRender = React.useCallback(
      (element: HTMLElement, render: () => void): (() => void) => {
        const entry: QueuedRender = { element, render };
        renderQueueRef.current.push(entry);
        if (drainRafRef.current === null)
          drainRafRef.current = requestAnimationFrame(drainRenderQueue);
        return () => {
          const index = renderQueueRef.current.indexOf(entry);
          if (index !== -1) renderQueueRef.current.splice(index, 1);
        };
      },
      [drainRenderQueue, renderQueueRef],
    );

    // Indices the observers currently report inside the runway. Both the
    // retention window and the background pass are centred on these.
    const visibleIndicesRef = useLazyRef(() => new Set<number>());
    const wakePrerenderRef = React.useRef<(() => void) | null>(null);

    /**
     * Middle of the runway, or the active slide before anything has been
     * observed. Deliberately not the scroll offset: this has to be meaningful
     * for miniatures that no preview has mounted, which have no geometry.
     */
    const getFocusIndex = React.useCallback(() => {
      const visible = visibleIndicesRef.current;
      if (visible.size === 0) {
        const activeSlideId = store.getState().activeSlideId;
        return activeSlideId ? Math.max(0, store.getSlideIndex(activeSlideId)) : 0;
      }
      let lowest = Number.POSITIVE_INFINITY;
      let highest = Number.NEGATIVE_INFINITY;
      for (const index of visible) {
        if (index < lowest) lowest = index;
        if (index > highest) highest = index;
      }
      return Math.round((lowest + highest) / 2);
    }, [store, visibleIndicesRef]);

    const setPreviewVisible = React.useCallback(
      (index: number, isVisible: boolean) => {
        const visible = visibleIndicesRef.current;
        if (isVisible === visible.has(index)) return;
        if (isVisible) visible.add(index);
        else visible.delete(index);
        // The window the background pass works on has moved.
        wakePrerenderRef.current?.();
      },
      [visibleIndicesRef],
    );

    const retainedPreviewsRef = useLazyRef(() => new Map<string, RetainedPreview>());

    /**
     * Retention is a window around what is on screen, so what gets reclaimed is
     * whatever sits furthest from it in the deck. Deliberately not
     * least-recently-used: the background pass renders outward from the
     * viewport, which would make its earliest and most useful work the oldest,
     * and so the first thing thrown away.
     */
    const retainPreview = React.useCallback(
      (slideId: string, preview: RetainedPreview): (() => void) => {
        const retained = retainedPreviewsRef.current;
        retained.set(slideId, preview);

        const focusIndex = getFocusIndex();
        const trim = (hasChart: boolean, cap: number) => {
          let count = 0;
          const reclaimable: { id: string; entry: RetainedPreview }[] = [];
          for (const [id, entry] of retained) {
            if (entry.hasChart !== hasChart) continue;
            count++;
            if (id !== slideId && !entry.isVisible()) reclaimable.push({ id, entry });
          }
          if (count <= cap) return;

          reclaimable.sort(
            (a, b) => Math.abs(b.entry.index - focusIndex) - Math.abs(a.entry.index - focusIndex),
          );
          for (const { id, entry } of reclaimable.slice(0, count - cap)) {
            retained.delete(id);
            entry.evict();
          }
        };

        trim(true, MAX_RETAINED_CHART_PREVIEWS);
        trim(false, MAX_RETAINED_PREVIEWS);

        return () => {
          if (retained.get(slideId) === preview) retained.delete(slideId);
        };
      },
      [retainedPreviewsRef, getFocusIndex],
    );

    React.useEffect(() => {
      const retainedPreviews = retainedPreviewsRef.current;
      return () => {
        sharedResizeObserverRef.current?.disconnect();
        sharedResizeObserverRef.current = null;
        if (drainRafRef.current !== null) {
          cancelAnimationFrame(drainRafRef.current);
          drainRafRef.current = null;
        }
        renderQueueRef.current = [];
        retainedPreviews.clear();
      };
    }, [sharedResizeObserverRef, renderQueueRef, retainedPreviewsRef]);

    /**
     * Renders the deck ahead of the scroll, in the browser's idle time, working
     * outward from what is on screen.
     *
     * This is what keeps scrolling smooth. Rendering a slide costs several
     * milliseconds, so at any real scrolling speed more previews come into view
     * per frame than can be filled in one, and filling on demand is a race the
     * scroll wins. Idle time is not scarce here: a couple of hundred slides is
     * about a second of work in total, spread over as many idle periods as it
     * takes, and once it is done scrolling has nothing left to pay for.
     *
     * Chart slides are left out. Their cost is a live chart engine rather than
     * DOM, so they get a much smaller retention window, and rendering them
     * ahead would only push each other out of it.
     */
    React.useEffect(() => {
      let idle: IdleHandle | null = null;

      // Read through the store rather than closing over the slide array: a
      // presentation keeps its identity across edits, so a captured array can
      // outlive the deck it described.
      const getSlides = () => store.getState().presentation?.slides ?? [];

      /**
       * Whether the pass should render this slide: either nothing is cached for
       * it, or what is cached predates an edit and no preview is showing it (a
       * preview refreshes its own miniature in place).
       */
      /**
       * Slides the pass has found to carry a chart and will not touch again.
       * Chart slides can only be recognised once materialized, so this is
       * filled in as the pass goes rather than known up front.
       */
      const chartSlideIds = new Set<string>();

      const needsRender = (slide: SlideData): boolean => {
        const cached = slideHandleCache.get(slide.id);
        if (!cached) return true;
        if (cached.revision === store.getSlideRevision(slide.id)) return false;
        return !retainedPreviewsRef.current.get(slide.id)?.isAttached;
      };

      const nextIndexToRender = (): number => {
        const slides = getSlides();
        const focusIndex = getFocusIndex();
        let considered = 0;
        for (let step = 0; step < slides.length; step++) {
          const candidates = step === 0 ? [focusIndex] : [focusIndex - step, focusIndex + step];
          for (const index of candidates) {
            const slide = slides[index];
            if (!slide || chartSlideIds.has(slide.id)) continue;
            if (considered >= MAX_RETAINED_PREVIEWS) return -1;
            considered++;
            if (needsRender(slide)) return index;
          }
        }
        return -1;
      };

      // What one slide has been costing lately. The pass takes whatever idle
      // slice the browser offers, which on a busy page is a few milliseconds,
      // so a fixed threshold either starves it or overruns the period; this
      // measures the deck in front of it instead of guessing.
      let renderCostMs = PRERENDER_COST_GUESS_MS;

      const renderAhead = (index: number) => {
        const deck = store.getState().presentation;
        const slide = deck?.slides[index];
        if (!deck || !slide) return;

        const startedAt = performance.now();

        // Parsing the slide is the only way to find out whether it holds a
        // chart, and it is work the first render would do anyway. Chart slides
        // stop here: they are the on-demand path's to render.
        materializeSlide(deck, slide);
        if (slideHasChart(slide)) {
          chartSlideIds.add(slide.id);
          return;
        }

        slideHandleCache.get(slide.id)?.slideHandle.dispose();
        const slideHandle = renderSlide(deck, slide, { mediaUrlCache });
        slideHandleCache.set(slide.id, {
          slideHandle,
          revision: store.getSlideRevision(slide.id),
        });
        retainPreview(slide.id, {
          index,
          hasChart: false,
          isAttached: false,
          isVisible: () => false,
          evict: () => {
            slideHandleCache.get(slide.id)?.slideHandle.dispose();
            slideHandleCache.delete(slide.id);
          },
        });

        const cost = performance.now() - startedAt;
        renderCostMs = Math.max(PRERENDER_MIN_SLICE_MS, renderCostMs * 0.7 + cost * 0.3);
      };

      const schedule = () => {
        idle ??= scheduleIdle(run);
      };

      const run = (timeRemaining: () => number) => {
        idle = null;
        // Previews the user is waiting on come first, and a render mid-scroll
        // is the jank this pass exists to avoid.
        if (
          renderQueueRef.current.length > 0 ||
          performance.now() - lastScrollAtRef.current < SCROLL_IDLE_MS
        ) {
          schedule();
          return;
        }

        // Always take one slide, then keep going for as long as the slice
        // lasts. Without that first one the pass deadlocks wherever a slide
        // costs more than the slice on offer, which is any browser or machine
        // slower than the budget assumes, and those are exactly the ones that
        // cannot afford to render during the scroll instead.
        let index = nextIndexToRender();
        let isFirst = true;
        while (index !== -1 && (isFirst || timeRemaining() > renderCostMs)) {
          renderAhead(index);
          isFirst = false;
          index = nextIndexToRender();
        }
        // Nothing left to do: the pass stops until the window moves.
        if (index !== -1) schedule();
      };

      wakePrerenderRef.current = schedule;
      // Edits, reorders and slide additions all land here, and a pass with
      // nothing to do costs one walk of the deck.
      const unsubscribe = store.subscribe(schedule);
      schedule();

      return () => {
        wakePrerenderRef.current = null;
        unsubscribe();
        idle?.cancel();
        idle = null;
      };
    }, [
      store,
      mediaUrlCache,
      slideHandleCache,
      getFocusIndex,
      retainPreview,
      renderQueueRef,
      retainedPreviewsRef,
    ]);

    // Resolved before the previews' observers are created: child layout effects
    // run ahead of this one, but every passive effect runs after it.
    React.useLayoutEffect(() => {
      scrollRootRef.current = findScrollRoot(listRef.current);
    }, [status]);

    React.useEffect(() => {
      const scrollTarget: EventTarget = scrollRootRef.current ?? window;
      const onScroll = () => {
        lastScrollAtRef.current = performance.now();
      };
      scrollTarget.addEventListener("scroll", onScroll, { passive: true });
      return () => scrollTarget.removeEventListener("scroll", onScroll);
    }, [status]);

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
        retainPreview,
        setPreviewVisible,
        scrollRootRef,
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
        retainPreview,
        setPreviewVisible,
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
            ref: [listRef, forwardedRef],
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
    const store = useStoreContext(THUMBNAIL_ITEM_NAME);
    const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_NAME);

    const isActive = useStoreSelector(store, (s) => s.activeSlideId === slideId, false);

    const displayIndex = useSlideIndex(store, slideId);

    const setSize = useStoreSelector(store, (s) => s.presentation?.slides.length ?? 0, 0);

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
 * The miniature usually comes from the list's cache, filled ahead of time by
 * its background pass. Failing that, an IntersectionObserver with a large
 * rootMargin queues a `renderSlide()` before the element scrolls into view.
 *
 * Once rendered, the miniature stays in the DOM. `content-visibility: auto`
 * lets the browser skip layout and paint for the offscreen ones, which is what
 * removing them used to buy, except that scrolling back has nothing to fill in
 * and so cannot flash a skeleton. Only the list's retention limits reclaim
 * them, furthest from the viewport first.
 */
export const ThumbnailItemPreview = React.forwardRef<HTMLDivElement, ThumbnailItemPreviewProps>(
  function ThumbnailItemPreview({ render, ...thumbnailItemPreviewProps }, forwardedRef) {
    const itemContext = useThumbnailItemContext(THUMBNAIL_ITEM_PREVIEW_NAME);
    const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_PREVIEW_NAME);
    const { presentation } = usePresentation();
    const store = useStoreContext(THUMBNAIL_ITEM_PREVIEW_NAME);

    const itemPreviewRef = React.useRef<HTMLDivElement>(null);
    const slideHandleRef = React.useRef<SlideHandle | null>(null);
    const hasRenderedRef = React.useRef(false);
    const isVisibleRef = React.useRef(false);
    const releaseRetainRef = React.useRef<(() => void) | null>(null);
    const {
      mediaUrlCache,
      slideHandleCache,
      observeResize,
      scheduleRender,
      retainPreview,
      setPreviewVisible,
      scrollRootRef,
    } = rovingContext;

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

    // The IntersectionObserver fires when this preview enters or leaves the
    // rootMargin runway. On entry:
    //   - Cache hit  → attach synchronously, which is the usual case once the
    //     list's background pass has been through this part of the deck
    //   - Cache miss → enqueue renderSlide() on the budgeted queue
    // Leaving only reports the preview as out of the runway: what is rendered
    // stays rendered until the list reclaims it, so scrolling back has nothing
    // to fill in.
    React.useEffect(() => {
      const itemPreviewElement = itemPreviewRef.current;
      if (!itemPreviewElement || !presentation || !slide) return;
      if (typeof IntersectionObserver === "undefined") return;

      const detach = () => {
        const slideHandle = slideHandleRef.current;
        slideHandleRef.current = null;
        slideHandle?.element.remove();
        hasRenderedRef.current = false;
        const element = itemPreviewRef.current;
        if (element) element.dataset.pending = "";
        return slideHandle;
      };

      // Reclaim path: unlike detach() this also drops the handle, since the
      // point of reclaiming is to stop paying for the miniature at all.
      const evict = () => {
        detach()?.dispose();
        slideHandleCache.delete(slide.id);
      };

      const attach = (element: HTMLDivElement, slideHandle: SlideHandle) => {
        if (slideHandleRef.current === slideHandle) return; // already attached
        const currentScale =
          widthRef.current > 0 ? widthRef.current / presentationWidthRef.current : 0;
        if (currentScale > 0) applySlideScale(slideHandle.element, currentScale);
        element.appendChild(slideHandle.element);
        slideHandleRef.current = slideHandle;
        hasRenderedRef.current = true;
        delete element.dataset.pending;

        releaseRetainRef.current?.();
        releaseRetainRef.current = retainPreview(slide.id, {
          index: slideIndex,
          hasChart: slideHasChart(slide),
          isAttached: true,
          isVisible: () => isVisibleRef.current,
          evict,
        });
      };

      let cancelRender: (() => void) | null = null;

      const intersectionObserver = new IntersectionObserver(
        (entries) => {
          const entry = entries.at(-1); // latest state wins
          if (!entry) return;
          isVisibleRef.current = entry.isIntersecting;
          setPreviewVisible(slideIndex, entry.isIntersecting);
          if (!entry.isIntersecting) return;

          const element = itemPreviewRef.current;
          if (!element || slideHandleRef.current) return;

          cancelRender?.();
          cancelRender = null;

          const cached = slideHandleCache.get(slide.id);
          if (cached && cached.revision === revisionRef.current) {
            attach(element, cached.slideHandle);
            return;
          }
          if (cached) {
            // Discard because it was rendered under an older edit revision.
            cached.slideHandle.dispose();
            slideHandleCache.delete(slide.id);
          }

          cancelRender = scheduleRender(element, () => {
            cancelRender = null;
            const mountedElement = itemPreviewRef.current;
            if (!mountedElement || slideHandleRef.current) return;
            const slideHandle = renderSlide(presentation, slide, { mediaUrlCache });
            slideHandleCache.set(slide.id, { slideHandle, revision: revisionRef.current });
            attach(mountedElement, slideHandle);
          });
        },
        { root: scrollRootRef.current, rootMargin: INTERSECTION_OBSERVER_ROOT_MARGIN },
      );

      intersectionObserver.observe(itemPreviewElement);

      return () => {
        intersectionObserver.disconnect();
        cancelRender?.();
        cancelRender = null;
        setPreviewVisible(slideIndex, false);
        releaseRetainRef.current?.();
        releaseRetainRef.current = null;
        // The handle stays cached so a remount (slide reorder, list re-key)
        // re-attaches it instead of rendering again; the list disposes the cache
        // when the presentation changes or it unmounts.
        detach();
      };
    }, [
      presentation,
      slide,
      slideIndex,
      mediaUrlCache,
      slideHandleCache,
      scheduleRender,
      retainPreview,
      setPreviewVisible,
      revisionRef,
      presentationWidthRef,
      scrollRootRef,
    ]);

    // Re-render the miniature in place when an edit bumps the revision,
    // without tearing down the IntersectionObserver (the detach→re-attach
    // gap was the cause of the thumbnail flash on every edit). Only fires
    // when the slide is currently visible/attached; off-screen slides are
    // re-rendered on demand by the IntersectionObserver callback using revisionRef.
    React.useEffect(() => {
      const element = itemPreviewRef.current;
      if (!element || !presentation || !slide) return;
      if (!slideHandleRef.current) return; // not visible; IO handles it

      const oldHandle = slideHandleRef.current;
      oldHandle.element.remove();
      oldHandle.dispose();
      slideHandleRef.current = null;
      slideHandleCache.delete(slide.id);

      const slideHandle = renderSlide(presentation, slide, { mediaUrlCache });
      const currentScale =
        widthRef.current > 0 ? widthRef.current / presentationWidthRef.current : 0;
      if (currentScale > 0) applySlideScale(slideHandle.element, currentScale);
      element.appendChild(slideHandle.element);
      slideHandleRef.current = slideHandle;
      slideHandleCache.set(slide.id, { slideHandle, revision });
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
              // Rendered miniatures are left attached, so the browser is what
              // skips work for the offscreen ones. This implies layout, style
              // and paint containment; the box keeps its height because width
              // and aspect-ratio give it one without measuring its contents.
              contentVisibility: "auto",
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
