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
import { materializeSlideNodes, renderSlide } from "@diceui/pptx-parser";

import { usePresentation, usePresentationStore } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";

const THUMBNAIL_LIST_NAME = "PresentationThumbnailList";
const THUMBNAIL_ITEM_NAME = "PresentationThumbnailItem";
const THUMBNAIL_ITEM_PREVIEW_NAME = "PresentationThumbnailItemPreview";
const THUMBNAIL_ITEM_NUMBER_NAME = "PresentationThumbnailItemNumber";

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

function wrapArray<T>(array: T[], startIndex: number): T[] {
  return array.map((_, i) => array[(startIndex + i) % array.length] as T);
}

/**
 * Nearest scrollable ancestor (`overflow-y: auto | scroll | overlay`), or
 * `null` when the element scrolls with the window.
 *
 * Deliberately does NOT require `scrollHeight > clientHeight`: the scroll
 * container must be identified at observer-creation time, before slides have
 * loaded and stretched the container.
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

interface ScrollportRect {
  top: number;
  bottom: number;
  height: number;
}

/** Visible box of a scroll container (or the window), clipped to the window. */
function scrollportRectOf(scrollContainer: HTMLElement | null): ScrollportRect {
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight;
  if (!scrollContainer) {
    return { top: 0, bottom: viewportHeight, height: viewportHeight };
  }
  const rect = scrollContainer.getBoundingClientRect();
  const top = Math.max(rect.top, 0);
  const bottom = Math.min(rect.bottom, viewportHeight);
  return { top, bottom, height: Math.max(bottom - top, 0) };
}

/**
 * Dev-only render-queue instrumentation (Phase 0 of the thumbnail overhaul).
 *
 * Each drain frame reports how many visible-slide renders ran, how long the
 * frame's render work took, and the remaining backlog. Aggregates accumulate
 * on `window.__pptxThumbnailPerf` and a `pptx:thumbnail-perf` CustomEvent is
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
  currentTabStopId: string | null;
  loop: boolean;
  itemsRef: React.RefObject<Map<string, HTMLButtonElement>>;
  onItemFocus: (slideId: string) => void;
  onItemRegister: (slideId: string, el: HTMLButtonElement) => void;
  onItemUnregister: (slideId: string) => void;
  /** Shared object-URL cache so each image is decoded once across all previews. */
  mediaUrlCache: Map<string, string>;
  /**
   * Register with the list-level shared ResizeObserver.
   * Returns a cleanup function that unregisters the element.
   */
  observeResize: (el: Element, cb: (width: number) => void) => () => void;
  /**
   * Enqueue a slide render, drained per animation frame.
   * Entries currently visible in the scrollport render unconditionally this
   * frame (an empty visible thumbnail is worse UX than one slightly longer
   * frame). Off-screen entries render nearest-to-scrollport-first within an
   * 8ms per-frame budget so background fill never causes jank.
   * Returns a cancel function that removes the entry from the queue.
   */
  scheduleRender: (el: Element, fn: () => void) => () => void;
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
 * Handles keyboard navigation (↑↓ / Home / End) with roving focus so
 * only the active item lives in the tab order at any time.
 */
export const ThumbnailList = React.forwardRef<HTMLDivElement, ThumbnailListProps>(
  function ThumbnailList({ render, children, loop = false, ...thumbnailListProps }, forwardedRef) {
    const { presentation, status } = usePresentation();
    const store = usePresentationStore(THUMBNAIL_LIST_NAME);

    const [currentTabStopId, setCurrentTabStopId] = React.useState<string | null>(null);
    const isClickFocusRef = React.useRef(false);
    const itemsRef = React.useRef<Map<string, HTMLButtonElement>>(new Map());
    // One shared object-URL cache for all previews in this list. Each image
    // path is decoded once and reused, regardless of how many slides reference
    // it. Keyed by presentation identity so a new load always starts fresh.
    const mediaCacheRef = React.useRef<{ key: object; cache: Map<string, string> } | null>(null);
    if (!mediaCacheRef.current || mediaCacheRef.current.key !== presentation) {
      mediaCacheRef.current = { key: presentation ?? {}, cache: new Map() };
    }
    const mediaUrlCache = mediaCacheRef.current.cache;

    // --- Shared ResizeObserver ---
    // One ResizeObserver serves every preview in the list; per-item observers
    // add meaningful setup and bookkeeping overhead on large decks.
    const resizeCallbacksRef = React.useRef(new Map<Element, (width: number) => void>());
    const sharedRORef = React.useRef<ResizeObserver | null>(null);
    // The rail's scroll container — resolved lazily on the first observeResize
    // call (when a DOM element is available) and cached for the render queue's
    // visibility partitioning, avoiding per-frame ancestor walks.
    const scrollContainerRef = React.useRef<HTMLElement | null>(null);

    if (!sharedRORef.current && typeof ResizeObserver !== "undefined") {
      sharedRORef.current = new ResizeObserver((entries) => {
        for (const entry of entries) {
          resizeCallbacksRef.current.get(entry.target)?.(entry.contentRect.width);
        }
      });
    }

    const observeResize = React.useCallback((el: Element, cb: (width: number) => void) => {
      // Resolve the scroll container once from the first registered element.
      if (!scrollContainerRef.current) {
        scrollContainerRef.current = findScrollContainer(el);
      }
      resizeCallbacksRef.current.set(el, cb);
      sharedRORef.current?.observe(el);
      return () => {
        resizeCallbacksRef.current.delete(el);
        sharedRORef.current?.unobserve(el);
      };
    }, []);

    // --- Central render queue ---
    // All renderSlide() work is serialised here and drained per animation
    // frame. Entries inside the rail's visible scrollport render
    // unconditionally in the same frame (a visible blank is worse UX than a
    // single long frame); lookahead entries render nearest-first within a
    // time budget so pre-rendering never janks an in-progress scroll.
    const renderQueueRef = React.useRef<Array<{ el: Element; fn: () => void }>>([]);
    const drainRafRef = React.useRef<number | null>(null);

    const drainRenderQueue = React.useCallback(() => {
      drainRafRef.current = null;
      const queue = renderQueueRef.current;
      if (queue.length === 0) return;

      const scrollport = scrollportRectOf(scrollContainerRef.current);
      const scrollportCenter = (scrollport.top + scrollport.bottom) / 2;

      // Snapshot positions once, then partition visible vs. lookahead.
      const measured = queue.map((entry) => {
        const rect = entry.el.getBoundingClientRect();
        return {
          entry,
          visible: rect.bottom > scrollport.top && rect.top < scrollport.bottom,
          distance: Math.abs((rect.top + rect.bottom) / 2 - scrollportCenter),
        };
      });

      const visible = measured.filter((m) => m.visible).sort((a, b) => a.distance - b.distance);
      // Sorted farthest-first so pop() yields the nearest entry in O(1).
      const lookahead = measured.filter((m) => !m.visible).sort((a, b) => b.distance - a.distance);

      queue.length = 0;
      for (const m of lookahead) queue.push(m.entry);

      const frameStart = performance.now();
      let rendered = 0;

      for (const m of visible) {
        m.entry.fn();
        rendered += 1;
      }

      // ~8ms budget: leaves headroom in a 16.7ms frame for style/layout/paint
      // of the slides just built. Simple slides take 1-3ms, so several render
      // per frame; complex ones degrade to ~1 per frame naturally.
      const budgetMs = 8;
      while (queue.length > 0 && performance.now() - frameStart < budgetMs) {
        queue.pop()?.fn();
        rendered += 1;
      }

      recordThumbnailPerf(rendered, performance.now() - frameStart, queue.length);

      if (queue.length > 0) {
        drainRafRef.current = requestAnimationFrame(drainRenderQueue);
      }
    }, []);

    const scheduleRender = React.useCallback(
      (el: Element, fn: () => void): (() => void) => {
        const entry = { el, fn };
        renderQueueRef.current.push(entry);
        if (drainRafRef.current === null) {
          drainRafRef.current = requestAnimationFrame(drainRenderQueue);
        }
        return () => {
          const idx = renderQueueRef.current.indexOf(entry);
          if (idx !== -1) renderQueueRef.current.splice(idx, 1);
        };
      },
      [drainRenderQueue],
    );

    React.useEffect(() => {
      return () => {
        sharedRORef.current?.disconnect();
        sharedRORef.current = null;
        scrollContainerRef.current = null;
        if (drainRafRef.current !== null) {
          cancelAnimationFrame(drainRafRef.current);
          drainRafRef.current = null;
        }
      };
    }, []);

    const activeSlideId = React.useSyncExternalStore(
      store.subscribe,
      () => store.getState().activeSlideId,
      () => null,
    );

    // Auto-focus the active (or first) thumbnail whenever a new presentation
    // loads. This lets arrow-key navigation work immediately without requiring
    // the user to Tab into the list first.
    React.useEffect(() => {
      if (!presentation) return;
      const items = itemsRef.current;
      const activeItem = activeSlideId ? items.get(activeSlideId) : undefined;
      const firstItem = activeItem ?? items.values().next().value;
      firstItem?.focus({ preventScroll: true });
    }, [presentation, activeSlideId]);

    // Fallback to the store's selected slide so its button gets tabIndex=0
    // before any keyboard interaction sets currentTabStopId explicitly.
    const effectiveTabStopId = currentTabStopId ?? activeSlideId;

    const rovingContextValue = React.useMemo<ThumbnailRovingContextValue>(
      () => ({
        currentTabStopId: effectiveTabStopId,
        loop,
        itemsRef,
        onItemFocus: setCurrentTabStopId,
        onItemRegister: (slideId, el) => itemsRef.current.set(slideId, el),
        onItemUnregister: (slideId) => itemsRef.current.delete(slideId),
        mediaUrlCache,
        observeResize,
        scheduleRender,
      }),
      // mediaUrlCache identity is stable per-presentation (guarded by the ref
      // above); the observer/scheduler callbacks are stable useCallbacks.
      [effectiveTabStopId, loop, mediaUrlCache, observeResize, scheduleRender],
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
                // When a button owns tabIndex=0 the container steps out of the tab
                // order: the list has exactly ONE external tab stop (the active
                // button). Shift+Tab from the button then skips the container and
                // exits the list in a single key press.
                // When no button has a tab stop yet (e.g. before auto-focus fires),
                // the container acts as the entry point and redirects focus.
                tabIndex: effectiveTabStopId ? -1 : 0,
                onMouseDown: (event: React.MouseEvent<HTMLDivElement>) => {
                  if (event.target === event.currentTarget) isClickFocusRef.current = true;
                },
                onFocus: (event: React.FocusEvent<HTMLDivElement>) => {
                  // Container only receives keyboard focus when effectiveTabStopId
                  // is null (no button owns tabIndex=0 yet). Redirect to first button.
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

    // Register this button in the roving context's ordered map so the list
    // can navigate without querySelectorAll. The callback ref fires when the
    // DOM element is attached/detached.
    const registerRef = React.useCallback(
      (element: HTMLButtonElement | null) => {
        if (element) {
          rovingContext.onItemRegister(slideId, element);
        } else {
          rovingContext.onItemUnregister(slideId);
        }
      },
      // slideId is stable for the lifetime of this item instance
      // eslint-disable-next-line react-hooks/exhaustive-deps
      [slideId],
    );

    // This item owns the single tab stop inside the list when it matches the
    // roving context's currentTabStopId: all other items get tabIndex=-1.
    const isCurrentTabStop = rovingContext.currentTabStopId === slideId;

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

  /**
   * Css scale factor applied to the slide element
   * (container width / presentation width). `0` before the container is measured.
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
 */
export const ThumbnailItemPreview = React.forwardRef<HTMLDivElement, ThumbnailItemPreviewProps>(
  function ThumbnailItemPreview({ render, ...thumbnailItemPreviewProps }, forwardedRef) {
    const itemContext = useThumbnailItemContext(THUMBNAIL_ITEM_PREVIEW_NAME);
    const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_PREVIEW_NAME);
    const { presentation } = usePresentation();
    const store = usePresentationStore(THUMBNAIL_ITEM_PREVIEW_NAME);

    const itemPreviewRef = React.useRef<HTMLDivElement>(null);
    const slideHandleRef = React.useRef<SlideHandle | null>(null);
    // Shared list-level infrastructure: one media cache, one ResizeObserver,
    // one render queue.
    const { mediaUrlCache, observeResize, scheduleRender } = rovingContext;
    const [containerWidth, setContainerWidth] = React.useState(0);
    // Tracks whether the slide DOM has actually been appended, used imperatively
    // to remove data-pending without triggering a React re-render per slide.
    const hasRenderedRef = React.useRef(false);

    // Ref so the queued render callback can read the current scale without
    // being a dependency of the render effect (resize must not re-render slides).
    const scaleRef = React.useRef(0);

    // O(1) via the store's id→index map; avoids an O(N) slides.find() on
    // every render of every preview.
    const slideIndex = store.getSlideIndex(itemContext.slideId);
    const slide = presentation?.slides[slideIndex] ?? null;
    const pWidth = presentation?.width ?? 1;
    const pHeight = presentation?.height ?? 1;
    const scale = containerWidth > 0 ? containerWidth / pWidth : 0;
    scaleRef.current = scale;

    // Measure container width synchronously before first paint so the slide
    // element is created with the correct transform immediately.
    React.useLayoutEffect(() => {
      const el = itemPreviewRef.current;
      if (el && el.offsetWidth > 0) setContainerWidth(el.offsetWidth);
    }, []);

    // Wire the shared ResizeObserver to keep width in sync on subsequent
    // container resizes.
    React.useEffect(() => {
      const el = itemPreviewRef.current;
      if (!el) return;
      return observeResize(el, setContainerWidth);
    }, [observeResize]);

    // RENDER-ONCE MODEL — no lazy / IO-based gating.
    //
    // Every item schedules its render immediately (via the list's shared
    // priority queue) and never disposes on scroll. The queue handles
    // prioritisation: items currently visible in the scrollport render
    // unconditionally this frame; off-screen items fill in nearest-first
    // within the 8ms budget. Once rendered, the slide DOM stays for the
    // lifetime of the (presentation, slide) pair — visited regions can never
    // go blank again.
    //
    // This approach removed the IntersectionObserver entirely. The IO had a
    // fundamental timing issue in React 18 Strict Mode: React actually
    // destroys and recreates DOM nodes during its simulated unmount/remount
    // cycle for effect cleanup verification. When effects re-run and
    // `io.observe(el)` is called on freshly inserted nodes, the IO fires its
    // initial callback before the browser has performed a layout pass — every
    // `entry.boundingClientRect` is `{ width: 0, height: 0 }`, so no element
    // ever reports `isIntersecting: true`, and items after the initial seed
    // zone are permanently stuck with `data-pending` set.
    //
    // The cleanup runs only on unmount or when presentation/slide changes.
    // Scrolling never triggers it.
    React.useEffect(() => {
      const el = itemPreviewRef.current;
      if (!el || !presentation || !slide) return;

      const cancel = scheduleRender(el, () => {
        const element = itemPreviewRef.current;
        // Guard against a second render being queued before the first
        // cleanup runs (e.g. Strict Mode remount, effect identity change).
        if (!element || slideHandleRef.current) return;

        if (!slide.nodesMaterialized) materializeSlideNodes(presentation, slide);
        const slideHandle = renderSlide(presentation, slide, { mediaUrlCache });
        slideHandle.element.style.transformOrigin = "top left";
        if (scaleRef.current > 0) {
          slideHandle.element.style.transform = `scale(${scaleRef.current})`;
        }
        element.appendChild(slideHandle.element);
        slideHandleRef.current = slideHandle;
        // Flip data-pending imperatively: avoids triggering N React re-renders
        // while the initial fill is draining the queue.
        hasRenderedRef.current = true;
        delete element.dataset.pending;
      });

      return () => {
        cancel();
        if (slideHandleRef.current) {
          slideHandleRef.current.dispose();
          slideHandleRef.current = null;
        }
        const element = itemPreviewRef.current;
        if (element) {
          element.innerHTML = "";
          hasRenderedRef.current = false;
          element.dataset.pending = "";
        }
      };
    }, [presentation, slide, mediaUrlCache, scheduleRender, itemContext.slideId]);

    // Apply scale imperatively: avoids a full slide teardown on every resize.
    React.useEffect(() => {
      if (!slideHandleRef.current || scale === 0) return;
      slideHandleRef.current.element.style.transform = `scale(${scale})`;
    }, [scale]);

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
              // Full containment: mutations inside one preview (slide DOM
              // landing) cannot invalidate layout/paint of the rail or page,
              // and offscreen previews contribute no paint work. Safe with
              // size containment because the box is sized by width +
              // aspect-ratio, never by its contents.
              // `layout style paint` gives us layout isolation (mutations
              // inside don't reflow the page) and paint isolation (offscreen
              // previews are skipped by the compositor) WITHOUT `size`
              // containment. `contain: strict` includes size containment,
              // which suppresses `aspect-ratio`-derived height, leaving
              // elements at 0px tall — the IntersectionObserver then sees
              // zero-area targets and never fires `isIntersecting: true`.
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
