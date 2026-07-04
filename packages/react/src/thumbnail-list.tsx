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

// Visually-hidden style for the default slide number inside each thumbnail.
// Hoisted outside render so React.memo on ThumbnailItem sees a stable reference.
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

/**
 * Time-budgeted scheduler for thumbnail slide rendering.
 *
 * Parsing + DOM generation for a slide can take anywhere from 2ms to 300ms+.
 * Running every mounted preview synchronously in one commit lets a single
 * heavy slide block first paint for the whole page. Instead, tasks run in
 * order against a per-frame budget: work proceeds synchronously until the
 * budget is spent, then yields to the browser and resumes next frame.
 *
 * Cheap slides (the common case) still complete before the first paint; an
 * expensive slide only delays the thumbnails behind it in the queue.
 */
const THUMBNAIL_FRAME_BUDGET_MS = 12;

let thumbnailSyncSpentMs = 0;
let thumbnailPumpScheduled = false;
const thumbnailRenderQueue: Array<() => void> = [];

function pumpThumbnailQueue() {
  thumbnailPumpScheduled = false;
  thumbnailSyncSpentMs = 0;
  const frameStart = performance.now();
  while (
    thumbnailRenderQueue.length > 0 &&
    performance.now() - frameStart < THUMBNAIL_FRAME_BUDGET_MS
  ) {
    thumbnailRenderQueue.shift()?.();
  }
  if (thumbnailRenderQueue.length > 0) scheduleThumbnailPump();
}

function scheduleThumbnailPump() {
  if (thumbnailPumpScheduled) return;
  thumbnailPumpScheduled = true;
  requestAnimationFrame(pumpThumbnailQueue);
}

/**
 * Run `task` synchronously if the current frame still has render budget,
 * otherwise queue it for an upcoming frame. Returns a cancel function.
 */
function scheduleThumbnailRender(task: () => void): () => void {
  if (thumbnailSyncSpentMs < THUMBNAIL_FRAME_BUDGET_MS) {
    const start = performance.now();
    task();
    thumbnailSyncSpentMs += performance.now() - start;
    // Reset the sync budget on the next frame so later commits get their own.
    scheduleThumbnailPump();
    return () => {};
  }
  thumbnailRenderQueue.push(task);
  scheduleThumbnailPump();
  return () => {
    const index = thumbnailRenderQueue.indexOf(task);
    if (index !== -1) thumbnailRenderQueue.splice(index, 1);
  };
}

/**
 * Bounding rect of the nearest scrollable ancestor's visible area, clipped to
 * the window. Falls back to the window viewport when no scroll container exists.
 */
function findScrollportRect(element: HTMLElement): {
  top: number;
  bottom: number;
  height: number;
} {
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight;
  let ancestor = element.parentElement;
  while (ancestor) {
    const { overflowY } = getComputedStyle(ancestor);
    if (
      (overflowY === "auto" || overflowY === "scroll" || overflowY === "overlay") &&
      ancestor.scrollHeight > ancestor.clientHeight
    ) {
      const rect = ancestor.getBoundingClientRect();
      const top = Math.max(rect.top, 0);
      const bottom = Math.min(rect.bottom, viewportHeight);
      return { top, bottom, height: Math.max(bottom - top, 0) };
    }
    ancestor = ancestor.parentElement;
  }
  return { top: 0, bottom: viewportHeight, height: viewportHeight };
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
      }),
      // mediaUrlCache identity is stable per-presentation (guarded by the ref
      // above), so it's safe to include without triggering spurious re-renders.
      [effectiveTabStopId, loop, mediaUrlCache],
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

    const itemPreviewRef = React.useRef<HTMLDivElement>(null);
    const slideHandleRef = React.useRef<SlideHandle | null>(null);
    // Shared across all previews in the list — images decoded once, reused everywhere.
    const { mediaUrlCache } = rovingContext;
    const [containerWidth, setContainerWidth] = React.useState(0);
    // Previews render lazily: parsing + DOM generation for a slide only happens
    // once its item scrolls near the viewport. One-way latch (false → true).
    const [isNearViewport, setIsNearViewport] = React.useState(false);

    // Ref so the slide render effect can read the current scale without
    // being listed as a dependency (avoids tearing down the slide on resize).
    const scaleRef = React.useRef(0);

    const slide = presentation?.slides.find((s) => s.id === itemContext.slideId) ?? null;
    const pWidth = presentation?.width ?? 1;
    const pHeight = presentation?.height ?? 1;
    const scale = containerWidth > 0 ? containerWidth / pWidth : 0;
    scaleRef.current = scale;

    // Measure the container width synchronously before the first paint so that
    // the slide element is created with the correct transform right away.
    // Also check initial visibility here: waiting for the IntersectionObserver
    // (which fires after paint) would flash on-screen thumbnails empty for a
    // frame on first render.
    React.useLayoutEffect(() => {
      const itemPreviewElement = itemPreviewRef.current;
      if (!itemPreviewElement) return;
      if (itemPreviewElement.offsetWidth > 0) {
        setContainerWidth(itemPreviewElement.offsetWidth);
      }
      // Visibility must be measured against the thumbnail rail's scrollport,
      // not the window: items clipped by the rail's overflow are off-screen
      // even when they'd fall inside the window's bounds.
      const scrollport = findScrollportRect(itemPreviewElement);
      const rect = itemPreviewElement.getBoundingClientRect();
      // Same half-screenful lookahead as the IntersectionObserver's rootMargin.
      const lookahead = scrollport.height * 0.5;
      if (rect.top < scrollport.bottom + lookahead && rect.bottom > scrollport.top - lookahead) {
        setIsNearViewport(true);
      }
    }, []);

    // Keep width in sync on subsequent resizes.
    React.useEffect(() => {
      const itemPreviewElement = itemPreviewRef.current;
      if (!itemPreviewElement) return;
      const resizeObserver = new ResizeObserver((entries) => {
        const entry = entries[0];
        if (entry) setContainerWidth(entry.contentRect.width);
      });
      resizeObserver.observe(itemPreviewElement);
      return () => resizeObserver.disconnect();
    }, []);

    // Defer expensive slide parsing/rendering until this item is near the
    // scrollport. rootMargin pre-renders one screenful ahead so scrolling
    // rarely reveals an empty placeholder.
    React.useEffect(() => {
      if (isNearViewport) return;
      const itemPreviewElement = itemPreviewRef.current;
      if (!itemPreviewElement) return;
      if (typeof IntersectionObserver === "undefined") {
        setIsNearViewport(true);
        return;
      }
      const intersectionObserver = new IntersectionObserver(
        (entries) => {
          if (entries.some((entry) => entry.isIntersecting)) setIsNearViewport(true);
        },
        { rootMargin: "50% 0px" },
      );
      intersectionObserver.observe(itemPreviewElement);
      return () => intersectionObserver.disconnect();
    }, [isNearViewport]);

    // Render (or re-render) the slide DOM. Does NOT depend on `scale`: a
    // resize only changes the CSS transform, which is handled by the effect
    // below without tearing down and re-creating the slide element.
    // Layout effect + budgeted scheduler: cheap slides render before the first
    // paint (no placeholder flash); an expensive slide yields to the browser
    // instead of blocking every thumbnail behind it.
    React.useLayoutEffect(() => {
      if (!isNearViewport) return;
      if (!presentation || !slide) return;

      let disposed = false;

      const cancel = scheduleThumbnailRender(() => {
        const itemPreviewElement = itemPreviewRef.current;
        if (disposed || !itemPreviewElement) return;

        if (slideHandleRef.current) {
          slideHandleRef.current.dispose();
          slideHandleRef.current = null;
        }
        itemPreviewElement.innerHTML = "";

        if (!slide.nodesMaterialized) materializeSlideNodes(presentation, slide);
        const slideHandle = renderSlide(presentation, slide, { mediaUrlCache });
        slideHandle.element.style.transformOrigin = "top left";
        // Apply the current scale immediately so the slide is never visible at
        // full size before the separate scale effect fires.
        if (scaleRef.current > 0) {
          slideHandle.element.style.transform = `scale(${scaleRef.current})`;
        }
        itemPreviewElement.appendChild(slideHandle.element);
        slideHandleRef.current = slideHandle;
      });

      return () => {
        disposed = true;
        cancel();
        if (slideHandleRef.current) {
          slideHandleRef.current.dispose();
          slideHandleRef.current = null;
        }
      };
    }, [presentation, slide, mediaUrlCache, isNearViewport]);

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
            // Prevent Tab from entering focusable PPTX content (links, forms, etc.)
            inert: true,
            style: {
              width: "100%",
              // Prevents height-collapse layout shift while waiting for the ResizeObserver measurement.
              aspectRatio: `${pWidth} / ${pHeight}`,
              overflow: "hidden",
              pointerEvents: "none",
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
