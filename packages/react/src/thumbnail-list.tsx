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

// ---------------------------------------------------------------------------
// Roving focus utilities
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Component name constants
// ---------------------------------------------------------------------------

const THUMBNAIL_LIST_NAME = "PresentationThumbnailList";
const THUMBNAIL_ITEM_NAME = "PresentationThumbnailItem";
const THUMBNAIL_ITEM_PREVIEW_NAME = "PresentationThumbnailItemPreview";
const THUMBNAIL_ITEM_NUMBER_NAME = "PresentationThumbnailItemNumber";

// ---------------------------------------------------------------------------
// Internal contexts
// ---------------------------------------------------------------------------

interface ThumbnailRovingContextValue {
  currentTabStopId: string | null;
  // Ordered registry of item button elements, populated by each ThumbnailItem
  // on mount. Using a Map keyed by slideId preserves insertion order (which
  // matches DOM/slide order) and avoids querySelectorAll on every keypress.
  itemsRef: React.RefObject<Map<string, HTMLButtonElement>>;
  onItemFocus: (slideId: string) => void;
  onItemRegister: (slideId: string, el: HTMLButtonElement) => void;
  onItemUnregister: (slideId: string) => void;
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
  /** Zero-based position of this slide in the current slide list. */
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

// ---------------------------------------------------------------------------
// ThumbnailList
// ---------------------------------------------------------------------------

export interface ThumbnailListState {
  total: number;
  currentSlideId: string | null;
  /** Derived display index — use `currentSlideId` as identity. */
  currentIndex: number;
}

export interface ThumbnailListRenderState {
  slides: SlideData[];
  currentSlideId: string | null;
  /** Derived display index — use `currentSlideId` as identity. */
  currentIndex: number;
  goTo: (slideId: string) => void;
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
}

/**
 * Scrollable `listbox` container listing all slide thumbnails.
 * Renders nothing until the presentation is `"ready"`.
 *
 * Handles keyboard navigation (↑↓ / Home / End) with roving focus so
 * only the active item lives in the tab order at any time.
 *
 * @example
 * // Default — one ThumbnailItem per slide
 * <Presentation.ThumbnailList />
 *
 * @example
 * // Custom items via function children
 * <Presentation.ThumbnailList>
 *   {({ slides }) =>
 *     slides.map((slide) => (
 *       <Presentation.ThumbnailItem key={slide.id} slideId={slide.id}>
 *         <Presentation.ThumbnailItemPreview />
 *         <Presentation.ThumbnailItemNumber />
 *       </Presentation.ThumbnailItem>
 *     ))
 *   }
 * </Presentation.ThumbnailList>
 */
export const ThumbnailList = React.forwardRef<HTMLDivElement, ThumbnailListProps>(
  function ThumbnailList({ className, style, render, children, ...elementProps }, forwardedRef) {
    const { presentation, status } = usePresentation();
    const store = usePresentationStore(THUMBNAIL_LIST_NAME);

    // Which slide's button currently owns tabIndex=0.
    const [currentTabStopId, setCurrentTabStopId] = React.useState<string | null>(null);
    // Distinguishes mouse-click focus from keyboard focus so the entry-focus
    // redirect only fires for keyboard (matching Radix roving-focus behaviour).
    const isClickFocusRef = React.useRef(false);

    // Registry of slideId → button element, maintained by each ThumbnailItem.
    // Map insertion order matches slide/DOM order so keyboard navigation can
    // read from it directly without querySelectorAll.
    const itemsRef = React.useRef<Map<string, HTMLButtonElement>>(new Map());

    const currentSlideId = React.useSyncExternalStore(
      store.subscribe,
      () => store.getState().currentSlideId,
      () => null,
    );

    // Auto-focus the active (or first) thumbnail whenever a new presentation
    // loads. This lets arrow-key navigation work immediately without requiring
    // the user to Tab into the list first.
    React.useEffect(() => {
      if (!presentation) return;
      const items = itemsRef.current;
      // Prefer the active slide's button; fall back to the first registered item.
      const activeBtn = currentSlideId ? items.get(currentSlideId) : undefined;
      const firstBtn = activeBtn ?? items.values().next().value;
      firstBtn?.focus({ preventScroll: true });
    }, [presentation, currentSlideId]);

    // Fallback to the store's selected slide so its button gets tabIndex=0
    // before any keyboard interaction sets currentTabStopId explicitly.
    const effectiveTabStopId = currentTabStopId ?? currentSlideId;

    const rovingContextValue = React.useMemo<ThumbnailRovingContextValue>(
      () => ({
        currentTabStopId: effectiveTabStopId,
        itemsRef,
        onItemFocus: setCurrentTabStopId,
        onItemRegister: (slideId, el) => itemsRef.current.set(slideId, el),
        onItemUnregister: (slideId) => itemsRef.current.delete(slideId),
      }),
      [effectiveTabStopId],
    );

    if (status !== "ready" || !presentation) return null;

    const total = presentation.slides.length;
    const currentIndex = currentSlideId
      ? presentation.slides.findIndex((s) => s.id === currentSlideId)
      : -1;

    const state: ThumbnailListState = { total, currentSlideId, currentIndex };

    let resolvedChildren: React.ReactNode;
    if (typeof children === "function") {
      resolvedChildren = children({
        slides: presentation.slides,
        currentSlideId,
        currentIndex,
        goTo: (id) => store.goTo(id),
        goToIndex: (i) => store.goToIndex(i),
      });
    } else if (children != null) {
      resolvedChildren = children;
    } else {
      resolvedChildren = presentation.slides.map((slide) => (
        <ThumbnailItem key={slide.id} slideId={slide.id} />
      ));
    }

    return (
      <ThumbnailRovingContext.Provider value={rovingContextValue}>
        {renderElement(
          "div",
          { render, className, style },
          {
            state,
            ref: forwardedRef,
            props: {
              ...elementProps,
              role: "listbox",
              "aria-label": elementProps["aria-label"] ?? "Slide thumbnails",
              "aria-orientation": "vertical",
              // When a button owns tabIndex=0 the container steps out of the tab
              // order — the list has exactly ONE external tab stop (the active
              // button). Shift+Tab from the button then skips the container and
              // exits the list in a single key press.
              // When no button has a tab stop yet (e.g. before auto-focus fires),
              // the container acts as the entry point and redirects focus.
              tabIndex: effectiveTabStopId ? -1 : 0,
              style: { overflowY: "auto", outline: "none" },
              onMouseDown: (e: React.MouseEvent<HTMLDivElement>) => {
                elementProps.onMouseDown?.(e);
                if (e.target === e.currentTarget) isClickFocusRef.current = true;
              },
              onFocus: (e: React.FocusEvent<HTMLDivElement>) => {
                elementProps.onFocus?.(e);
                // Container only receives keyboard focus when effectiveTabStopId
                // is null (no button owns tabIndex=0 yet). Redirect to first button.
                if (e.target !== e.currentTarget) return;
                if (isClickFocusRef.current) {
                  isClickFocusRef.current = false;
                  return;
                }
                focusFirst(Array.from(itemsRef.current.values()), true);
              },
              children: resolvedChildren,
            },
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

// ---------------------------------------------------------------------------
// ThumbnailItem
// ---------------------------------------------------------------------------

export interface ThumbnailItemState {
  slideId: string;
  isActive: boolean;
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
   *
   * @example
   * render={(props, { isActive }) => (
   *   <motion.button {...props} animate={{ scale: isActive ? 1.04 : 1 }} />
   * )}
   */
  render?: RenderProp<ThumbnailItemState>;
}

/**
 * Clickable `option` button for a single slide in a `ThumbnailList`.
 *
 * - Identity is derived from `slide.id` — stable across list mutations.
 * - Auto-wires `onClick → goTo`, `data-active`, `aria-selected`, and roving
 *   `tabIndex` (0 when active, -1 otherwise).
 * - Provides context so nested `ThumbnailItemPreview` and `ThumbnailItemNumber`
 *   need no explicit props.
 * - Defaults to rendering `<ThumbnailItemPreview />` + `<ThumbnailItemNumber />`
 *   when no children are provided.
 */
export const ThumbnailItem = React.memo(
  React.forwardRef<HTMLButtonElement, ThumbnailItemProps>(function ThumbnailItem(
    { slideId, children, className, style, render, ...elementProps },
    forwardedRef,
  ) {
    const store = usePresentationStore(THUMBNAIL_ITEM_NAME);
    const rovingContext = useThumbnailRovingContext(THUMBNAIL_ITEM_NAME);

    const isActive = React.useSyncExternalStore(
      store.subscribe,
      () => store.getState().currentSlideId === slideId,
      () => false,
    );

    // Subscribe narrowly to the index itself (a number) rather than the whole
    // presentation object, so only actual slide reorders cause a re-render here.
    const displayIndex = React.useSyncExternalStore(
      store.subscribe,
      () => store.getState().presentation?.slides.findIndex((s) => s.id === slideId) ?? -1,
      () => -1,
    );

    // Register this button in the roving context's ordered map so the list
    // can navigate without querySelectorAll. The callback ref fires when the
    // DOM element is attached/detached.
    const registerRef = React.useCallback(
      (el: HTMLButtonElement | null) => {
        if (el) {
          rovingContext.onItemRegister(slideId, el);
        } else {
          rovingContext.onItemUnregister(slideId);
        }
      },
      // slideId is stable for the lifetime of this item instance
      // eslint-disable-next-line react-hooks/exhaustive-deps
      [slideId],
    );

    // This item owns the single tab stop inside the list when it matches the
    // roving context's currentTabStopId — all other items get tabIndex=-1.
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
          { render, className, style },
          {
            state,
            ref: [registerRef, forwardedRef],
            props: {
              ...elementProps,
              type: "button",
              role: "option",
              "aria-selected": isActive,
              "aria-label": `Slide ${displayIndex + 1}`,
              "data-active": isActive || undefined,
              "data-slide-id": slideId,
              tabIndex: isCurrentTabStop ? 0 : -1,
              style: {
                display: "block",
                width: "100%",
                padding: 0,
                border: "none",
                background: "none",
                overflow: "hidden",
                position: "relative",
              },
              onClick: () => store.goTo(slideId),
              onFocus: (e: React.FocusEvent<HTMLButtonElement>) => {
                elementProps.onFocus?.(e);
                rovingContext.onItemFocus(slideId);
                store.goTo(slideId);
              },
              onMouseDown: (e: React.MouseEvent<HTMLButtonElement>) => {
                elementProps.onMouseDown?.(e);
                rovingContext.onItemFocus(slideId);
              },
              onKeyDown: (e: React.KeyboardEvent<HTMLButtonElement>) => {
                elementProps.onKeyDown?.(e);

                if (e.target !== e.currentTarget) return;

                const focusIntent = MAP_KEY_TO_INTENT[e.key];
                if (!focusIntent || e.metaKey || e.ctrlKey || e.altKey || e.shiftKey) return;
                e.preventDefault();

                let candidates = Array.from(rovingContext.itemsRef.current.values());

                if (focusIntent === "last") {
                  candidates = candidates.reverse();
                } else if (focusIntent === "prev" || focusIntent === "next") {
                  if (focusIntent === "prev") candidates = candidates.reverse();
                  const idx = candidates.indexOf(e.currentTarget);
                  candidates = candidates.slice(idx + 1);
                }

                setTimeout(() => focusFirst(candidates));
              },
              children: children ?? (
                <>
                  <ThumbnailItemPreview />
                  <ThumbnailItemNumber />
                </>
              ),
            },
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

// ---------------------------------------------------------------------------
// ThumbnailItemPreview
// ---------------------------------------------------------------------------

export interface ThumbnailItemPreviewState {
  slideId: string;
  scale: number;
}

export interface ThumbnailItemPreviewProps extends React.ComponentProps<"div"> {
  /**
   * Replace the preview element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   *
   * The rendered element is the clipping container — the parsed slide DOM is
   * appended to it imperatively. Preserve `overflow: hidden` and dimensions.
   */
  render?: RenderProp<ThumbnailItemPreviewState>;
}

/**
 * Renders the slide miniature for the enclosing `ThumbnailItem`.
 *
 * Must be a descendant of `<Presentation.ThumbnailItem>` — reads slide ID
 * from context. Width is measured automatically via `ResizeObserver` so no
 * sizing props are required.
 *
 * The rendered element IS the clipping container: parsed slide DOM is appended
 * to it and CSS-scaled to fit. Marked `aria-hidden` since the enclosing
 * button's `aria-label` already identifies the slide.
 */
export const ThumbnailItemPreview = React.forwardRef<HTMLDivElement, ThumbnailItemPreviewProps>(
  function ThumbnailItemPreview({ className, style, render, ...elementProps }, forwardedRef) {
    const context = useThumbnailItemContext(THUMBNAIL_ITEM_PREVIEW_NAME);
    const { presentation } = usePresentation();

    const containerRef = React.useRef<HTMLDivElement>(null);
    const handleRef = React.useRef<SlideHandle | null>(null);
    const mediaUrlCache = React.useRef(new Map<string, string>()).current;

    const [containerWidth, setContainerWidth] = React.useState(0);
    React.useEffect(() => {
      const el = containerRef.current;
      if (!el) return;
      const ro = new ResizeObserver((entries) => {
        const entry = entries[0];
        if (entry) setContainerWidth(entry.contentRect.width);
      });
      ro.observe(el);
      return () => ro.disconnect();
    }, []);

    const slide = presentation?.slides.find((s) => s.id === context.slideId) ?? null;
    const pWidth = presentation?.width ?? 1;
    const pHeight = presentation?.height ?? 1;
    const scale = containerWidth > 0 ? containerWidth / pWidth : 0;
    const thumbHeight = pHeight * scale;

    // Render (or re-render) the slide DOM. Does NOT depend on `scale` — a
    // resize only changes the CSS transform, which is handled by the effect
    // below without tearing down and re-creating the slide element.
    React.useEffect(() => {
      const container = containerRef.current;
      if (!container || !presentation || !slide) return;

      if (handleRef.current) {
        handleRef.current.dispose();
        handleRef.current = null;
      }
      container.innerHTML = "";

      if (!slide.nodesMaterialized) materializeSlideNodes(presentation, slide);
      const handle = renderSlide(presentation, slide, { mediaUrlCache });
      handle.element.style.transformOrigin = "top left";
      container.appendChild(handle.element);
      handleRef.current = handle;

      return () => {
        if (handleRef.current) {
          handleRef.current.dispose();
          handleRef.current = null;
        }
      };
    }, [presentation, slide, mediaUrlCache]);

    // Apply scale imperatively — avoids a full slide teardown on every resize.
    React.useEffect(() => {
      if (!handleRef.current || scale === 0) return;
      handleRef.current.element.style.transform = `scale(${scale})`;
    }, [scale]);

    return renderElement(
      "div",
      { render, className, style },
      {
        state: { slideId: context.slideId, scale },
        ref: [containerRef, forwardedRef],
        props: {
          ...elementProps,
          "aria-hidden": "true",
          "data-active": context.isActive || undefined,
          // Prevent Tab from entering focusable PPTX content (links, forms, etc.)
          inert: true,
          style: {
            width: "100%",
            height: thumbHeight || undefined,
            overflow: "hidden",
            pointerEvents: "none",
          },
        },
      },
    );
  },
);

export namespace ThumbnailItemPreview {
  export type State = ThumbnailItemPreviewState;
  export type Props = ThumbnailItemPreviewProps;
}

// ---------------------------------------------------------------------------
// ThumbnailItemNumber
// ---------------------------------------------------------------------------

export interface ThumbnailItemNumberProps extends Omit<React.ComponentProps<"span">, "children"> {
  /**
   * Optionally replace the number span element.
   * - ReactElement: cloned with composed props
   * - Function: (props, state) => ReactElement
   */
  render?: RenderProp<{ isActive: boolean; displayIndex: number; slideId: string }>;
  children?: React.ReactNode;
}

/**
 * Renders the 1-based slide number for the enclosing `ThumbnailItem`.
 *
 * Must be a descendant of `<Presentation.ThumbnailItem>` — reads the display
 * number from context. Marked `aria-hidden` since the enclosing button's
 * `aria-label` already announces the slide number.
 *
 * Completely unstyled — add `className` / `style` for visual treatment.
 */
export const ThumbnailItemNumber = React.forwardRef<HTMLSpanElement, ThumbnailItemNumberProps>(
  function ThumbnailItemNumber(
    { className, style, children, render, ...elementProps },
    forwardedRef,
  ) {
    const context = useThumbnailItemContext(THUMBNAIL_ITEM_NUMBER_NAME);

    return renderElement(
      "span",
      { render, className, style },
      {
        state: {
          isActive: context.isActive,
          displayIndex: context.displayIndex,
          slideId: context.slideId,
        },
        ref: forwardedRef,
        props: {
          ...elementProps,
          "aria-hidden": "true",
          "data-active": context.isActive || undefined,
          style: { userSelect: "none" },
          children: children ?? context.displayIndex + 1,
        },
      },
    );
  },
);

export namespace ThumbnailItemNumber {
  export type Props = ThumbnailItemNumberProps;
}
