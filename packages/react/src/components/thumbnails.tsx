import * as React from "react";
import { usePresentation, usePresentationStoreRef } from "../context";
import { renderSlide } from "@diceui/pptx-parser";
import type { SlideData, SlideHandle } from "@diceui/pptx-parser";
import { mergeRefs, renderElement } from "../utils/render";
import type { RenderProp } from "../utils/render";

// ---------------------------------------------------------------------------
// Roving focus utilities (pattern from @radix-ui/react-roving-focus)
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
// Internal contexts
// ---------------------------------------------------------------------------

interface ThumbnailRovingCtxValue {
  currentTabStopId: string | null;
  containerRef: React.RefObject<HTMLDivElement | null>;
  onItemFocus: (slideId: string) => void;
}

const ThumbnailRovingCtx = React.createContext<ThumbnailRovingCtxValue | null>(null);

interface ThumbnailItemCtxValue {
  slideId: string;
  /** Zero-based position of this slide in the current slide list. */
  displayIndex: number;
  isActive: boolean;
}

const ThumbnailItemCtx = React.createContext<ThumbnailItemCtxValue | null>(null);

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
  render?: RenderProp<React.ComponentProps<"div">, ThumbnailListState>;
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
 * <Presentation.ThumbnailList thumbWidth={160} />
 *
 * @example
 * // Custom items via function children
 * <Presentation.ThumbnailList thumbWidth={160}>
 *   {({ slides }) =>
 *     slides.map((slide, i) => (
 *       <Presentation.ThumbnailItem key={slide.id} slideId={slide.id}>
 *         <Presentation.ThumbnailItemPreview />
 *         <Presentation.ThumbnailItemIndex />
 *       </Presentation.ThumbnailItem>
 *     ))
 *   }
 * </Presentation.ThumbnailList>
 */
export const ThumbnailList = React.forwardRef<HTMLDivElement, ThumbnailListProps>(
  function ThumbnailList({ className, style, render, children, ...elementProps }, forwardedRef) {
    const { presentation, status } = usePresentation();
    const store = usePresentationStoreRef();
    const containerRef = React.useRef<HTMLDivElement>(null);

    // Which slide's button currently owns tabIndex=0.
    const [currentTabStopId, setCurrentTabStopId] = React.useState<string | null>(null);
    // Distinguishes mouse-click focus from keyboard focus so the entry-focus
    // redirect only fires for keyboard (matching Radix roving-focus behaviour).
    const isClickFocusRef = React.useRef(false);

    const currentSlideId = React.useSyncExternalStore(
      store.subscribe.bind(store),
      () => store.getState().currentSlideId,
      () => null,
    );

    // Auto-focus the active (or first) thumbnail whenever a new presentation
    // loads. This lets arrow-key navigation work immediately without requiring
    // the user to Tab into the list first.
    React.useEffect(() => {
      if (!presentation || !containerRef.current) return;
      const btn =
        containerRef.current.querySelector<HTMLButtonElement>("[data-active]") ??
        containerRef.current.querySelector<HTMLButtonElement>("button");
      btn?.focus({ preventScroll: true });
    }, [presentation]);

    if (status !== "ready" || !presentation) return null;

    const total = presentation.slides.length;
    const currentIndex = currentSlideId
      ? presentation.slides.findIndex((s) => s.id === currentSlideId)
      : -1;

    const state: ThumbnailListState = { total, currentSlideId, currentIndex };

    // Fallback to the store's selected slide so its button gets tabIndex=0
    // before any keyboard interaction sets currentTabStopId explicitly.
    const effectiveTabStopId = currentTabStopId ?? currentSlideId;

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
      <ThumbnailRovingCtx.Provider
        value={{
          currentTabStopId: effectiveTabStopId,
          containerRef,
          onItemFocus: setCurrentTabStopId,
        }}
      >
        {renderElement(
          "div",
          render as RenderProp<Record<string, unknown>, ThumbnailListState> | undefined,
          {
            ...elementProps,
            ref: mergeRefs(containerRef, forwardedRef),
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
            className,
            style: { overflowY: "auto", outline: "none", ...style },
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
              const el = containerRef.current;
              if (!el) return;
              const allBtns = Array.from(el.querySelectorAll<HTMLButtonElement>("button"));
              focusFirst(allBtns, true);
            },
            children: resolvedChildren,
          },
          state,
        )}
      </ThumbnailRovingCtx.Provider>
    );
  },
);

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
  render?: RenderProp<React.ComponentProps<"button">, ThumbnailItemState>;
}

/**
 * Clickable `option` button for a single slide in a `ThumbnailList`.
 *
 * - Identity is derived from `slide.id` — stable across list mutations.
 * - Auto-wires `onClick → goTo`, `data-active`, `aria-selected`, and roving
 *   `tabIndex` (0 when active, -1 otherwise).
 * - Provides context so nested `ThumbnailItemPreview` and `ThumbnailItemIndex`
 *   need no explicit props.
 * - Defaults to rendering `<ThumbnailItemPreview />` + `<ThumbnailItemIndex />`
 *   when no children are provided.
 */
export const ThumbnailItem = React.forwardRef<HTMLButtonElement, ThumbnailItemProps>(
  function ThumbnailItem(
    { slideId, children, className, style, render, ...elementProps },
    forwardedRef,
  ) {
    const store = usePresentationStoreRef();
    const rovingCtx = React.useContext(ThumbnailRovingCtx);

    const isActive = React.useSyncExternalStore(
      store.subscribe.bind(store),
      () => store.getState().currentSlideId === slideId,
      () => false,
    );

    const presentation = React.useSyncExternalStore(
      store.subscribe.bind(store),
      () => store.getState().presentation,
      () => null,
    );
    const displayIndex = presentation ? presentation.slides.findIndex((s) => s.id === slideId) : -1;

    // This item owns the single tab stop inside the list when it matches the
    // roving context's currentTabStopId — all other items get tabIndex=-1.
    const isCurrentTabStop = rovingCtx?.currentTabStopId === slideId;

    const state: ThumbnailItemState = { slideId, isActive, displayIndex };

    return (
      <ThumbnailItemCtx.Provider value={{ slideId, displayIndex, isActive }}>
        {renderElement(
          "button",
          render as RenderProp<Record<string, unknown>, ThumbnailItemState> | undefined,
          {
            ...elementProps,
            type: "button",
            role: "option",
            "aria-selected": isActive,
            "aria-label": `Slide ${displayIndex + 1}`,
            "data-active": isActive || undefined,
            // Queried by the container's entry-focus handler to restore the last stop.
            "data-slide-id": slideId,
            ref: forwardedRef,
            tabIndex: isCurrentTabStop ? 0 : -1,
            className,
            style: {
              display: "block",
              width: "100%",
              padding: 0,
              border: "none",
              background: "none",
              overflow: "hidden",
              position: "relative",
              ...style,
            },
            onClick: () => store.goTo(slideId),
            onFocus: (e: React.FocusEvent<HTMLButtonElement>) => {
              elementProps.onFocus?.(e);
              rovingCtx?.onItemFocus(slideId);
              store.goTo(slideId);
            },
            onMouseDown: (e: React.MouseEvent<HTMLButtonElement>) => {
              elementProps.onMouseDown?.(e);
              rovingCtx?.onItemFocus(slideId);
            },
            onKeyDown: (e: React.KeyboardEvent<HTMLButtonElement>) => {
              elementProps.onKeyDown?.(e);

              if (e.target !== e.currentTarget) return;

              const focusIntent = MAP_KEY_TO_INTENT[e.key];
              if (!focusIntent || e.metaKey || e.ctrlKey || e.altKey || e.shiftKey) return;
              e.preventDefault();

              const container = rovingCtx?.containerRef.current;
              if (!container) return;

              let candidates = Array.from(container.querySelectorAll<HTMLButtonElement>("button"));

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
          state,
        )}
      </ThumbnailItemCtx.Provider>
    );
  },
);

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
  render?: RenderProp<React.ComponentProps<"div">, ThumbnailItemPreviewState>;
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
    const ctx = React.useContext(ThumbnailItemCtx);
    const { presentation } = usePresentation();

    if (!ctx) {
      throw new Error(
        "[pptx/react] <Presentation.ThumbnailItemPreview> must be rendered inside <Presentation.ThumbnailItem>",
      );
    }

    const containerRef = React.useRef<HTMLDivElement>(null);
    const handleRef = React.useRef<SlideHandle | null>(null);
    const mediaUrlCache = React.useRef(new Map<string, string>()).current;

    // Measure the container's actual rendered width so the caller never has
    // to pass a manual thumbWidth — the scale always matches the CSS width.
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

    const slide = presentation?.slides.find((s) => s.id === ctx.slideId) ?? null;
    const pWidth = presentation?.width ?? 1;
    const pHeight = presentation?.height ?? 1;
    const scale = containerWidth > 0 ? containerWidth / pWidth : 0;
    const thumbHeight = pHeight * scale;

    React.useEffect(() => {
      const container = containerRef.current;
      // Skip until we have a measured width — avoids a wasted render at scale 0.
      if (!container || !presentation || !slide || scale === 0) return;

      if (handleRef.current) {
        handleRef.current.dispose();
        handleRef.current = null;
      }
      container.innerHTML = "";

      const handle = renderSlide(presentation, slide, { mediaUrlCache });
      handle.element.style.transform = `scale(${scale})`;
      handle.element.style.transformOrigin = "top left";
      container.appendChild(handle.element);
      handleRef.current = handle;

      return () => {
        if (handleRef.current) {
          handleRef.current.dispose();
          handleRef.current = null;
        }
      };
    }, [presentation, slide, scale, mediaUrlCache]);

    const state: ThumbnailItemPreviewState = { slideId: ctx.slideId, scale };

    return renderElement(
      "div",
      render as RenderProp<Record<string, unknown>, ThumbnailItemPreviewState> | undefined,
      {
        ...elementProps,
        ref: mergeRefs(containerRef, forwardedRef),
        "aria-hidden": "true",
        "data-active": ctx.isActive || undefined,
        // Prevent Tab from entering focusable PPTX content (links, forms, etc.)
        inert: true,
        className,
        style: {
          width: "100%",
          height: thumbHeight || undefined,
          overflow: "hidden",
          pointerEvents: "none",
          ...style,
        },
      },
      state,
    );
  },
);

// ---------------------------------------------------------------------------
// ThumbnailItemNumber
// ---------------------------------------------------------------------------

export interface ThumbnailItemNumberProps extends Omit<React.ComponentProps<"span">, "children"> {
  /**
   * Optionally replace the number span element.
   * - ReactElement: cloned with composed props
   * - Function: (props, state) => ReactElement
   */
  render?: RenderProp<
    React.ComponentProps<"span">,
    { isActive: boolean; displayIndex: number; slideId: string }
  >;
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
    const ctx = React.useContext(ThumbnailItemCtx);

    if (!ctx) {
      throw new Error(
        "[pptx/react] <Presentation.ThumbnailItemNumber> must be rendered inside <Presentation.ThumbnailItem>",
      );
    }

    const state = {
      isActive: ctx.isActive,
      displayIndex: ctx.displayIndex,
      slideId: ctx.slideId,
    };

    return renderElement(
      "span",
      render as RenderProp<Record<string, unknown>, typeof state> | undefined,
      {
        ...elementProps,
        ref: forwardedRef,
        "aria-hidden": "true",
        "data-active": ctx.isActive || undefined,
        className,
        style: { userSelect: "none", ...style },
        children: children ?? ctx.displayIndex + 1,
      },
      state,
    );
  },
);
