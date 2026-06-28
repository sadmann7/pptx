import * as React from "react";
import { flushSync } from "react-dom";
import { usePresentation, usePresentationStoreRef } from "../context";
import { renderSlide } from "@diceui/pptx-parser";
import type { SlideData, SlideHandle } from "@diceui/pptx-parser";
import { mergeRefs, renderElement } from "../utils/render";
import type { RenderProp } from "../utils/render";

// ---------------------------------------------------------------------------
// Internal contexts
// ---------------------------------------------------------------------------

interface ThumbnailItemCtxValue {
  slideId: string;
  /** Zero-based position of this slide in the current slide list. */
  displayIndex: number;
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

export interface ThumbnailListProps extends Omit<React.HTMLAttributes<HTMLDivElement>, "children"> {
  /**
   * Replace the list container element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<React.HTMLAttributes<HTMLDivElement>, ThumbnailListState>;
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
    // Set to true before Shift+Tab exits an item so the container's onFocus
    // redirect is skipped — prevents the focus from cycling back inside the list.
    const skipNextEntryFocusRef = React.useRef(false);

    // Read current navigation state directly from the store (not via hook) to
    // avoid creating a snapshot object inside getSnapshot (infinite loop risk).
    const currentSlideId = React.useSyncExternalStore(
      store.subscribe.bind(store),
      () => store.getState().currentSlideId,
      () => null,
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
      <>
        {renderElement(
          "div",
          render as RenderProp<Record<string, unknown>, ThumbnailListState> | undefined,
          {
            ...elementProps,
            ref: mergeRefs(containerRef, forwardedRef),
            role: "listbox",
            "aria-label": elementProps["aria-label"] ?? "Slide thumbnails",
            "aria-orientation": "vertical",
            // Container is the single external tab stop (Radix roving-focus pattern).
            // All item buttons have tabIndex=-1; focus is 100% programmatic inside.
            tabIndex: elementProps.tabIndex ?? 0,
            className,
            style: { overflowY: "auto", outline: "none", ...style },
            onFocus: (e: React.FocusEvent<HTMLDivElement>) => {
              elementProps.onFocus?.(e);
              // Only handle focus that landed on the container itself (not bubbled
              // from a child button). This fires when Tab brings focus into the list.
              if (e.target !== e.currentTarget) return;
              if (skipNextEntryFocusRef.current) {
                // Shift+Tab exited an item — let focus rest on the container so
                // the next Shift+Tab naturally escapes the list entirely.
                skipNextEntryFocusRef.current = false;
                return;
              }
              // Redirect to the active item; fall back to the very first button.
              const target =
                containerRef.current?.querySelector<HTMLButtonElement>("[data-active]") ??
                containerRef.current?.querySelector<HTMLButtonElement>("button");
              target?.focus({ preventScroll: true });
            },
            onKeyDown: (e: React.KeyboardEvent<HTMLDivElement>) => {
              elementProps.onKeyDown?.(e);
              // Shift+Tab from an inner item would normally cycle back to the
              // container (tabIndex=0). Set the flag so onFocus skips the redirect,
              // letting the next Shift+Tab escape to the element before the list.
              if (e.key === "Tab" && e.shiftKey && e.target !== e.currentTarget) {
                skipNextEntryFocusRef.current = true;
                return;
              }

              const { presentation: p, currentSlideId: id } = store.getState();
              if (!p) return;

              const n = p.slides.length;
              const cur = id ? p.slides.findIndex((s) => s.id === id) : -1;

              let next: number | null = null;
              if (e.key === "ArrowDown") next = Math.min(cur + 1, n - 1);
              else if (e.key === "ArrowUp") next = Math.max(cur - 1, 0);
              else if (e.key === "Home") next = 0;
              else if (e.key === "End") next = n - 1;

              if (next === null || next === cur) return;
              e.preventDefault();

              // flushSync forces React to commit the updated isActive/tabIndex
              // state before we query the DOM for the newly active item.
              flushSync(() => {
                store.goToIndex(next!);
              });
              containerRef.current
                ?.querySelector<HTMLButtonElement>("[data-active]")
                ?.focus({ preventScroll: true });
            },
            children: resolvedChildren,
          },
          state,
        )}
      </>
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

export interface ThumbnailItemProps extends Omit<
  React.HTMLAttributes<HTMLButtonElement>,
  "onClick"
> {
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
  render?: RenderProp<React.HTMLAttributes<HTMLButtonElement>, ThumbnailItemState>;
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
    const buttonRef = React.useRef<HTMLButtonElement>(null);

    const isActive = React.useSyncExternalStore(
      store.subscribe.bind(store),
      () => store.getState().currentSlideId === slideId,
      () => false,
    );

    // presentation ref is stable across navigation — only changes on file load
    const presentation = React.useSyncExternalStore(
      store.subscribe.bind(store),
      () => store.getState().presentation,
      () => null,
    );
    const displayIndex = presentation ? presentation.slides.findIndex((s) => s.id === slideId) : -1;

    const state: ThumbnailItemState = { slideId, isActive, displayIndex };

    return (
      <ThumbnailItemCtx.Provider value={{ slideId, displayIndex }}>
        {renderElement(
          "button",
          render as RenderProp<Record<string, unknown>, ThumbnailItemState> | undefined,
          {
            ...elementProps,
            ref: mergeRefs(buttonRef, forwardedRef),
            type: "button",
            role: "option",
            "aria-selected": isActive,
            "aria-label": `Slide ${displayIndex + 1}`,
            "data-active": isActive || undefined,
            // All items are removed from the natural tab order; the ThumbnailList
            // container (tabIndex=0) is the single external entry point and
            // redirects focus here programmatically on keyboard entry.
            tabIndex: -1,
            className,
            style: {
              // Reset button UA styles; visual opinions belong in the styled layer.
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
            // Safari doesn't focus buttons on click — force focus so the
            // listbox keyboard handler stays connected to the right item.
            onMouseDown: () => {
              buttonRef.current?.focus({ preventScroll: true });
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

export interface ThumbnailItemPreviewProps extends React.HTMLAttributes<HTMLDivElement> {
  /**
   * Replace the preview element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   *
   * The rendered element is the clipping container — the parsed slide DOM is
   * appended to it imperatively. Preserve `overflow: hidden` and dimensions.
   */
  render?: RenderProp<React.HTMLAttributes<HTMLDivElement>, ThumbnailItemPreviewState>;
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

export interface ThumbnailItemNumberProps extends React.HTMLAttributes<HTMLSpanElement> {}

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
  function ThumbnailItemNumber({ className, style, children, ...elementProps }, forwardedRef) {
    const ctx = React.useContext(ThumbnailItemCtx);

    if (!ctx) {
      throw new Error(
        "[pptx/react] <Presentation.ThumbnailItemNumber> must be rendered inside <Presentation.ThumbnailItem>",
      );
    }

    return (
      <span
        {...elementProps}
        ref={forwardedRef}
        aria-hidden="true"
        className={className}
        style={{ userSelect: "none", ...style }}
      >
        {children ?? ctx.displayIndex + 1}
      </span>
    );
  },
);
