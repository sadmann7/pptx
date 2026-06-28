import * as React from "react";
import { usePresentation, useSlide, usePresentationStoreRef } from "../context";
import { renderSlide } from "@pptx/parser";
import type { SlideData, SlideHandle } from "@pptx/parser";
import { mergeRefs, renderElement } from "../utils/render";
import type { RenderProp } from "../utils/render";

// ---------------------------------------------------------------------------
// Internal contexts — cascade thumbWidth and slide identity through the tree
// ---------------------------------------------------------------------------

interface ThumbnailListCtxValue {
  thumbWidth: number;
}

/** Provides `thumbWidth` to descendant `ThumbnailItemCanvas` elements. */
const ThumbnailListCtx = React.createContext<ThumbnailListCtxValue | null>(null);

/**
 * Provides the stable slide ID (`SlideData.slidePath`) to a descendant
 * `ThumbnailItemCanvas` so it knows which slide to render without explicit props.
 */
const ThumbnailItemCtx = React.createContext<string | null>(null);

// ---------------------------------------------------------------------------
// ThumbnailList
// ---------------------------------------------------------------------------

export interface ThumbnailListState {
  total: number;
  currentSlideId: string | null;
  /** Derived display index — use `currentSlideId` as identity, not this. */
  currentIndex: number;
}

export interface ThumbnailListRenderState {
  slides: SlideData[];
  currentSlideId: string | null;
  /** Derived display index — use `currentSlideId` as identity, not this. */
  currentIndex: number;
  goTo: (slideId: string) => void;
  goToIndex: (index: number) => void;
}

export interface ThumbnailListProps extends Omit<React.HTMLAttributes<HTMLDivElement>, "children"> {
  /**
   * Width of each thumbnail in pixels. Default: `140`.
   * Also read from `style.width` for backward compatibility.
   */
  thumbWidth?: number;
  /**
   * Replace the list container element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<React.HTMLAttributes<HTMLDivElement>, ThumbnailListState>;
  /**
   * - Absent → default `ThumbnailItem + ThumbnailItemCanvas` list
   * - ReactNode → rendered as-is inside the container
   * - Function → called with slide state when the presentation is ready
   */
  children?: React.ReactNode | ((state: ThumbnailListRenderState) => React.ReactNode);
}

/**
 * Scrollable container listing all slide thumbnails.
 * Renders nothing until the presentation is `"ready"`.
 *
 * @example
 * // Default — one ThumbnailItem per slide
 * <Presentation.ThumbnailList thumbWidth={160} />
 *
 * @example
 * // Custom items via function children
 * <Presentation.ThumbnailList thumbWidth={160}>
 *   {({ slides, currentSlideId, goTo }) =>
 *     slides.map((slide, i) => (
 *       <Presentation.ThumbnailItem key={slide.slidePath} slideId={slide.slidePath}>
 *         <Presentation.ThumbnailItemCanvas />
 *         <span>{i + 1} / {slides.length}</span>
 *       </Presentation.ThumbnailItem>
 *     ))
 *   }
 * </Presentation.ThumbnailList>
 */
export const ThumbnailList = React.forwardRef<HTMLDivElement, ThumbnailListProps>(
  function ThumbnailList(
    { className, style, thumbWidth, render, children, ...elementProps },
    forwardedRef,
  ) {
    const { presentation, status } = usePresentation();
    const { slideId: currentSlideId, index: currentIndex, goTo, goToIndex } = useSlide();

    if (status !== "ready" || !presentation) return null;

    const resolvedThumbWidth = thumbWidth ?? (style?.width as number | undefined) ?? 140;

    const state: ThumbnailListState = {
      total: presentation.slides.length,
      currentSlideId,
      currentIndex,
    };

    let resolvedChildren: React.ReactNode;
    if (typeof children === "function") {
      resolvedChildren = children({
        slides: presentation.slides,
        currentSlideId,
        currentIndex,
        goTo,
        goToIndex,
      });
    } else if (children != null) {
      resolvedChildren = children;
    } else {
      resolvedChildren = presentation.slides.map((slide) => (
        <ThumbnailItem key={slide.slidePath} slideId={slide.slidePath} />
      ));
    }

    return (
      <ThumbnailListCtx.Provider value={{ thumbWidth: resolvedThumbWidth }}>
        {renderElement(
          "div",
          render as RenderProp<Record<string, unknown>, ThumbnailListState> | undefined,
          {
            ...elementProps,
            ref: forwardedRef,
            className,
            style: { overflowY: "auto", padding: 8, ...style },
            children: resolvedChildren,
          },
          state,
        )}
      </ThumbnailListCtx.Provider>
    );
  },
);

// ---------------------------------------------------------------------------
// ThumbnailItem
// ---------------------------------------------------------------------------

export interface ThumbnailItemState {
  slideId: string;
  isActive: boolean;
}

export interface ThumbnailItemProps extends Omit<
  React.HTMLAttributes<HTMLButtonElement>,
  "onClick"
> {
  /**
   * Stable slide identity (`SlideData.slidePath`). Required.
   * Use this instead of a positional index so the item remains correct
   * across reorders, insertions, and deletions.
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
 * Clickable thumbnail button for a single slide.
 *
 * Accepts `slideId` (`SlideData.slidePath`) as its identity — stable across
 * any slide list mutations. Auto-wires `onClick → goTo(slideId)` and
 * `data-active` from the store. Defaults to rendering `<ThumbnailItemCanvas />`
 * when no `children` are provided.
 *
 * Provides `ThumbnailItemCtx` so a nested `<ThumbnailItemCanvas>` knows
 * which slide to render without needing an explicit prop.
 */
export const ThumbnailItem = React.forwardRef<HTMLButtonElement, ThumbnailItemProps>(
  function ThumbnailItem(
    { slideId, children, className, style, render, ...elementProps },
    forwardedRef,
  ) {
    const store = usePresentationStoreRef();
    const isActive = React.useSyncExternalStore(
      store.subscribe.bind(store),
      () => store.getState().currentSlideId === slideId,
      () => false,
    );

    const state: ThumbnailItemState = { slideId, isActive };

    return (
      <ThumbnailItemCtx.Provider value={slideId}>
        {renderElement(
          "button",
          render as RenderProp<Record<string, unknown>, ThumbnailItemState> | undefined,
          {
            ...elementProps,
            ref: forwardedRef,
            type: "button",
            "data-active": isActive || undefined,
            className,
            style: {
              display: "block",
              width: "100%",
              marginBottom: 8,
              padding: 4,
              border: isActive ? "2px solid #3b82f6" : "1px solid #e5e7eb",
              borderRadius: 4,
              background: "#fff",
              cursor: "pointer",
              overflow: "hidden",
              position: "relative",
              ...style,
            },
            onClick: () => store.goTo(slideId),
            children: children ?? <ThumbnailItemCanvas />,
          },
          state,
        )}
      </ThumbnailItemCtx.Provider>
    );
  },
);

// ---------------------------------------------------------------------------
// ThumbnailItemCanvas
// ---------------------------------------------------------------------------

export interface ThumbnailItemCanvasState {
  slideId: string;
  scale: number;
}

export interface ThumbnailItemCanvasProps extends React.HTMLAttributes<HTMLDivElement> {
  /**
   * Override the thumb width used for scaling.
   * Falls back to the `thumbWidth` set on the parent `ThumbnailList`.
   */
  thumbWidth?: number;
  /**
   * Replace the canvas element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   *
   * The rendered element is the clipping container — slides are appended
   * to it imperatively. Preserve `overflow: hidden` and the correct
   * dimensions when customising.
   */
  render?: RenderProp<React.HTMLAttributes<HTMLDivElement>, ThumbnailItemCanvasState>;
}

/**
 * Renders the slide miniature for the enclosing `ThumbnailItem`.
 *
 * Must be a descendant of `<Presentation.ThumbnailItem>` (reads the slide ID
 * from context). Reads `thumbWidth` from the parent `ThumbnailList` context
 * as a fallback — no props are required in the common case.
 *
 * The rendered element IS the clipping container: the parsed slide DOM is
 * appended to it and CSS-scaled to fit. `containerRef` and `forwardedRef`
 * are merged onto it via `mergeRefs`.
 */
export const ThumbnailItemCanvas = React.forwardRef<HTMLDivElement, ThumbnailItemCanvasProps>(
  function ThumbnailItemCanvas(
    { thumbWidth: thumbWidthProp, className, style, render, ...elementProps },
    forwardedRef,
  ) {
    const slideId = React.useContext(ThumbnailItemCtx);
    const listCtx = React.useContext(ThumbnailListCtx);
    const { presentation } = usePresentation();

    if (slideId === null) {
      throw new Error(
        "[pptx/react] <Presentation.ThumbnailItemCanvas> must be rendered inside <Presentation.ThumbnailItem>",
      );
    }

    const containerRef = React.useRef<HTMLDivElement>(null);
    const handleRef = React.useRef<SlideHandle | null>(null);
    const mediaUrlCache = React.useRef(new Map<string, string>()).current;

    const slide = presentation?.slides.find((s) => s.slidePath === slideId) ?? null;
    const thumbWidth = thumbWidthProp ?? listCtx?.thumbWidth ?? 140;
    const pWidth = presentation?.width ?? 1;
    const pHeight = presentation?.height ?? 1;
    const scale = (thumbWidth - 16) / pWidth;
    const thumbHeight = pHeight * scale;

    React.useEffect(() => {
      const container = containerRef.current;
      if (!container || !presentation || !slide) return;

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

    const state: ThumbnailItemCanvasState = { slideId, scale };

    return renderElement(
      "div",
      render as RenderProp<Record<string, unknown>, ThumbnailItemCanvasState> | undefined,
      {
        ...elementProps,
        // containerRef drives the DOM manipulation; forwardedRef exposes the element to consumers
        ref: mergeRefs(containerRef, forwardedRef),
        className,
        style: {
          width: thumbWidth - 16,
          height: thumbHeight,
          overflow: "hidden",
          pointerEvents: "none",
          ...style,
        },
      },
      state,
    );
  },
);
