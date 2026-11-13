import * as React from "react";

import type {
  PresentationData,
  SlideData,
  SlideHandle,
} from "@diceui/pptx-core";
import { materializeSlide, PPTX_ATTRS, renderSlide } from "@diceui/pptx-core";

import { DEFAULT_STORE_STATE } from "./constant";
import {
  useActiveSlide,
  useSlideContentRevision,
  useSlideRevision,
  useStoreContext,
  useStoreSelector,
} from "./context";
import { useLatestRef, useLazyRef } from "./hooks";
import type { PrimitiveProps } from "./render";
import { renderElement } from "./render";
import { getPasteboardOverhang } from "./selection/geometry";
import type { PresentationStatus } from "./store";

const SLIDE_NAME = "Presentation.Slide";

function buildNodeSnapshot(slide: SlideData) {
  const map = new Map<
    string,
    { x: number; y: number; w: number; h: number; rotation: number }
  >();
  for (const node of slide.nodes) {
    map.set(node.id, {
      x: node.position.x,
      y: node.position.y,
      w: node.size.w,
      h: node.size.h,
      rotation: node.rotation,
    });
  }
  return map;
}

export interface SlideState {
  /** Current load status, also set as `data-status` on the element. */
  status: PresentationStatus;

  /** 0-based index of the active slide in the loaded presentation. */
  index: number;
}

export interface SlideProps extends PrimitiveProps<"div", SlideState> {
  /**
   * Event handler called when a shape fails to render and is skipped. It replaces the default
   * `console.warn`.
   *
   * ```tsx
   * onNodeError={(nodeId, error) => reportError(error, { nodeId })}
   * ```
   */
  onNodeError?: (nodeId: string, error: unknown) => void;
}

export const Slide = React.forwardRef<HTMLDivElement, SlideProps>(
  function Slide(
    { children, render, onNodeError, ...slideProps },
    forwardedRef,
  ) {
    const store = useStoreContext(SLIDE_NAME);
    const presentation = useStoreSelector(store, (s) => s.presentation, null);
    const status = useStoreSelector(
      store,
      (s) => s.status,
      DEFAULT_STORE_STATE.status,
    );
    const zoom = useStoreSelector(store, (s) => s.zoom, 1);
    const slide = useActiveSlide(store);
    const index = React.useSyncExternalStore(
      store.subscribe,
      store.getActiveSlideIndex,
      () => -1,
    );

    const slideId = slide?.id;
    const revision = useSlideRevision(store, slideId);
    const contentRevision = useSlideContentRevision(store, slideId);

    const slideContent =
      presentation && slide ? (
        <SlideImpl
          presentation={presentation}
          slide={slide}
          zoom={zoom}
          revision={revision}
          contentRevision={contentRevision}
          onNodeError={onNodeError}
        >
          {children}
        </SlideImpl>
      ) : null;

    return renderElement(
      "div",
      { render },
      // oxlint-disable-next-line react/refs -- renderElement attaches the ref without reading it
      {
        state: { status, index },
        ref: forwardedRef,
        props: [
          {
            "aria-busy": status === "loading",
            "data-presentation-slide": "",
            "data-status": status,
            style: {
              display: "flex",
              // No flex centering, which makes start-side overflow unreachable; the slide uses `margin: auto`.
              overflow: "auto",
              // Otherwise the scroll container grows with its content at high zoom and pushes its
              // scrollbars off screen. It resolves to `none` when the parent's size is indefinite.
              maxWidth: "100%",
              maxHeight: "100%",
            },
            children: slideContent,
          },
          slideProps,
        ],
      },
    );
  },
);

interface SlideImplProps extends React.ComponentProps<"div"> {
  presentation: PresentationData;
  slide: SlideData;
  zoom: number;
  // A change forces a fresh render.
  revision: number;
  // Unchanged by edits that only move or resize nodes.
  contentRevision: number;
  onNodeError?: (nodeId: string, error: unknown) => void;
}

function SlideImpl({
  presentation,
  slide,
  zoom,
  revision,
  contentRevision,
  onNodeError,
  children,
  style,
  ...slideImplProps
}: SlideImplProps) {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const slideHandleRef = React.useRef<SlideHandle | null>(null);
  const mediaUrlCacheRef = useLazyRef(() => new Map<string, string>());
  const nodeSnapshotRef =
    React.useRef<
      Map<
        string,
        { x: number; y: number; w: number; h: number; rotation: number }
      >
    >(null);
  const contentRevisionRef = React.useRef<number>(null);
  const renderedSlideIdRef = React.useRef<string>(null);
  // In a ref so an inline handler doesn't rebuild the slide DOM.
  const onNodeErrorRef = useLatestRef(onNodeError);

  React.useLayoutEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    if (!slide.nodesMaterialized) materializeSlide(presentation, slide);

    // Only position changes are patched in place, since a resize needs re-layout. Different
    // slides can have identical node snapshots, like one full-bleed picture each.
    if (
      slideHandleRef.current &&
      nodeSnapshotRef.current &&
      renderedSlideIdRef.current === slide.id &&
      contentRevisionRef.current === contentRevision
    ) {
      const prevSnapshot = nodeSnapshotRef.current;
      const canPatch =
        slide.nodes.length === prevSnapshot.size &&
        slide.nodes.every((node) => {
          const prev = prevSnapshot.get(node.id);
          if (!prev) return false;
          return (
            prev.w === node.size.w &&
            prev.h === node.size.h &&
            prev.rotation === node.rotation
          );
        });

      if (canPatch) {
        const elements = new Map<string, HTMLElement>();
        for (const el of container.querySelectorAll<HTMLElement>(
          `[${PPTX_ATTRS.nodeId}]`,
        )) {
          const id = el.getAttribute(PPTX_ATTRS.nodeId);
          if (id !== null && !elements.has(id)) elements.set(id, el);
        }
        for (const node of slide.nodes) {
          const el = elements.get(node.id);
          if (el) {
            el.style.left = `${node.position.x}px`;
            el.style.top = `${node.position.y}px`;
          }
        }
        nodeSnapshotRef.current = buildNodeSnapshot(slide);
        contentRevisionRef.current = contentRevision;
        renderedSlideIdRef.current = slide.id;
        return;
      }
    }

    if (slideHandleRef.current) {
      slideHandleRef.current.dispose();
      slideHandleRef.current = null;
    }
    container.innerHTML = "";

    const isEditable = presentation.sourcePackage != null;
    const slideHandle = renderSlide(presentation, slide, {
      mediaUrlCache: mediaUrlCacheRef.current,
      placeholderPrompts: isEditable,
      // Shapes dragged past the slide edge stay visible, like PowerPoint's pasteboard.
      clipContent: !isEditable,
      onNodeError: (nodeId, error) => {
        const handler = onNodeErrorRef.current;
        if (handler) handler(nodeId, error);
        else console.warn(`[pptx] Node render error: ${nodeId}`, error);
      },
    });

    container.appendChild(slideHandle.element);
    slideHandleRef.current = slideHandle;
    nodeSnapshotRef.current = buildNodeSnapshot(slide);
    contentRevisionRef.current = contentRevision;
    renderedSlideIdRef.current = slide.id;
  }, [
    presentation,
    slide,
    revision,
    contentRevision,
    mediaUrlCacheRef,
    onNodeErrorRef,
  ]);

  React.useEffect(() => {
    const mediaUrlCache = mediaUrlCacheRef.current;
    return () => {
      if (slideHandleRef.current) {
        slideHandleRef.current.dispose();
        slideHandleRef.current = null;
      }
      for (const url of mediaUrlCache.values()) URL.revokeObjectURL(url);
      mediaUrlCache.clear();
    };
  }, [mediaUrlCacheRef]);

  const { width, height } = presentation;
  const isEditable = presentation.sourcePackage != null;

  // Scroll containers only extend right and down, so the pasteboard reserves room on every
  // overhanging side to keep shapes past the top or left edge reachable.
  const overhang = isEditable
    ? getPasteboardOverhang(slide.nodes, width, height)
    : { left: 0, top: 0, right: 0, bottom: 0 };

  return (
    <div
      style={{
        width: (width + overhang.left + overhang.right) * zoom,
        height: (height + overhang.top + overhang.bottom) * zoom,
        flexShrink: 0,
        position: "relative",
        // Unlike flex centering, auto margins collapse to 0 on overflow, keeping start edges scrollable.
        margin: "auto",
        ...style,
      }}
      {...slideImplProps}
    >
      {/* The selection overlay reads the slide's client rect from this frame, so it must match the slide bounds exactly. */}
      <div
        style={{
          position: "absolute",
          left: overhang.left * zoom,
          top: overhang.top * zoom,
          width: width * zoom,
          height: height * zoom,
        }}
      >
        <div
          ref={containerRef}
          style={{
            width,
            height,
            // Layout-time scaling keeps hairlines on whole device pixels at fractional zoom.
            zoom,
            position: "relative",
            // The inner container unclips in edit mode, so this wrapper must not clip either.
            overflow: isEditable ? "visible" : "hidden",
          }}
        />
        {children && (
          <div
            style={{ position: "absolute", inset: 0, pointerEvents: "none" }}
          >
            {children}
          </div>
        )}
      </div>
    </div>
  );
}

export namespace Slide {
  export type State = SlideState;
  export type Props = SlideProps;
}
