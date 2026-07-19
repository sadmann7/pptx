import * as React from "react";

import type { PresentationData, SlideData, SlideHandle } from "@diceui/pptx-core";
import { materializeSlide, PPTX_ATTRS, renderSlide } from "@diceui/pptx-core";

import {
  usePresentation,
  useSlide,
  useSlideContentRevision,
  useSlideRevision,
  useStoreContext,
  useZoom,
} from "./context";
import { useLazyRef } from "./hook";
import type { RenderProp } from "./render";
import { renderElement } from "./render";
import { getPasteboardOverhang } from "./selection";
import type { PresentationStatus } from "./store";

const SLIDE_NAME = "Presentation.Slide";

function buildNodeSnapshot(slide: SlideData) {
  const map = new Map<string, { x: number; y: number; w: number; h: number; rotation: number }>();
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
  /**
   * Current parse/load status. Reflected as `data-status` on the wrapper
   * element so CSS can target states directly (e.g. `[data-status="loading"]`).
   */
  status: PresentationStatus;

  /** 0-based index of the active slide in the loaded presentation. */
  index: number;
}

export interface SlideProps extends React.ComponentProps<"div"> {
  /**
   * Replace the slide wrapper element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<SlideState>;
}

/**
 * Renders the active slide inside a centered wrapper `<div>`.
 *
 * The wrapper always mounts so that sibling layout is stable. Slide content
 * is absent until the presentation is `"ready"`. Use `<Presentation.Loading>`
 * and `<Presentation.Error>` alongside this component to cover other states.
 *
 * The wrapper carries a `data-status` attribute matching the store status,
 * enabling css-driven state styles without extra js.
 */
export const Slide = React.forwardRef<HTMLDivElement, SlideProps>(function Slide(
  { children, render, ...slideProps },
  forwardedRef,
) {
  const { presentation, status } = usePresentation();
  const { slide, index } = useSlide();
  const { zoom } = useZoom();
  const store = useStoreContext(SLIDE_NAME);

  // Bumped when an edit/undo/redo touches this slide; re-renders the content.
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
      >
        {children}
      </SlideImpl>
    ) : null;

  return renderElement(
    "div",
    { render },
    {
      state: { status, index },
      ref: forwardedRef,
      props: [
        {
          "aria-busy": status === "loading",
          "data-status": status,
          style: {
            display: "flex",
            // No align/justify centering: with a centered flex container the
            // start-side overflow becomes unreachable when content is bigger
            // than the viewport. The slide wrapper centers itself with
            // `margin: auto`, which degrades correctly to scrollable.
            overflow: "auto",
          },
          children: slideContent,
        },
        slideProps,
      ],
    },
  );
});

interface SlideImplProps {
  presentation: PresentationData;
  slide: SlideData;
  zoom: number;
  /** Edit revision of this slide; a change forces a fresh render. */
  revision: number;
  /** Content revision: unchanged when only node transforms were edited. */
  contentRevision: number;
  children?: React.ReactNode;
}

function SlideImpl({
  presentation,
  slide,
  zoom,
  revision,
  contentRevision,
  children,
}: SlideImplProps) {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const slideHandleRef = React.useRef<SlideHandle | null>(null);
  const mediaUrlCacheRef = useLazyRef(() => new Map<string, string>());
  const nodeSnapshotRef =
    React.useRef<Map<string, { x: number; y: number; w: number; h: number; rotation: number }>>(
      null,
    );
  const contentRevisionRef = React.useRef<number>(null);
  const renderedSlideIdRef = React.useRef<string>(null);

  React.useLayoutEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    if (!slide.nodesMaterialized) materializeSlide(presentation, slide);

    // Try to patch positions in-place instead of a full rebuild.
    // Safe when: same slide as the one currently rendered, node content
    // unchanged (contentRevision), and node sizes/rotations unchanged
    // (a resize needs SVG/text re-layout). The slide id check matters:
    // different slides can have structurally identical node snapshots
    // (e.g. one full-bleed picture per slide), which would otherwise
    // skip the rebuild when navigating between them.
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
            prev.w === node.size.w && prev.h === node.size.h && prev.rotation === node.rotation
          );
        });

      if (canPatch) {
        for (const node of slide.nodes) {
          const el = container.querySelector<HTMLElement>(
            `[${PPTX_ATTRS.nodeId}="${CSS.escape(node.id)}"]`,
          );
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

    // Full rebuild required.
    if (slideHandleRef.current) {
      slideHandleRef.current.dispose();
      slideHandleRef.current = null;
    }
    container.innerHTML = "";

    const isEditable = presentation.sourcePackage != null;
    const slideHandle = renderSlide(presentation, slide, {
      mediaUrlCache: mediaUrlCacheRef.current,
      placeholderPrompts: isEditable,
      // Editing surface: keep shapes dragged past the slide edge visible
      // (PowerPoint pasteboard behavior) instead of clipping at the bounds.
      clipContent: !isEditable,
      onNodeError: (nodeId, error) => {
        console.warn(`[pptx] Node render error: ${nodeId}`, error);
      },
    });

    container.appendChild(slideHandle.element);
    slideHandleRef.current = slideHandle;
    nodeSnapshotRef.current = buildNodeSnapshot(slide);
    contentRevisionRef.current = contentRevision;
    renderedSlideIdRef.current = slide.id;
  }, [presentation, slide, revision, contentRevision]);

  React.useEffect(() => {
    return () => {
      if (slideHandleRef.current) {
        slideHandleRef.current.dispose();
        slideHandleRef.current = null;
      }
    };
  }, []);

  const { width, height } = presentation;
  const isEditable = presentation.sourcePackage != null;

  // Pasteboard margin: scroll containers only extend their scrollable area
  // toward the right/bottom, so content overhanging the top/left of the
  // slide would be clipped and unreachable. Reserve room on every side the
  // shapes overhang, and offset the slide within it, so off-slide shapes can
  // always be scrolled into view.
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
        // Center within the flex viewport; unlike align/justify centering,
        // auto margins collapse to 0 when the content overflows, keeping the
        // start edges scrollable.
        margin: "auto",
      }}
    >
      {/*
       * Slide frame: sized and positioned to the slide itself. The selection
       * overlay resolves the slide's client rect from this element (its
       * grandparent from the selection root), so it must match the slide
       * bounds exactly.
       */}
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
            transformOrigin: "top left",
            transform: `scale(${zoom})`,
            position: "relative",
            // In edit mode the inner slide container unclips (pasteboard
            // behavior), so this wrapper must not reintroduce the clip.
            overflow: isEditable ? "visible" : "hidden",
          }}
        />
        {children && (
          <div style={{ position: "absolute", inset: 0, pointerEvents: "none" }}>{children}</div>
        )}
      </div>
    </div>
  );
}

export namespace Slide {
  export type State = SlideState;
  export type Props = SlideProps;
}
