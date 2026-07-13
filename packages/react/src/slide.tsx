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
          "data-status": status,
          style: {
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
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

  React.useLayoutEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    if (!slide.nodesMaterialized) materializeSlide(presentation, slide);

    // Try to patch positions in-place instead of a full rebuild.
    // Safe when: same slide handle, node content unchanged (contentRevision),
    // and node sizes/rotations unchanged (a resize needs SVG/text re-layout).
    if (
      slideHandleRef.current &&
      nodeSnapshotRef.current &&
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
        return;
      }
    }

    // Full rebuild required.
    if (slideHandleRef.current) {
      slideHandleRef.current.dispose();
      slideHandleRef.current = null;
    }
    container.innerHTML = "";

    const slideHandle = renderSlide(presentation, slide, {
      mediaUrlCache: mediaUrlCacheRef.current,
      placeholderPrompts: presentation.sourcePackage != null,
      onNodeError: (nodeId, error) => {
        console.warn(`[pptx] Node render error: ${nodeId}`, error);
      },
    });

    container.appendChild(slideHandle.element);
    slideHandleRef.current = slideHandle;
    nodeSnapshotRef.current = buildNodeSnapshot(slide);
    contentRevisionRef.current = contentRevision;
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

  return (
    <div
      style={{ width: width * zoom, height: height * zoom, flexShrink: 0, position: "relative" }}
    >
      <div
        ref={containerRef}
        style={{
          width,
          height,
          transformOrigin: "top left",
          transform: `scale(${zoom})`,
          position: "relative",
          overflow: "hidden",
        }}
      />
      {children && (
        <div style={{ position: "absolute", inset: 0, pointerEvents: "none" }}>{children}</div>
      )}
    </div>
  );
}

export namespace Slide {
  export type State = SlideState;
  export type Props = SlideProps;
}
