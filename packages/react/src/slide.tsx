import * as React from "react";

import type { PresentationData, SlideData, SlideHandle } from "@diceui/pptx-core";
import { materializeSlide, renderSlide } from "@diceui/pptx-core";

import { TYPOGRAPHY_RESET_STYLE } from "./constant";
import { usePresentation, useSlide, useSlideRevision, useStoreContext, useZoom } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";
import type { PresentationStatus } from "./store";

const SLIDE_NAME = "Presentation.Slide";

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

  const slideContent =
    presentation && slide ? (
      <SlideImpl presentation={presentation} slide={slide} zoom={zoom} revision={revision}>
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
  children?: React.ReactNode;
}

function SlideImpl({ presentation, slide, zoom, revision, children }: SlideImplProps) {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const slideHandleRef = React.useRef<SlideHandle | null>(null);
  const mediaUrlCacheRef = React.useRef<Map<string, string>>(null);
  if (mediaUrlCacheRef.current === null) mediaUrlCacheRef.current = new Map();

  React.useLayoutEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    if (slideHandleRef.current) {
      slideHandleRef.current.dispose();
      slideHandleRef.current = null;
    }
    container.innerHTML = "";

    if (!slide.nodesMaterialized) materializeSlide(presentation, slide);
    const slideHandle = renderSlide(presentation, slide, {
      mediaUrlCache: mediaUrlCacheRef.current!,
      // Editable presentations (loaded with readOnly: false) get PowerPoint-style
      // dashed outlines and prompt text on empty placeholders.
      placeholderPrompts: presentation.sourcePackage != null,
      onNodeError: (nodeId, error) => {
        console.warn(`[pptx] Node render error: ${nodeId}`, error);
      },
    });

    container.appendChild(slideHandle.element);
    slideHandleRef.current = slideHandle;

    return () => {
      if (slideHandleRef.current) {
        slideHandleRef.current.dispose();
        slideHandleRef.current = null;
      }
    };
  }, [presentation, slide, revision]);

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
          ...TYPOGRAPHY_RESET_STYLE,
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
