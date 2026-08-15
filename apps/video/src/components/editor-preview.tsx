import * as React from "react";

import { Presentation } from "@diceui/pptx";

import { geistSans } from "@/lib/fonts";
import { useDeck } from "@/lib/use-deck";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as React.CSSProperties;

/**
 * The selection overlay captures the pointer on press, which throws for an
 * event we synthesized because no such pointer is actually down. Neutering
 * capture for the duration of the dispatch lets the handler run to its
 * setState.
 */
function withoutPointerCapture(dispatch: () => void) {
  const { setPointerCapture, releasePointerCapture } = Element.prototype;
  Element.prototype.setPointerCapture = () => {};
  Element.prototype.releasePointerCapture = () => {};
  try {
    dispatch();
  } finally {
    Element.prototype.setPointerCapture = setPointerCapture;
    Element.prototype.releasePointerCapture = releasePointerCapture;
  }
}

/**
 * Presses the largest non-text shape so the frame shows real selection handles.
 * Text shapes are skipped because a click on one enters text editing and shows
 * a caret instead.
 */
function selectLargestShape(host: HTMLDivElement) {
  const overlay = host.querySelector<HTMLElement>("[data-pptx-selection]");
  const wrapper = host.querySelector('[data-status="ready"]');
  if (!overlay || !wrapper) return;

  const target = Array.from(wrapper.querySelectorAll<HTMLElement>("[data-pptx-node-id]"))
    .filter((node) => !node.querySelector("[data-pptx-p]"))
    .map((node) => ({ node, rect: node.getBoundingClientRect() }))
    .sort((a, b) => b.rect.width * b.rect.height - a.rect.width * a.rect.height)
    .at(0);
  if (!target) return;

  const clientX = target.rect.left + target.rect.width / 2;
  const clientY = target.rect.top + target.rect.height / 2;

  withoutPointerCapture(() => {
    for (const type of ["pointerdown", "pointerup"] as const) {
      overlay.dispatchEvent(
        new PointerEvent(type, {
          bubbles: true,
          cancelable: true,
          composed: true,
          clientX,
          clientY,
          button: 0,
          buttons: type === "pointerdown" ? 1 : 0,
          pointerId: 1,
          pointerType: "mouse",
          isPrimary: true,
        }),
      );
    }
  });
}

/**
 * How far along each part of the editor is, 0 to 1. The caller drives these off
 * the same timeline as the code snippet so the two stay in step.
 */
export interface EditorPreviewReveal {
  chrome: number;
  rail: number;
  canvas: number;
  slide: number;
  selection: number;
}

const FULLY_REVEALED: EditorPreviewReveal = {
  chrome: 1,
  rail: 1,
  canvas: 1,
  slide: 1,
  selection: 1,
};

export function EditorPreview({
  file,
  width,
  height,
  slideIndex = 0,
  railWidth = 130,
  padding = 26,
  reveal = FULLY_REVEALED,
  style,
}: {
  file: string;
  width: number;
  height: number;
  slideIndex?: number;
  railWidth?: number;
  padding?: number;
  reveal?: EditorPreviewReveal;
  style?: React.CSSProperties;
}) {
  const { data, hostRef, onLoad } = useDeck(file, selectLargestShape);

  return (
    <div
      ref={hostRef}
      style={{
        ...FONT_VARS,
        width,
        height,
        borderRadius: 18,
        overflow: "hidden",
        opacity: reveal.chrome,
        transform: `translateY(${(1 - reveal.chrome) * 20}px) scale(${0.98 + reveal.chrome * 0.02})`,
        background: "rgba(255,255,255,.04)",
        border: "1px solid rgba(255,255,255,.12)",
        boxShadow: `0 40px 120px rgba(0,0,0,${0.55 * reveal.chrome})`,
        ...style,
      }}
    >
      {data && (
        <Presentation.Root
          file={data}
          defaultSlideIndex={slideIndex}
          readOnly={false}
          onLoad={onLoad}
          style={{ display: "flex", width: "100%", height: "100%" }}
        >
          <Presentation.ThumbnailList
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 10,
              flex: "0 0 auto",
              width: railWidth,
              padding: 12,
              boxSizing: "border-box",
              overflow: "hidden",
              borderRight: "1px solid rgba(255,255,255,.1)",
              background: "rgba(0,0,0,.25)",
              // Opacity only: the previews observe their intersection with the
              // window, and moving them detaches and re-renders every frame.
              opacity: reveal.rail,
              // Fade the clipped last thumbnail so the rail reads as scrollable.
              maskImage: "linear-gradient(to bottom, black 78%, transparent)",
            }}
          />
          <Presentation.Viewport
            autoFit
            autoFitPadding={padding}
            style={{
              flex: 1,
              minWidth: 0,
              height: "100%",
              overflow: "hidden",
              // The slide sits in normal flow, so centre it rather than let any
              // leftover height pool at the bottom.
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              // Viewport draws nothing of its own, so give it a wash to make its
              // arrival visible when its line lands.
              background: `rgba(0,0,0,${0.22 * reveal.canvas})`,
            }}
          >
            <Presentation.Slide
              style={{
                opacity: reveal.slide,
                transform: `scale(${0.96 + reveal.slide * 0.04})`,
              }}
            >
              <Presentation.Selection style={{ opacity: reveal.selection }} />
            </Presentation.Slide>
          </Presentation.Viewport>
        </Presentation.Root>
      )}
    </div>
  );
}
