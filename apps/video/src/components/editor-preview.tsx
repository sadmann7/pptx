import * as React from "react";

import { Presentation } from "@diceui/pptx";

import { panelShadow } from "@/lib/theme";
import { useDeck } from "@/lib/use-deck";

export function EditorPreview({
  file,
  width,
  height,
  slideIndex = 0,
  railWidth = 130,
  padding = 26,
  reveal = 1,
  railIn = 1,
  railReveal = Number.POSITIVE_INFINITY,
  style,
}: {
  file: string;
  width: number;
  height: number;
  slideIndex?: number;
  railWidth?: number;
  padding?: number;
  /** Entrance progress, 0 to 1. Only the window moves; the editor inside is static. */
  reveal?: number;
  /**
   * How far the rail has arrived, 0 to 1. The rail holds its final layout box
   * the whole time and only fades, because mounting it late or moving it out of
   * the window would throw the IntersectionObserver that decides whether each
   * miniature draws at all. At 0 the canvas is shifted over the hidden column
   * so the slide sits centred in the full window, which is what makes the
   * arrival read as the slide sliding aside to make room.
   */
  railIn?: number;
  /**
   * Miniatures faded in up to this index, so the rail fills top to bottom.
   * Opacity, blur and a few pixels of travel only, all of which stay inside the
   * rail's own box and so leave the intersection checks alone.
   */
  railReveal?: number;
  style?: React.CSSProperties;
}) {
  const { data, hostRef, onLoad } = useDeck(file);

  const rail = (
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
        opacity: railIn,
        borderRight: "1px solid rgba(255,255,255,.1)",
        // Sits on the canvas material below, so this only needs to be the
        // little extra darkening that separates the rail from the canvas.
        background: "rgba(0,0,0,.04)",
        // Fade the clipped last thumbnail so the rail reads as scrollable.
        maskImage: "linear-gradient(to bottom, black 78%, transparent)",
      }}
    >
      {({ slides }) =>
        slides.map((slide, index) => {
          // Each fade lasts longer than the gap between them, so the rail
          // cascades instead of ticking one item at a time.
          const shown = Math.max(0, Math.min(1, (railReveal - index) / 1.8));
          return (
            <Presentation.ThumbnailItem
              key={slide.id}
              slideId={slide.id}
              style={{
                width: "100%",
                opacity: shown,
                filter: shown < 1 ? `blur(${(1 - shown) * 6}px)` : undefined,
                // Travels in from the left edge of the rail's padding, so it
                // never leaves the column and gets clipped out of view.
                transform: `translateX(${-12 * (1 - shown)}px)`,
              }}
            >
              <Presentation.ThumbnailItemPreview />
            </Presentation.ThumbnailItem>
          );
        })
      }
    </Presentation.ThumbnailList>
  );

  const canvas = (
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
        // Inserting the rail moves the canvas centre right by half the rail's
        // width, so undoing exactly that puts the slide where it would sit in
        // an empty window. It holds nothing but the slide, so unlike the rail
        // it is free to move.
        transform: `translateX(${(-railWidth / 2) * (1 - railIn)}px)`,
      }}
    >
      <Presentation.Slide>
        <Presentation.Selection />
      </Presentation.Slide>
    </Presentation.Viewport>
  );

  return (
    <div
      ref={hostRef}
      style={{
        width,
        height,
        borderRadius: 18,
        overflow: "hidden",
        // Opacity only. A transform here would move the thumbnails, and they
        // render themselves off their intersection with the window, so any
        // movement re-triggers that and blanks them mid-entrance.
        opacity: reveal,
        background: "rgba(255,255,255,.04)",
        border: "1px solid rgba(255,255,255,.12)",
        boxShadow: panelShadow(reveal),
        ...style,
      }}
    >
      {data && (
        <Presentation.Root
          file={data}
          defaultSlideIndex={slideIndex}
          readOnly={false}
          onLoad={onLoad}
          style={{
            display: "flex",
            width: "100%",
            height: "100%",
            // Carried here rather than on the canvas so the rail column is the
            // same material while it is still hidden.
            background: "rgba(0,0,0,.22)",
          }}
        >
          {rail}
          {canvas}
        </Presentation.Root>
      )}
    </div>
  );
}
