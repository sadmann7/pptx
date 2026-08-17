import * as React from "react";

import { Presentation } from "@diceui/pptx";

import { geistSans } from "@/lib/fonts";
import { useDeck } from "@/lib/use-deck";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as React.CSSProperties;

export function EditorPreview({
  file,
  width,
  height,
  slideIndex = 0,
  railWidth = 130,
  padding = 26,
  reveal = 1,
  showRail = true,
  slideOpacity = 1,
  through = false,
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
   * Whether the thumbnail rail is in the tree. Fixed for the life of the
   * instance: mounting it later would throw the IntersectionObserver and
   * blank every miniature.
   */
  showRail?: boolean;
  slideOpacity?: number;
  /**
   * Punch a hole through the canvas so a matching slide behind the window can
   * show through, used while the last showcase card flies into place.
   */
  through?: boolean;
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
        borderRight: "1px solid rgba(255,255,255,.1)",
        background: "rgba(0,0,0,.25)",
        // Fade the clipped last thumbnail so the rail reads as scrollable.
        maskImage: "linear-gradient(to bottom, black 78%, transparent)",
      }}
    />
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
        background: through ? "transparent" : "rgba(0,0,0,.22)",
      }}
    >
      <Presentation.Slide style={{ opacity: slideOpacity }}>
        <Presentation.Selection />
      </Presentation.Slide>
    </Presentation.Viewport>
  );

  return (
    <div
      ref={hostRef}
      style={{
        ...FONT_VARS,
        width,
        height,
        borderRadius: 18,
        overflow: "hidden",
        // Opacity only. A transform would move the thumbnails, and they render
        // themselves off their intersection with the window, so any movement
        // re-triggers that and blanks them mid-entrance.
        opacity: reveal,
        background: through ? "transparent" : "rgba(255,255,255,.04)",
        border: "1px solid rgba(255,255,255,.12)",
        boxShadow: `0 40px 120px rgba(0,0,0,${0.55 * reveal})`,
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
          }}
        >
          {showRail ? rail : null}
          {canvas}
        </Presentation.Root>
      )}
    </div>
  );
}
