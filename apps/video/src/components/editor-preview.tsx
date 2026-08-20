import type * as React from "react";

import { Presentation } from "@diceui/pptx";

import { useDeck } from "@/hooks/use-deck";
import { panelBox } from "@/lib/theme";

interface EditorPreviewProps extends React.ComponentProps<"div"> {
  file: string;
  width: number;
  height: number;
  slideIndex?: number;
  railWidth?: number;
  padding?: number;
  reveal?: number;
  railIn?: number;
  railReveal?: number;
}

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
  ...props
}: EditorPreviewProps) {
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
        background: "rgba(0,0,0,.04)",
        maskImage: "linear-gradient(to bottom, black 78%, transparent)",
      }}
    >
      {({ slides }) =>
        slides.map((slide, index) => {
          const shown = Math.max(0, Math.min(1, (railReveal - index) / 1.8));
          return (
            <Presentation.ThumbnailItem
              key={slide.id}
              slideId={slide.id}
              style={{
                width: "100%",
                opacity: shown,
                filter: shown < 1 ? `blur(${(1 - shown) * 6}px)` : undefined,
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
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
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
        boxSizing: "border-box",
        overflow: "hidden",
        opacity: reveal,
        ...panelBox(reveal),
        ...style,
      }}
      {...props}
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
