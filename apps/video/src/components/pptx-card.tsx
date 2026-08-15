import * as React from "react";

import { Presentation } from "@diceui/pptx";

import { geistSans } from "@/lib/fonts";
import { useDeck } from "@/lib/use-deck";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as React.CSSProperties;

interface PptxCardProps extends React.ComponentProps<"div"> {
  file: string;
  width: number;
  height: number;
  slideIndex?: number;
}

export function PptxCard({ file, width, height, slideIndex = 0, style, ...props }: PptxCardProps) {
  const { data, hostRef, onLoad } = useDeck(file);

  return (
    <div ref={hostRef} style={{ ...FONT_VARS, width, height, ...style }} {...props}>
      {data && (
        <Presentation.Root file={data} defaultSlideIndex={slideIndex} onLoad={onLoad}>
          <Presentation.Viewport autoFit style={{ width, height, overflow: "hidden" }}>
            <Presentation.Slide />
          </Presentation.Viewport>
        </Presentation.Root>
      )}
    </div>
  );
}
