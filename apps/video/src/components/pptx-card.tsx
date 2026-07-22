import * as React from "react";
import type { CSSProperties } from "react";

import type { PresentationStore } from "@diceui/pptx";
import { Presentation } from "@diceui/pptx";
import { continueRender, delayRender, staticFile } from "remotion";

import { geistSans } from "../fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as CSSProperties;

export function PptxCard({
  file,
  width,
  height,
  slideIndex = 0,
  style,
}: {
  file: string;
  width: number;
  height: number;
  slideIndex?: number;
  style?: CSSProperties;
}) {
  const [data, setData] = React.useState<ArrayBuffer | null>(null);
  const storeRef = React.useRef<PresentationStore | null>(null);
  const [handle] = React.useState(() => delayRender(`load ${file}`));

  React.useEffect(() => {
    let cancelled = false;
    fetch(staticFile(file))
      .then((r) => r.arrayBuffer())
      .then((buf) => {
        if (!cancelled) setData(buf);
      });
    return () => {
      cancelled = true;
    };
  }, [file]);

  return (
    <div style={{ ...FONT_VARS, width, height, ...style }}>
      {data && (
        <Presentation.Root
          file={data}
          onLoad={(store) => {
            storeRef.current = store;
            if (slideIndex > 0) store.goToIndex(slideIndex);
            continueRender(handle);
          }}
        >
          <Presentation.Viewport autoFit style={{ width, height, overflow: "hidden" }}>
            <Presentation.Slide />
          </Presentation.Viewport>
        </Presentation.Root>
      )}
    </div>
  );
}
