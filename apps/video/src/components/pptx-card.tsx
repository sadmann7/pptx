import * as React from "react";

import type { PresentationStore } from "@diceui/pptx";
import { Presentation } from "@diceui/pptx";
import { cancelRender, continueRender, delayRender, staticFile } from "remotion";

import { geistSans } from "@/lib/fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as React.CSSProperties;

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
  style?: React.CSSProperties;
}) {
  const [data, setData] = React.useState<ArrayBuffer | null>(null);
  const storeRef = React.useRef<PresentationStore | null>(null);
  const [pptxLoadToken] = React.useState(() => delayRender(`load ${file}`));
  const resolvedRef = React.useRef(false);

  React.useEffect(() => {
    let cancelled = false;
    fetch(staticFile(file))
      .then((r) => {
        // fetch resolves on 404, so an unchecked body would fail later as an
        // unreadable deck instead of a missing file.
        if (!r.ok) throw new Error(`${file}: ${r.status}`);
        return r.arrayBuffer();
      })
      .then((buf) => {
        if (!cancelled) setData(buf);
      })
      .catch((err) => {
        if (!cancelled) cancelRender(err);
      });
    return () => {
      cancelled = true;
      if (!resolvedRef.current) {
        continueRender(pptxLoadToken);
      }
    };
  }, [file, pptxLoadToken]);

  return (
    <div style={{ ...FONT_VARS, width, height, ...style }}>
      {data && (
        <Presentation.Root
          file={data}
          onLoad={(store) => {
            storeRef.current = store;
            if (slideIndex > 0) store.goToIndex(slideIndex);
            resolvedRef.current = true;
            continueRender(pptxLoadToken);
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
