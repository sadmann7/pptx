import * as React from "react";

import { Presentation } from "@diceui/pptx";
import { cancelRender, continueRender, delayRender, staticFile } from "remotion";

import { geistSans } from "@/lib/fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as React.CSSProperties;

/**
 * Frames to keep looking for the slide before giving up and releasing anyway.
 * A render that hangs here would fail on Remotion's own delayRender timeout
 * with nothing pointing at the cause, so a scene rendered slightly early is
 * the better failure.
 */
const SLIDE_WAIT_FRAMES = 120;

/**
 * Runs `onReady` once the slide is in the DOM and has had a frame to paint.
 * `Presentation.Slide` marks its wrapper `data-status="ready"` and only fills
 * it once the deck is parsed, so a wrapper with content is the signal.
 */
function waitForSlide(hostRef: React.RefObject<HTMLDivElement | null>, onReady: () => void) {
  let framesLeft = SLIDE_WAIT_FRAMES;

  const check = () => {
    const slide = hostRef.current?.querySelector('[data-status="ready"]');
    if (slide?.firstElementChild || framesLeft <= 0) {
      requestAnimationFrame(onReady);
      return;
    }
    framesLeft--;
    requestAnimationFrame(check);
  };

  check();
}

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
  const hostRef = React.useRef<HTMLDivElement>(null);
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
    <div ref={hostRef} style={{ ...FONT_VARS, width, height, ...style }}>
      {data && (
        <Presentation.Root
          file={data}
          defaultSlideIndex={slideIndex}
          onLoad={() => {
            // Parsing finishing is not the same as the slide being on screen:
            // `onLoad` runs a commit early, so releasing the handle here lets
            // Remotion screenshot an empty card on whichever frame this
            // component mounted on. Wait for the slide, then give it a frame
            // to paint.
            waitForSlide(hostRef, () => {
              resolvedRef.current = true;
              continueRender(pptxLoadToken);
            });
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
