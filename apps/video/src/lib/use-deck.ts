import * as React from "react";

import { staticFile, useDelayRender } from "remotion";

/** Give up and release rather than hang Remotion's delayRender timeout. */
const SLIDE_WAIT_FRAMES = 120;

function waitForSlide(hostRef: React.RefObject<HTMLDivElement | null>, onReady: () => void) {
  let framesLeft = SLIDE_WAIT_FRAMES;
  let raf = 0;
  let stopped = false;

  const check = () => {
    if (stopped) return;
    const slide = hostRef.current?.querySelector('[data-status="ready"]');
    if (slide?.firstElementChild || framesLeft <= 0) {
      raf = requestAnimationFrame(() => {
        if (!stopped) onReady();
      });
      return;
    }
    framesLeft--;
    raf = requestAnimationFrame(check);
  };

  raf = requestAnimationFrame(check);
  return () => {
    stopped = true;
    cancelAnimationFrame(raf);
  };
}

/**
 * Fetches a deck and holds Remotion's render until the slide has painted with
 * its embedded fonts, so no frame is captured against an empty card.
 *
 * `onPainted` runs after paint but before the render is released, which is the
 * only window where imperative setup (dispatching a selection, say) still lands
 * in the captured frame.
 */
export function useDeck(file: string, onPainted?: (host: HTMLDivElement) => void) {
  const { delayRender, continueRender, cancelRender } = useDelayRender();
  const [handle] = React.useState(() => delayRender(`load ${file}`));
  const [data, setData] = React.useState<ArrayBuffer | null>(null);
  const hostRef = React.useRef<HTMLDivElement>(null);
  const resolvedRef = React.useRef(false);
  const stopWaitRef = React.useRef<(() => void) | null>(null);
  const onPaintedRef = React.useRef(onPainted);
  onPaintedRef.current = onPainted;

  const release = React.useCallback(() => {
    if (resolvedRef.current) return;
    resolvedRef.current = true;
    continueRender(handle);
  }, [continueRender, handle]);

  React.useEffect(() => {
    let cancelled = false;
    fetch(staticFile(file))
      .then((r) => {
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
      stopWaitRef.current?.();
      release();
    };
  }, [file, cancelRender, release]);

  const onLoad = React.useCallback(() => {
    stopWaitRef.current?.();
    stopWaitRef.current = waitForSlide(hostRef, () => {
      const fontsReady = document.fonts?.ready ?? Promise.resolve();
      void fontsReady.then(() => {
        const host = hostRef.current;
        if (host) onPaintedRef.current?.(host);
        // A second frame so anything onPainted mutated has painted too.
        requestAnimationFrame(() => requestAnimationFrame(release));
      });
    });
  }, [release]);

  return { data, hostRef, onLoad };
}
