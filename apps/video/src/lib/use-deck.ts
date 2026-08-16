import * as React from "react";

import { staticFile, useDelayRender } from "remotion";

/** Give up and release rather than hang Remotion's delayRender timeout. */
const SLIDE_WAIT_FRAMES = 240;

/**
 * Thumbnails render through a budgeted async queue, so the count of unfinished
 * ones drops in bursts. Previews clipped out of the rail never render at all, so
 * waiting for zero would always time out; waiting for the count to hold steady
 * this many frames means the queue has drained as far as it will. The window has
 * to outlast the longest gap between bursts, or a slow thumbnail looks like the
 * end of the queue and the deck releases half drawn. Remotion renders pages in
 * parallel, so the gap widens with CPU contention and the window has to allow
 * for the slowest case, not the one a single-threaded render shows.
 */
const THUMBNAIL_SETTLE_FRAMES = 45;

function waitForSlide(hostRef: React.RefObject<HTMLDivElement | null>, onReady: () => void) {
  let framesLeft = SLIDE_WAIT_FRAMES;
  let raf = 0;
  let stopped = false;
  let pendingCount = -1;
  let steadyFrames = 0;

  const check = () => {
    if (stopped) return;
    const host = hostRef.current;
    const slide = host?.querySelector('[data-status="ready"]');

    // Charts count too: they initialise a frame or more after the slide they sit
    // on, so a deck released on thumbnails alone captures its charts blank and
    // then has them appear partway through the shot.
    const nextPending =
      host?.querySelectorAll("[data-pending], [data-pptx-chart-pending]").length ?? 0;
    if (nextPending === pendingCount) steadyFrames++;
    else steadyFrames = 0;
    pendingCount = nextPending;

    const settled = slide?.firstElementChild && steadyFrames >= THUMBNAIL_SETTLE_FRAMES;
    if (settled || framesLeft <= 0) {
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
 */
export function useDeck(file: string) {
  const { delayRender, continueRender, cancelRender } = useDelayRender();
  const [handle] = React.useState(() => delayRender(`load ${file}`));
  const [data, setData] = React.useState<ArrayBuffer | null>(null);
  const hostRef = React.useRef<HTMLDivElement>(null);
  const resolvedRef = React.useRef(false);
  const stopWaitRef = React.useRef<(() => void) | null>(null);

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
        requestAnimationFrame(release);
      });
    });
  }, [release]);

  return { data, hostRef, onLoad };
}
