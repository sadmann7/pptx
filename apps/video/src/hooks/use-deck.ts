import * as React from "react";

import { staticFile, useDelayRender } from "remotion";

const SLIDE_WAIT_FRAMES = 240;
const THUMBNAIL_SETTLE_FRAMES = 45;

function waitForSlide(hostRef: React.RefObject<HTMLDivElement | null>, onReady: () => void) {
  let framesLeft = SLIDE_WAIT_FRAMES;
  let raf = 0;
  let stopped = false;
  let pendingCount = -1;
  let steadyFrames = 0;

  function check() {
    if (stopped) return;
    const host = hostRef.current;
    const slide = host?.querySelector('[data-status="ready"]');
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
  }

  raf = requestAnimationFrame(check);
  return () => {
    stopped = true;
    cancelAnimationFrame(raf);
  };
}

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
        if (!r.ok) throw new Error(`${file}: ${r.status}.`);
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
