/**
 * Time-budgeted scheduler for thumbnail slide rendering.
 *
 * Parsing + DOM generation for a slide can take anywhere from 2ms to 300ms+.
 * Running every mounted preview synchronously in one commit lets a single
 * heavy slide block first paint for the whole page. Instead, tasks run in
 * order against a per-frame budget: work proceeds synchronously until the
 * budget is spent, then yields to the browser and resumes next frame.
 *
 * Cheap slides (the common case) still complete before the first paint; an
 * expensive slide only delays the thumbnails behind it in the queue.
 */

const FRAME_BUDGET_MS = 12;

let syncSpentMs = 0;
let pumpScheduled = false;
const renderQueue: Array<() => void> = [];

function pump() {
  pumpScheduled = false;
  syncSpentMs = 0;
  const frameStart = performance.now();
  while (renderQueue.length > 0 && performance.now() - frameStart < FRAME_BUDGET_MS) {
    renderQueue.shift()?.();
  }
  if (renderQueue.length > 0) schedulePump();
}

function schedulePump() {
  if (pumpScheduled) return;
  pumpScheduled = true;
  requestAnimationFrame(pump);
}

/**
 * Run `task` synchronously if the current frame still has render budget,
 * otherwise queue it for an upcoming frame. Returns a cancel function.
 */
export function scheduleThumbnailRender(task: () => void): () => void {
  if (syncSpentMs < FRAME_BUDGET_MS) {
    const start = performance.now();
    task();
    syncSpentMs += performance.now() - start;
    // Reset the sync budget on the next frame so later commits get their own.
    schedulePump();
    return () => {};
  }
  renderQueue.push(task);
  schedulePump();
  return () => {
    const index = renderQueue.indexOf(task);
    if (index !== -1) renderQueue.splice(index, 1);
  };
}

/**
 * Bounding rect of the nearest scrollable ancestor's visible area, clipped to
 * the window. Falls back to the window viewport when no scroll container exists.
 */
export function findScrollportRect(element: HTMLElement): {
  top: number;
  bottom: number;
  height: number;
} {
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight;
  let ancestor = element.parentElement;
  while (ancestor) {
    const { overflowY } = getComputedStyle(ancestor);
    if (
      (overflowY === "auto" || overflowY === "scroll" || overflowY === "overlay") &&
      ancestor.scrollHeight > ancestor.clientHeight
    ) {
      const rect = ancestor.getBoundingClientRect();
      const top = Math.max(rect.top, 0);
      const bottom = Math.min(rect.bottom, viewportHeight);
      return { top, bottom, height: Math.max(bottom - top, 0) };
    }
    ancestor = ancestor.parentElement;
  }
  return { top: 0, bottom: viewportHeight, height: viewportHeight };
}
