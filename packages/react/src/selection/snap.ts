interface SnapRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Screen pixels within which an edge or center snaps. */
export const SNAP_THRESHOLD = 6;

/**
 * A guide line in slide pixels. An `x` guide is a vertical line at
 * `x = position`, running from `start` to `end` down the y axis; a `y` guide
 * is horizontal.
 */
export interface SnapGuide {
  axis: "x" | "y";
  position: number;
  start: number;
  end: number;
}

export interface SnapResult {
  dx: number;
  dy: number;
  guides: SnapGuide[];
}

type Axis = "x" | "y";

/** Start, center, and end of a rect along an axis. */
function getAnchors(rect: SnapRect, axis: Axis): [number, number, number] {
  const start = axis === "x" ? rect.x : rect.y;
  const size = axis === "x" ? rect.w : rect.h;
  return [start, start + size / 2, start + size];
}

/** The rect's extent across the other axis, which is where a guide runs. */
function getCrossSpan(rect: SnapRect, axis: Axis): [number, number] {
  return axis === "x" ? [rect.y, rect.y + rect.h] : [rect.x, rect.x + rect.w];
}

interface AxisSnap {
  correction: number;
  position: number;
}

/** The smallest correction that puts one of `edges` on a target anchor, if any is close enough. */
function getNearestSnap(
  edges: number[],
  targets: SnapRect[],
  axis: Axis,
  threshold: number,
): AxisSnap | null {
  let best: AxisSnap | null = null;
  for (const target of targets) {
    for (const anchor of getAnchors(target, axis)) {
      for (const edge of edges) {
        const correction = anchor - edge;
        if (Math.abs(correction) > threshold) continue;
        if (!best || Math.abs(correction) < Math.abs(best.correction)) {
          best = { correction, position: anchor };
        }
      }
    }
  }
  return best;
}

const EPSILON = 0.01;

/** A guide through `position` spanning the snapped rect and every target aligned on it. */
function getGuide(snapped: SnapRect, targets: SnapRect[], axis: Axis, position: number): SnapGuide {
  let [start, end] = getCrossSpan(snapped, axis);
  for (const target of targets) {
    if (!getAnchors(target, axis).some((anchor) => Math.abs(anchor - position) < EPSILON)) {
      continue;
    }
    const [targetStart, targetEnd] = getCrossSpan(target, axis);
    start = Math.min(start, targetStart);
    end = Math.max(end, targetEnd);
  }
  return { axis, position, start, end };
}

function getGuides(
  snapped: SnapRect,
  targets: SnapRect[],
  snaps: Partial<Record<Axis, AxisSnap | null>>,
): SnapGuide[] {
  const guides: SnapGuide[] = [];
  for (const axis of ["x", "y"] as const) {
    const snap = snaps[axis];
    if (snap) guides.push(getGuide(snapped, targets, axis, snap.position));
  }
  return guides;
}

/**
 * Snaps a move. `rect` is the dragged bounds before the move (the union of a
 * multi-selection), `dx`/`dy` the drag so far. An axis set to `false` in
 * `axes` keeps its delta, for Shift-constrained drags.
 */
export function snapMove(
  rect: SnapRect,
  dx: number,
  dy: number,
  targets: SnapRect[],
  threshold: number,
  axes: { x: boolean; y: boolean } = { x: true, y: true },
): SnapResult {
  const moved = { ...rect, x: rect.x + dx, y: rect.y + dy };
  const xSnap = axes.x ? getNearestSnap(getAnchors(moved, "x"), targets, "x", threshold) : null;
  const ySnap = axes.y ? getNearestSnap(getAnchors(moved, "y"), targets, "y", threshold) : null;

  const nextDx = dx + (xSnap?.correction ?? 0);
  const nextDy = dy + (ySnap?.correction ?? 0);
  const snapped = { ...rect, x: rect.x + nextDx, y: rect.y + nextDy };
  return { dx: nextDx, dy: nextDy, guides: getGuides(snapped, targets, { x: xSnap, y: ySnap }) };
}

/**
 * Snaps a resize: only the edges the handle moves snap. `handle` is a compass
 * direction ("e", "nw", ...). Corner drags with the aspect ratio locked are
 * left alone, since snapping one axis would break the ratio.
 */
export function snapResize(
  rect: SnapRect,
  handle: string,
  dx: number,
  dy: number,
  targets: SnapRect[],
  threshold: number,
  lockAspect = false,
): SnapResult {
  if (lockAspect && handle.length === 2) return { dx, dy, guides: [] };

  const movesLeft = handle.includes("w");
  const movesRight = handle.includes("e");
  const movesTop = handle.includes("n");
  const movesBottom = handle.includes("s");

  const xEdge = movesLeft ? rect.x + dx : movesRight ? rect.x + rect.w + dx : null;
  const yEdge = movesTop ? rect.y + dy : movesBottom ? rect.y + rect.h + dy : null;
  const xSnap = xEdge === null ? null : getNearestSnap([xEdge], targets, "x", threshold);
  const ySnap = yEdge === null ? null : getNearestSnap([yEdge], targets, "y", threshold);

  const nextDx = dx + (xSnap?.correction ?? 0);
  const nextDy = dy + (ySnap?.correction ?? 0);
  const snapped = {
    x: movesLeft ? rect.x + nextDx : rect.x,
    y: movesTop ? rect.y + nextDy : rect.y,
    w: rect.w + (movesLeft ? -nextDx : movesRight ? nextDx : 0),
    h: rect.h + (movesTop ? -nextDy : movesBottom ? nextDy : 0),
  };
  return { dx: nextDx, dy: nextDy, guides: getGuides(snapped, targets, { x: xSnap, y: ySnap }) };
}

export function getUnionRect(rects: SnapRect[]): SnapRect | null {
  if (rects.length === 0) return null;
  let x1 = Infinity;
  let y1 = Infinity;
  let x2 = -Infinity;
  let y2 = -Infinity;
  for (const rect of rects) {
    x1 = Math.min(x1, rect.x);
    y1 = Math.min(y1, rect.y);
    x2 = Math.max(x2, rect.x + rect.w);
    y2 = Math.max(y2, rect.y + rect.h);
  }
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
}
