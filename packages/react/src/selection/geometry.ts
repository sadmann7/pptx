import type { NodePosition, SlideNode } from "@diceui/pptx-core";

/** Degrees Shift steps a rotation by. */
export const ROTATION_STEP = 15;

/** An angle in degrees, wrapped into [0, 360). */
export function normalizeAngle(degrees: number): number {
  const wrapped = ((degrees % 360) + 360) % 360;
  // Floating-point error can leave 359.9999…, which is 0.
  return Math.abs(wrapped - 360) < 1e-6 ? 0 : wrapped;
}

/** The pointer's angle around a center, in degrees, clockwise from the x axis. */
export function getPointerAngle(centerX: number, centerY: number, x: number, y: number): number {
  return (Math.atan2(y - centerY, x - centerX) * 180) / Math.PI;
}

/** Minimum shape size (slide px) a resize can shrink to. */
export const MIN_SIZE = 8;

/** Screen-px movement before a pointer-down becomes a drag instead of a click. */
export const DRAG_THRESHOLD = 3;

export const GRIP_DIRECTIONS = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;
export type GripDirection = (typeof GRIP_DIRECTIONS)[number];

export const GRIP_CURSORS: Record<GripDirection, string> = {
  nw: "nwse-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  e: "ew-resize",
  se: "nwse-resize",
  s: "ns-resize",
  sw: "nesw-resize",
  w: "ew-resize",
};

const CORNER_GRIPS: ReadonlySet<GripDirection> = new Set(["nw", "ne", "se", "sw"]);

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Whether a press on empty canvas has travelled far enough to be a rubber band
 * rather than a click. Below the threshold nothing about the selection has
 * changed yet, so the band stays invisible and the handles stay up.
 */
export function isBandDrag(startX: number, startY: number, curX: number, curY: number): boolean {
  return Math.hypot(curX - startX, curY - startY) > DRAG_THRESHOLD;
}

/**
 * Applies a resize drag (slide-px deltas) to the original rect, clamped to
 * MIN_SIZE. With `lockAspect` (Shift held), corner handles scale both
 * dimensions proportionally around the opposite corner, like PowerPoint.
 */
export function resizeRect(
  origin: Rect,
  grip: GripDirection,
  dx: number,
  dy: number,
  lockAspect = false,
): Rect {
  let { x, y, w, h } = origin;

  if (grip.includes("e")) w = origin.w + dx;
  if (grip.includes("s")) h = origin.h + dy;
  if (grip.includes("w")) {
    w = origin.w - dx;
    x = origin.x + dx;
  }
  if (grip.includes("n")) {
    h = origin.h - dy;
    y = origin.y + dy;
  }

  if (lockAspect && CORNER_GRIPS.has(grip) && origin.w > 0 && origin.h > 0) {
    let scale =
      Math.abs(w / origin.w - 1) >= Math.abs(h / origin.h - 1) ? w / origin.w : h / origin.h;
    scale = Math.max(scale, MIN_SIZE / Math.min(origin.w, origin.h));
    w = origin.w * scale;
    h = origin.h * scale;
    if (grip.includes("w")) x = origin.x + origin.w - w;
    if (grip.includes("n")) y = origin.y + origin.h - h;
    return { x, y, w, h };
  }

  if (w < MIN_SIZE) {
    if (grip.includes("w")) x = origin.x + origin.w - MIN_SIZE;
    w = MIN_SIZE;
  }
  if (h < MIN_SIZE) {
    if (grip.includes("n")) y = origin.y + origin.h - MIN_SIZE;
    h = MIN_SIZE;
  }

  return { x, y, w, h };
}

/**
 * Applies a move drag (slide-px deltas). With `lockAxis` (Shift held) the
 * larger delta wins and the other is dropped, so the shape travels straight
 * along one axis like PowerPoint. Ties keep the horizontal delta.
 */
export function constrainMove(dx: number, dy: number, lockAxis = false): NodePosition {
  if (!lockAxis) return { x: dx, y: dy };
  return Math.abs(dx) >= Math.abs(dy) ? { x: dx, y: 0 } : { x: 0, y: dy };
}

/**
 * Union of an existing selection and the ids a marquee just enclosed, in
 * selection order. Additive banding only ever grows the selection: unlike
 * Shift+click it does not toggle an already selected shape back out, which
 * would otherwise drop shapes the band swept over.
 */
export function mergeSelection(baseIds: string[], addedIds: string[]): string[] {
  const baseIdsSet = new Set(baseIds);
  return [...baseIds, ...addedIds.filter((id) => !baseIdsSet.has(id))];
}

/**
 * Shift/Ctrl+click semantics: add the shape to the selection, or take it back
 * out when it is already selected.
 */
export function toggleSelection(selectedIds: string[], nodeId: string): string[] {
  return selectedIds.includes(nodeId)
    ? selectedIds.filter((id) => id !== nodeId)
    : [...selectedIds, nodeId];
}

export function getNodeRect(node: SlideNode): Rect {
  return { x: node.position.x, y: node.position.y, w: node.size.w, h: node.size.h };
}

/** Factors a resize drag scaled the dragged shape by, per axis. */
export interface ResizeScale {
  x: number;
  y: number;
}

/**
 * How much a resize drag scaled the shape whose handle was grabbed. A side
 * handle leaves the other axis at 1, and a zero-extent axis cannot define a
 * factor, so it reports 1 rather than dividing by zero.
 */
export function gripScale(
  origin: Rect,
  grip: GripDirection,
  dx: number,
  dy: number,
  lockAspect = false,
): ResizeScale {
  const next = resizeRect(origin, grip, dx, dy, lockAspect);
  return {
    x: origin.w === 0 ? 1 : next.w / origin.w,
    y: origin.h === 0 ? 1 : next.h / origin.h,
  };
}

/**
 * Scales a rect by the factors the dragged shape was scaled by, anchored at the
 * edges opposite the grip, matching `resizeRect`.
 *
 * PowerPoint resizes every shape in a multi-selection when one shape's handle
 * is dragged, but it scales each in place: the gaps between them do not scale,
 * which is why scaling a layout proportionally needs Scale Height/Width in the
 * Size pane instead. Applied to the dragged shape itself this reproduces
 * `resizeRect` exactly, clamps included.
 */
export function scaleRectFromGrip(rect: Rect, grip: GripDirection, scale: ResizeScale): Rect {
  const w = Math.max(MIN_SIZE, rect.w * scale.x);
  const h = Math.max(MIN_SIZE, rect.h * scale.y);
  return {
    x: grip.includes("w") ? rect.x + rect.w - w : rect.x,
    y: grip.includes("n") ? rect.y + rect.h - h : rect.y,
    w,
    h,
  };
}

export interface PasteboardOverhang {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * How far the slide's nodes overhang each slide edge, in slide-space px.
 * Zero on every side when all shapes sit within the slide bounds. Used to
 * size the pasteboard area around the slide in edit mode: both the visible
 * margin (so off-slide shapes stay scrollable into view) and the selection
 * event surface (so they stay clickable).
 */
export function getPasteboardOverhang(
  nodes: SlideNode[],
  slideWidth: number,
  slideHeight: number,
): PasteboardOverhang {
  const overhang = { left: 0, top: 0, right: 0, bottom: 0 };
  for (const node of nodes) {
    let { x, y, w, h } = getNodeRect(node);
    if (node.rotation !== 0) {
      const radians = (node.rotation * Math.PI) / 180;
      const halfW = (Math.abs(Math.cos(radians)) * w + Math.abs(Math.sin(radians)) * h) / 2;
      const halfH = (Math.abs(Math.sin(radians)) * w + Math.abs(Math.cos(radians)) * h) / 2;
      const cx = x + w / 2;
      const cy = y + h / 2;
      x = cx - halfW;
      y = cy - halfH;
      w = halfW * 2;
      h = halfH * 2;
    }
    overhang.left = Math.max(overhang.left, -x);
    overhang.top = Math.max(overhang.top, -y);
    overhang.right = Math.max(overhang.right, x + w - slideWidth);
    overhang.bottom = Math.max(overhang.bottom, y + h - slideHeight);
  }
  return overhang;
}
