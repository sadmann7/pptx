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

export const HANDLE_DIRECTIONS = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;
export type HandleDirection = (typeof HANDLE_DIRECTIONS)[number];

export const HANDLE_CURSORS: Record<HandleDirection, string> = {
  nw: "nwse-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  e: "ew-resize",
  se: "nwse-resize",
  s: "ns-resize",
  sw: "nesw-resize",
  w: "ew-resize",
};

const CORNER_HANDLES: ReadonlySet<HandleDirection> = new Set(["nw", "ne", "se", "sw"]);

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
  handle: HandleDirection,
  dx: number,
  dy: number,
  lockAspect = false,
): Rect {
  let { x, y, w, h } = origin;

  if (handle.includes("e")) w = origin.w + dx;
  if (handle.includes("s")) h = origin.h + dy;
  if (handle.includes("w")) {
    w = origin.w - dx;
    x = origin.x + dx;
  }
  if (handle.includes("n")) {
    h = origin.h - dy;
    y = origin.y + dy;
  }

  if (lockAspect && CORNER_HANDLES.has(handle) && origin.w > 0 && origin.h > 0) {
    let scale =
      Math.abs(w / origin.w - 1) >= Math.abs(h / origin.h - 1) ? w / origin.w : h / origin.h;
    scale = Math.max(scale, MIN_SIZE / Math.min(origin.w, origin.h));
    w = origin.w * scale;
    h = origin.h * scale;
    if (handle.includes("w")) x = origin.x + origin.w - w;
    if (handle.includes("n")) y = origin.y + origin.h - h;
    return { x, y, w, h };
  }

  if (w < MIN_SIZE) {
    if (handle.includes("w")) x = origin.x + origin.w - MIN_SIZE;
    w = MIN_SIZE;
  }
  if (h < MIN_SIZE) {
    if (handle.includes("n")) y = origin.y + origin.h - MIN_SIZE;
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

/** The slide-axis box a node covers once rotated about its center. */
export function getNodeBounds(node: SlideNode): Rect {
  const rect = getNodeRect(node);
  if (node.rotation === 0) return rect;
  const radians = (node.rotation * Math.PI) / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  const w = rect.w * cos + rect.h * sin;
  const h = rect.w * sin + rect.h * cos;
  return { x: rect.x + (rect.w - w) / 2, y: rect.y + (rect.h - h) / 2, w, h };
}

/**
 * Resizes a rotated shape from one of its handles. The drag is measured along
 * the shape's own axes, and the edges opposite the handle stay put on the
 * slide. `origin` is the box before rotation, which turns about its center.
 */
export function resizeRotatedRect(
  origin: Rect,
  rotation: number,
  handle: HandleDirection,
  dx: number,
  dy: number,
  lockAspect = false,
): Rect {
  if (rotation === 0) return resizeRect(origin, handle, dx, dy, lockAspect);
  const radians = (rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const local = resizeRect(origin, handle, dx * cos + dy * sin, -dx * sin + dy * cos, lockAspect);
  return placeRotated(origin, rotation, local);
}

/**
 * Where a rect resized in its own axes lands once rotated: the center moves by
 * the local offset turned into slide axes, so the anchored edges stay put.
 */
function placeRotated(origin: Rect, rotation: number, local: Rect): Rect {
  if (rotation === 0) return local;
  const radians = (rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const offsetX = local.x + local.w / 2 - (origin.x + origin.w / 2);
  const offsetY = local.y + local.h / 2 - (origin.y + origin.h / 2);
  const centerX = origin.x + origin.w / 2 + offsetX * cos - offsetY * sin;
  const centerY = origin.y + origin.h / 2 + offsetX * sin + offsetY * cos;
  return { x: centerX - local.w / 2, y: centerY - local.h / 2, w: local.w, h: local.h };
}

/** The resize cursor for a handle on a shape turned by `rotation` degrees. */
export function getHandleCursor(handle: HandleDirection, rotation: number): string {
  const steps = Math.round(normalizeAngle(rotation) / 45);
  return HANDLE_CURSORS[HANDLE_DIRECTIONS[(HANDLE_DIRECTIONS.indexOf(handle) + steps) % 8]];
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
export function handleScale(
  origin: Rect,
  handle: HandleDirection,
  dx: number,
  dy: number,
  lockAspect = false,
): ResizeScale {
  const next = resizeRect(origin, handle, dx, dy, lockAspect);
  return {
    x: origin.w === 0 ? 1 : next.w / origin.w,
    y: origin.h === 0 ? 1 : next.h / origin.h,
  };
}

/**
 * Scales a rect by the dragged shape's scale factors, anchored at the edges
 * opposite the handle. PowerPoint scales each shape in a multi-selection in place,
 * so the gaps between them don't scale. On the dragged shape itself this matches
 * `resizeRect` exactly, clamps included.
 */
export function scaleRectFromHandle(rect: Rect, handle: HandleDirection, scale: ResizeScale): Rect {
  const w = Math.max(MIN_SIZE, rect.w * scale.x);
  const h = Math.max(MIN_SIZE, rect.h * scale.y);
  return {
    x: handle.includes("w") ? rect.x + rect.w - w : rect.x,
    y: handle.includes("n") ? rect.y + rect.h - h : rect.y,
    w,
    h,
  };
}

/**
 * Where each selected shape lands mid-resize. The dragged handle's shape resizes along
 * its own axes, and every other shape scales by the same factors along its
 * own axes, each anchored at its own edges opposite the handle.
 */
export function getResizedRects(
  nodes: SlideNode[],
  handleNode: SlideNode,
  handle: HandleDirection,
  dx: number,
  dy: number,
  lockAspect: boolean,
): Map<string, Rect> {
  const origin = getNodeRect(handleNode);
  const resized = resizeRotatedRect(origin, handleNode.rotation, handle, dx, dy, lockAspect);
  const scale: ResizeScale = {
    x: origin.w === 0 ? 1 : resized.w / origin.w,
    y: origin.h === 0 ? 1 : resized.h / origin.h,
  };
  return new Map(
    nodes.map((node) => {
      if (node === handleNode) return [node.id, resized];
      const rect = getNodeRect(node);
      return [node.id, placeRotated(rect, node.rotation, scaleRectFromHandle(rect, handle, scale))];
    }),
  );
}

export interface PasteboardOverhang {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * How far the slide's nodes overhang each slide edge, in slide-space px, or zero
 * when everything fits. Sizes the edit-mode pasteboard, both the visible margin
 * and the selection surface, so off-slide shapes stay scrollable and clickable.
 */
export function getPasteboardOverhang(
  nodes: SlideNode[],
  slideWidth: number,
  slideHeight: number,
): PasteboardOverhang {
  const overhang = { left: 0, top: 0, right: 0, bottom: 0 };
  for (const node of nodes) {
    const { x, y, w, h } = getNodeBounds(node);
    overhang.left = Math.max(overhang.left, -x);
    overhang.top = Math.max(overhang.top, -y);
    overhang.right = Math.max(overhang.right, x + w - slideWidth);
    overhang.bottom = Math.max(overhang.bottom, y + h - slideHeight);
  }
  return overhang;
}
