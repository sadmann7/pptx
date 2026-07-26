import * as React from "react";

import type {
  NodePosition,
  SetTextBodyParagraph,
  ShapeNodeData,
  SlideNode,
} from "@diceui/pptx-core";
import { PPTX_ATTRS, PPTX_DATASET } from "@diceui/pptx-core";

import { usePresentation, useSlide, useSlideRevision, useStoreContext, useZoom } from "./context";
import { useLatestRef } from "./hook";
import type { RenderProp } from "./render";
import { mergeRefs, renderElement } from "./render";

const SELECTION_NAME = "Presentation.Selection";

const ENABLE_DEBUG_LOG = false;

/** Screen-px thickness of the selection-frame grab hitboxes shown while editing text. */
const FRAME_GRAB_SIZE = 10;

const FRAME_GRAB_INSET = FRAME_GRAB_SIZE / 2;

/** Top, bottom, left, and right frame grab hitboxes straddling the selection border. */
const FRAME_GRAB_HITBOXES: React.CSSProperties[] = [
  {
    left: -FRAME_GRAB_INSET,
    right: -FRAME_GRAB_INSET,
    top: -FRAME_GRAB_INSET,
    height: FRAME_GRAB_SIZE,
  },
  {
    left: -FRAME_GRAB_INSET,
    right: -FRAME_GRAB_INSET,
    bottom: -FRAME_GRAB_INSET,
    height: FRAME_GRAB_SIZE,
  },
  {
    left: -FRAME_GRAB_INSET,
    top: FRAME_GRAB_INSET,
    bottom: FRAME_GRAB_INSET,
    width: FRAME_GRAB_SIZE,
  },
  {
    right: -FRAME_GRAB_INSET,
    top: FRAME_GRAB_INSET,
    bottom: FRAME_GRAB_INSET,
    width: FRAME_GRAB_SIZE,
  },
];

/** Minimum shape size (slide px) a resize can shrink to. */
export const MIN_SIZE = 8;

/** Screen-px movement before a pointer-down becomes a drag instead of a click. */
export const DRAG_THRESHOLD = 3;

/**
 * Whether a press on empty canvas has travelled far enough to be a rubber band
 * rather than a click. Below the threshold nothing about the selection has
 * changed yet, so the band stays invisible and the handles stay up.
 */
export function isBandDrag(startX: number, startY: number, curX: number, curY: number): boolean {
  return Math.hypot(curX - startX, curY - startY) > DRAG_THRESHOLD;
}

const GRIP_DIRECTIONS = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;
export type GripDirection = (typeof GRIP_DIRECTIONS)[number];

const GRIP_CURSORS: Record<GripDirection, string> = {
  nw: "nwse-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  e: "ew-resize",
  se: "nwse-resize",
  s: "ns-resize",
  sw: "nesw-resize",
  w: "ew-resize",
};

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

const CORNER_GRIPS: ReadonlySet<GripDirection> = new Set(["nw", "ne", "se", "sw"]);

function debugLog(...args: unknown[]): void {
  if (!ENABLE_DEBUG_LOG) return;
  console.debug("[pptx-selection]", ...args);
}

/**
 * Modifiers PowerPoint treats as "extend the selection" rather than "start a
 * new one", for both clicks and rubber-band drags.
 */
function isMultiSelectEvent(event: React.PointerEvent): boolean {
  return event.shiftKey || event.ctrlKey || event.metaKey;
}

function getIsNativeUndoTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  return (
    target.isContentEditable ||
    target.tagName === "INPUT" ||
    target.tagName === "TEXTAREA" ||
    target.tagName === "SELECT"
  );
}

/**
 * Apply a resize drag (slide-px deltas) to the original rect, clamped to
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
 * Apply a move drag (slide-px deltas). With `lockAxis` (Shift held) the
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
 * Scale a rect by the factors the dragged shape was scaled by, anchored at the
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
      // Axis-aligned bounds of the rotated rect (rotation is about center).
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

/**
 * Strip zero-width spaces: line-height spacer spans from the renderer must
 * not reach the model.
 */
export function cleanText(raw: string | null | undefined): string {
  return (raw ?? "").replace(/\u200B/g, "");
}

function textContainerOf(shapeElement: HTMLElement): HTMLElement | null {
  // Empty placeholders also render a prompt overlay ("Click to add text")
  // whose paragraphs carry the paragraph attribute; skip it: only the real
  // text container is editable.
  for (const paragraphElement of Array.from(
    shapeElement.querySelectorAll<HTMLElement>(`[${PPTX_ATTRS.paragraph}]`),
  )) {
    if (paragraphElement.closest(`[${PPTX_ATTRS.placeholderPrompt}]`)) continue;
    return paragraphElement.parentElement;
  }
  return null;
}

function placeholderPromptOf(shapeElement: HTMLElement): HTMLElement | null {
  return shapeElement.querySelector<HTMLElement>(`[${PPTX_ATTRS.placeholderPrompt}]`);
}

/**
 * Walk the edited contentEditable container and produce a `setTextBody`
 * paragraphs payload. Uses the paragraph/run data attributes (see
 * `PPTX_ATTRS`) to map back to source indices for style inheritance.
 */
export function readBackTextBody(container: HTMLElement): SetTextBodyParagraph[] {
  const paragraphs: SetTextBodyParagraph[] = [];

  const children = Array.from(container.children).filter(
    (element) => element instanceof HTMLElement,
  ) as HTMLElement[];

  const paragraphDivs = children.filter(
    (element) => element.dataset[PPTX_DATASET.run] === undefined,
  );
  const effectiveDivs =
    paragraphDivs.length > 0 && paragraphDivs.length === children.length
      ? paragraphDivs
      : [container];
  let lastSourceParagraph = 0;

  for (const paragraphDiv of effectiveDivs) {
    const paragraphIndexAttr = paragraphDiv.dataset?.[PPTX_DATASET.paragraph];
    const sourceParagraphIndex =
      paragraphIndexAttr !== undefined ? Number(paragraphIndexAttr) : lastSourceParagraph;
    lastSourceParagraph = sourceParagraphIndex;

    const runs = readRunsFromParagraphElement(paragraphDiv, sourceParagraphIndex);
    paragraphs.push({ sourceParagraphIndex, runs });
  }

  return paragraphs;
}

function readRunsFromParagraphElement(
  paragraphDiv: HTMLElement,
  defaultSourceParagraph: number,
): SetTextBodyParagraph["runs"] {
  const runs: SetTextBodyParagraph["runs"] = [];
  let lastSourceRun: [number, number] | undefined;

  for (const child of Array.from(paragraphDiv.childNodes)) {
    if (child instanceof HTMLElement && child.dataset[PPTX_DATASET.bullet] !== undefined) {
      continue;
    }

    if (child instanceof HTMLElement && child.dataset[PPTX_DATASET.run] !== undefined) {
      const runIndex = Number(child.dataset[PPTX_DATASET.run]);
      const sourceRun: [number, number] = [defaultSourceParagraph, runIndex];
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun });
        lastSourceRun = sourceRun;
      }
    } else if (child instanceof HTMLBRElement) {
      // Browsers insert <br> for empty paragraphs; skip.
    } else if (child.nodeType === Node.TEXT_NODE) {
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceRun });
      }
    } else if (child instanceof HTMLElement) {
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceRun });
      }
    }
  }

  if (runs.length === 0) {
    runs.push({ text: "", sourceRun: lastSourceRun });
  }

  return runs;
}

/**
 * Compare the read-back paragraphs to the model to detect changes.
 */
export function textBodyChanged(node: SlideNode, readBack: SetTextBodyParagraph[]): boolean {
  if (node.nodeType !== "shape") return false;
  const shape = node as ShapeNodeData;
  const paragraphs = shape.textBody?.paragraphs;
  if (!paragraphs) return false;
  if (paragraphs.length !== readBack.length) return true;
  for (let i = 0; i < paragraphs.length; i++) {
    const origText = paragraphs[i].runs.map((run) => run.text ?? "").join("");
    const newText = readBack[i].runs.map((run) => run.text).join("");
    if (origText !== newText) return true;
  }
  return false;
}

type InternalState =
  | { mode: "idle" }
  | { mode: "selected"; nodeIds: string[] }
  | {
      mode: "move";
      /** All nodes moving together (multi-selection drags as a unit). */
      nodeIds: string[];
      /** The node under the pointer; drives click (no-drag) behavior. */
      primaryId: string;
      startX: number;
      startY: number;
      dx: number;
      dy: number;
      /** True once movement exceeded the drag threshold; gates the commit. */
      moved: boolean;
      /** Move started from text-mode border drag: return to text mode after. */
      resumeText?: boolean;
      /**
       * Selection to apply if this gesture turns out to be a click rather than
       * a drag. Set only for Shift/Ctrl presses, where a click toggles the
       * shape instead of selecting it or entering text.
       */
      clickIds?: string[];
    }
  | {
      /** Rubber-band selection from a drag on empty canvas (client coords). */
      mode: "marquee";
      startX: number;
      startY: number;
      curX: number;
      curY: number;
      /**
       * Selection the band extends, captured at pointer-down when Shift/Ctrl
       * was held. Empty for a plain band, which replaces the selection.
       */
      baseIds: string[];
    }
  | {
      mode: "resize";
      /** Every node being scaled; one entry for a solo selection. */
      nodeIds: string[];
      /** The node whose handle is being dragged; it sets the scale factors. */
      gripNodeId: string;
      grip: GripDirection;
      startX: number;
      startY: number;
      dx: number;
      dy: number;
      /** Shift held during the drag: corner grips keep the aspect ratio. */
      lockAspect: boolean;
    }
  | {
      mode: "text";
      nodeId: string;
      /** The text container element that has contentEditable. */
      editingElement: HTMLElement;
    };

/** The current interaction state of the edit layer. */
export interface SelectionState {
  /** Interaction mode of the selection. */
  mode: "idle" | "selected" | "move" | "resize" | "text" | "marquee";
  /**
   * The slide node currently selected, or `null` when nothing is selected.
   * With a multi-selection this is the first selected node.
   */
  selectedNode: SlideNode | null;
  /** All selected slide nodes (empty when nothing is selected). */
  selectedNodes: SlideNode[];
}

export interface SelectionProps extends React.ComponentProps<"div"> {
  /**
   * Replace the root overlay element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<SelectionState>;
  /** Called after `Ctrl+Z` fires. */
  onUndo?: (status: "success" | "empty", error?: unknown) => void;
  /** Called after `Ctrl+Shift+Z` / `Ctrl+Y` fires. */
  onRedo?: (status: "success" | "empty", error?: unknown) => void;
  /** Called after a node delete is attempted (keyboard or pointer). */
  onNodeDelete?: (nodeId: string, error?: unknown) => void;
  /** Called after a node move or resize is attempted. */
  onNodeTransform?: (nodeId: string, error?: unknown) => void;
  /** Called after inline text editing is committed (or errors). */
  onTextChange?: (nodeId: string, error?: unknown) => void;
}

/**
 * PowerPoint-style shape selection and manipulation overlay.
 *
 * Place inside `<Presentation.Slide>`. Renders nothing unless the loaded
 * presentation was opened with `readOnly={false}`.
 *
 * Interaction model mirrors PowerPoint:
 * - Click inside a text shape → edit text (caret appears).
 * - Drag any shape → move it (a multi-selection moves as a unit).
 * - Escape from text editing → select the shape (handles appear).
 * - Escape from selection → deselect.
 * - Click a non-text shape → select it.
 * - Shift/Ctrl+click toggles shapes in and out of the selection.
 * - Shift+drag a shape → move it along one axis only.
 * - Ctrl+drag a shape duplicates it in PowerPoint; unsupported here, so the
 *   press only toggles the selection.
 * - Drag on empty canvas → marquee-select fully enclosed shapes.
 * - Shift/Ctrl+drag on empty canvas → add the enclosed shapes to the selection.
 * - Every selected shape keeps its own handles, and dragging any one of them
 *   scales the whole selection, each shape in place.
 * - Ctrl/Cmd+A selects every shape on the slide.
 * - Arrow keys nudge, Delete removes: applied to the whole selection.
 */
const SelectionImpl = React.forwardRef<HTMLDivElement, SelectionProps>(function SelectionImpl(
  { render, onNodeDelete, onNodeTransform, onTextChange, ...selectionProps },
  forwardedRef,
) {
  const store = useStoreContext(SELECTION_NAME);
  const { presentation } = usePresentation();
  const { slide, slideId } = useSlide();
  const { zoom } = useZoom();

  const rootRef = React.useRef<HTMLDivElement>(null);
  const [state, setState] = React.useState<InternalState>({ mode: "idle" });

  // Stable refs for document-level listeners (avoids stale closures).
  const stateRef = useLatestRef(state);

  // Shallow clone of a styled run span, captured per shape on entering text
  // mode. When the browser destroys all run spans (select-all + delete),
  // typed text is re-wrapped with this template so it keeps the run's
  // styling.
  const runTemplateRef = React.useRef<{ nodeId: string; span: HTMLElement } | null>(null);

  // Edit revision of the active slide. A bump means SlideImpl will replace
  // the slide DOM in its effect; the text-mode repair effect below uses this
  // to re-attach contentEditable to the fresh DOM.
  const slideRevision = useSlideRevision(store, slideId);

  const selectedIds: string[] =
    state.mode === "selected" || state.mode === "move" || state.mode === "resize"
      ? state.nodeIds
      : state.mode === "text"
        ? [state.nodeId]
        : // An additive band keeps its base selection outlined while dragging,
          // so it stays visible what the band is adding to.
          state.mode === "marquee"
          ? state.baseIds
          : [];
  const selectedNodes: SlideNode[] = slide
    ? selectedIds
        .map((id) => slide.nodes.find((node) => node.id === id))
        .filter((node): node is SlideNode => node !== undefined)
    : [];
  /** The single selected node; with a multi-selection, the first one. */
  const selectedNode = selectedNodes[0] ?? null;
  const isSoloSelection = selectedNodes.length === 1;

  const isTextMode = state.mode === "text";
  const publicState: SelectionState = { mode: state.mode, selectedNode, selectedNodes };

  const isBanding =
    state.mode === "marquee" && isBandDrag(state.startX, state.startY, state.curX, state.curY);

  // Every selected shape previews the scale taken from the dragged shape.
  const grippedNode =
    state.mode === "resize" ? selectedNodes.find((node) => node.id === state.gripNodeId) : null;
  const resizeScale =
    state.mode === "resize" && grippedNode
      ? gripScale(getNodeRect(grippedNode), state.grip, state.dx, state.dy, state.lockAspect)
      : null;

  // Pasteboard hit area: shapes dragged past the slide edge render unclipped
  // in edit mode, but a surface sized to the slide (`inset: 0`) would never
  // receive pointerdowns over their off-slide portion. Extend the surface to
  // cover every node's bounds so those shapes stay clickable and draggable.
  //
  // Memoized because this component re-renders on every pointermove during
  // drags. Nodes are mutated in place, so the node array is not a usable
  // dependency; `slideRevision` bumps on every committed geometry edit and is
  // the correct cache key.
  const pasteboard = React.useMemo(
    () =>
      slide && presentation
        ? getPasteboardOverhang(slide.nodes, presentation.width, presentation.height)
        : { left: 0, top: 0, right: 0, bottom: 0 },
    // eslint-disable-next-line react-hooks/exhaustive-deps -- nodes mutate in place; revision is the change signal
    [slide, slideRevision, presentation],
  );

  function getSlideWrapper(): HTMLElement | null {
    return rootRef.current?.parentElement?.parentElement ?? null;
  }

  function getShapeElement(nodeId: string): HTMLElement | null {
    return (
      getSlideWrapper()?.querySelector<HTMLElement>(
        `[${PPTX_ATTRS.nodeId}="${CSS.escape(nodeId)}"]`,
      ) ?? null
    );
  }

  /**
   * Drop any document text selection left behind by earlier gestures.
   * A stale selection under the pointer turns the next press-and-drag into a
   * native drag-and-drop, cancelling our pointer stream mid-resize/move.
   */
  function clearDocumentSelection(): void {
    const selection = document.getSelection();
    if (selection && !selection.isCollapsed) selection.removeAllRanges();
  }

  /** Convert viewport (client) coordinates to slide-space px. */
  function clientToSlide(clientX: number, clientY: number): NodePosition {
    const wrapperRect = getSlideWrapper()?.getBoundingClientRect();
    if (!wrapperRect) return { x: 0, y: 0 };
    return { x: (clientX - wrapperRect.left) / zoom, y: (clientY - wrapperRect.top) / zoom };
  }

  function hitTest(clientX: number, clientY: number): string | null {
    const root = rootRef.current;
    const wrapper = getSlideWrapper();
    if (!root || !wrapper) return null;
    for (const element of document.elementsFromPoint(clientX, clientY)) {
      if (root.contains(element)) continue;
      if (!wrapper.contains(element)) continue;
      const nodeElement = (element as HTMLElement).closest<HTMLElement>(`[${PPTX_ATTRS.nodeId}]`);
      if (nodeElement) return nodeElement.getAttribute(PPTX_ATTRS.nodeId);
    }
    return null;
  }

  function nodeHasText(nodeId: string): boolean {
    const node = slide!.nodes.find((node) => node.id === nodeId);
    return node?.nodeType === "shape" && Boolean((node as ShapeNodeData).textBody);
  }

  /**
   * PowerPoint: a single click starts text editing only on dedicated text
   * boxes and placeholders. Regular shapes with text select on click and
   * need a double click (or typing) to edit.
   */
  function nodeEditsOnClick(nodeId: string): boolean {
    const node = slide!.nodes.find((node) => node.id === nodeId);
    if (node?.nodeType !== "shape") return false;
    const shape = node as ShapeNodeData;
    return Boolean(shape.textBody) && (Boolean(shape.isTextBox) || Boolean(shape.placeholder));
  }

  function commitEdit(
    action: () => Promise<unknown>,
    onRollback?: () => void,
    onSuccess?: () => void,
    onFailure?: (error: unknown) => void,
  ): void {
    action().then(onSuccess, (err) => {
      console.warn("[pptx] edit failed:", err);
      onRollback?.();
      onFailure?.(err);
    });
  }

  function enterTextMode(
    nodeId: string,
    clientX?: number,
    clientY?: number,
    insertText?: string,
  ): void {
    const shapeElement = getShapeElement(nodeId);
    const textElement = shapeElement ? textContainerOf(shapeElement) : null;
    debugLog("enterTextMode", {
      nodeId,
      shapeElementementFound: Boolean(shapeElement),
      textElementFound: Boolean(textElement),
      textElementConnected: textElement?.isConnected,
      textElementHtml: textElement?.outerHTML.slice(0, 200),
    });
    if (!textElement) {
      // No editable text container (e.g. decorative shape): fall back to
      // selection instead of leaving the previous mode (move) active.
      setState({ mode: "selected", nodeIds: [nodeId] });
      return;
    }

    textElement.contentEditable = "plaintext-only";
    if (!textElement.isContentEditable) {
      textElement.contentEditable = "true";
    }
    textElement.style.cursor = "text";
    textElement.style.outline = "none";

    // Capture a styling template before the browser can mutate the DOM.
    // When the DOM has no run spans left (e.g. re-entering a box that was
    // cleared without a re-render), keep the template captured earlier for
    // this shape instead of discarding it.
    const templateSource = textElement.querySelector<HTMLElement>(`[${PPTX_ATTRS.run}]`);
    if (templateSource) {
      runTemplateRef.current = {
        nodeId,
        span: templateSource.cloneNode(false) as HTMLElement,
      };
    } else if (runTemplateRef.current?.nodeId !== nodeId) {
      runTemplateRef.current = null;
    }

    // Hide the placeholder prompt overlay ("Click to add text") while
    // editing; it paints above the real text container and would cover
    // freshly typed text. The commit re-render restores or drops it.
    const placeholderPrompt = shapeElement ? placeholderPromptOf(shapeElement) : null;
    if (placeholderPrompt) placeholderPrompt.style.display = "none";

    // The caret-from-point APIs hit-test the DOM, and the overlay still
    // covers the text at this moment (React hasn't re-rendered with
    // pointerEvents: none yet). Disable it directly so the point resolves
    // into the text; the state render below keeps it disabled.
    if (rootRef.current) rootRef.current.style.pointerEvents = "none";

    textElement.focus({ preventScroll: true });

    if (clientX !== undefined && clientY !== undefined) {
      placeCaretAtPoint(clientX, clientY);
    } else {
      placeCaretAtEnd(textElement);
    }

    // PowerPoint: typing while a shape is selected starts editing with that
    // keystroke as the first character.
    if (insertText) {
      insertTextAtCaret(textElement, nodeId, insertText);
    }

    const selection = window.getSelection();
    debugLog("enterTextMode done", {
      activeElement: document.activeElement?.tagName,
      activeIsTextElement: document.activeElement === textElement,
      selectionAnchorInTextElement: selection?.anchorNode
        ? textElement.contains(selection.anchorNode)
        : null,
      selectionAnchor: selection?.anchorNode?.nodeName,
      isContentEditable: textElement.isContentEditable,
    });

    setState({ mode: "text", nodeId, editingElement: textElement });
  }

  /**
   * Re-enter text mode after a commit re-rendered the slide. The store
   * notify flushes React state at microtask end and the new slide DOM is
   * committed before the next paint, so a double rAF (not a timeout) is the
   * earliest reliable moment the replacement shape element exists.
   */
  function resumeTextEditing(nodeId: string): void {
    debugLog("resumeTextEditing scheduled", { nodeId });
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        debugLog("resumeTextEditing firing", { nodeId });
        enterTextMode(nodeId);
      });
    });
  }

  /**
   * Place the caret inside the last styled run span of `scope`. Typed text
   * must land inside a run span (`PPTX_ATTRS.run`) to inherit the run's font and
   * color: a bare text node at the container/paragraph level renders with
   * unstyled defaults (e.g. black text on a dark slide → invisible typing).
   * Returns false when the scope has no run spans.
   */
  function snapCaretIntoRun(scope: HTMLElement): boolean {
    const runs = scope.querySelectorAll<HTMLElement>(`[${PPTX_ATTRS.run}]`);
    const last = runs[runs.length - 1];
    if (!last) return false;
    const selection = window.getSelection();
    if (!selection) return false;

    let target = last.lastChild;
    if (!target || target.nodeType !== Node.TEXT_NODE) {
      target = document.createTextNode("");
      last.appendChild(target);
    }
    // Note: Chrome cannot keep the caret inside an empty text node; the
    // first keystroke would escape into the parent div and lose the run's
    // styling. The beforeinput interceptor (onDocBeforeInput) handles that
    // case by inserting the typed text into the span itself.
    const range = document.createRange();
    range.setStart(target, (target.textContent ?? "").length);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
    return true;
  }

  /** Move the caret to the given text node offset. */
  function setCaret(node: Node, offset: number): void {
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  }

  /**
   * Empty paragraphs are rendered with a placeholder `<br>` to preserve their
   * line height; once real text is inserted the `<br>` would push it onto a
   * second line. Remove it, but only when the paragraph holds nothing except
   * the just-inserted text; `<br>`s in paragraphs with other content are
   * genuine line breaks.
   */
  function removePlaceholderBreak(from: Node, insertedText: string): void {
    const element = from instanceof Element ? (from as HTMLElement) : from.parentElement;
    const paragraphElement = element?.closest<HTMLElement>(`[${PPTX_ATTRS.paragraph}]`);
    if (!paragraphElement) return;
    if (cleanText(paragraphElement.textContent) !== cleanText(insertedText)) return;
    for (const child of Array.from(paragraphElement.children)) {
      if (child instanceof HTMLBRElement) child.remove();
    }
  }

  /**
   * Intercept text insertion when the browser would drop it outside a styled
   * run span. Chrome cannot keep the caret inside an empty text node (the
   * keystroke escapes into the parent div), and destructive edits like
   * select-all + delete remove the run spans entirely; in both cases typed
   * text would render with unstyled defaults (e.g. near-white → invisible).
   * Instead of the default insertion we place the text into the run span at
   * the caret, or into a clone of the span captured on entering text mode.
   */
  /** The styling template span for a shape, if one was captured for it. */
  function runTemplateFor(nodeId: string): HTMLElement | null {
    const entry = runTemplateRef.current;
    return entry && entry.nodeId === nodeId ? entry.span : null;
  }

  /**
   * Insert `data` at the current caret, keeping it inside a styled run span
   * (existing span at the caret, or a clone of the shape's template span).
   * Returns false when there is nowhere sensible to insert.
   */
  function insertTextAtCaret(editingElement: HTMLElement, nodeId: string, data: string): boolean {
    const selection = window.getSelection();
    const anchor = selection?.anchorNode;
    if (!anchor || !selection?.isCollapsed || !editingElement.contains(anchor)) return false;

    const anchorElement =
      anchor instanceof Element ? (anchor as HTMLElement) : anchor.parentElement;
    const runSpan = anchorElement?.closest<HTMLElement>(`[${PPTX_ATTRS.run}]`);

    if (runSpan) {
      if (anchor.nodeType === Node.TEXT_NODE && (anchor.textContent ?? "").length > 0) {
        // Insert into the existing text node at the caret offset.
        const textNode = anchor as Text;
        const offset = Math.min(selection.anchorOffset, textNode.length);
        textNode.insertData(offset, data);
        setCaret(textNode, offset + data.length);
      } else {
        // Empty span: append a text node so the text stays inside.
        let textNode = runSpan.lastChild;
        if (!textNode || textNode.nodeType !== Node.TEXT_NODE) {
          textNode = document.createTextNode("");
          runSpan.appendChild(textNode);
        }
        textNode.textContent = (textNode.textContent ?? "") + data;
        setCaret(textNode, (textNode.textContent ?? "").length);
      }
      removePlaceholderBreak(runSpan, data);
      debugLog("inserted into run span", { data });
      return true;
    }

    // Caret is on a bare div (run spans destroyed, or the shape never had
    // any, e.g. a plain rectangle you just started typing into: wrap the
    // insertion in a clone of the template span when one was captured, else
    // a plain span that inherits the paragraph's styling. Read-back emits it
    // without a sourceRun and the edit falls back to paragraph defaults.
    const template = runTemplateFor(nodeId);
    const span = template
      ? (template.cloneNode(false) as HTMLElement)
      : document.createElement("span");
    const textNode = document.createTextNode(data);
    span.appendChild(textNode);
    if (anchorElement?.closest(`[${PPTX_ATTRS.paragraph}]`)) {
      selection.getRangeAt(0).insertNode(span);
    } else {
      // Caret sits at the container level (e.g. caret fallback when all run
      // spans were destroyed): inserting there would land the span *below*
      // the paragraph div, on its own line. Append into the last paragraph
      // instead.
      const paras = editingElement.querySelectorAll<HTMLElement>(`[${PPTX_ATTRS.paragraph}]`);
      const lastPara = paras[paras.length - 1];
      if (lastPara) lastPara.appendChild(span);
      else selection.getRangeAt(0).insertNode(span);
    }
    removePlaceholderBreak(span, data);
    setCaret(textNode, textNode.length);
    debugLog("rewrapped typing in run span", { spanHtml: span.outerHTML.slice(0, 200) });
    return true;
  }

  function interceptTextInsertion(
    event: InputEvent,
    editingElement: HTMLElement,
    nodeId: string,
  ): void {
    if (event.inputType !== "insertText" || !event.data) return;
    const selection = window.getSelection();
    const anchor = selection?.anchorNode;
    if (!anchor || !selection.isCollapsed || !editingElement.contains(anchor)) return;

    const anchorElement =
      anchor instanceof Element ? (anchor as HTMLElement) : anchor.parentElement;

    // Caret in a non-empty text node inside a styled element (run span,
    // hyperlink run, or a fallback span from a previous insertion): the
    // default insertion behaves correctly, don't interfere.
    if (
      anchor.nodeType === Node.TEXT_NODE &&
      (anchor.textContent ?? "").length > 0 &&
      anchorElement &&
      (anchorElement.closest(`[${PPTX_ATTRS.run}]`) || anchorElement.tagName === "SPAN")
    ) {
      return;
    }

    if (insertTextAtCaret(editingElement, nodeId, event.data)) {
      event.preventDefault();
    }
  }

  /**
   * Post-input safety net: if typed text still ended up as a bare text node
   * outside any run span (paths not covered by beforeinput, e.g. IME
   * composition or paste), reparent it into a clone of the template span.
   */
  function repairRunStyling(editingElement: HTMLElement, nodeId: string): void {
    const selection = window.getSelection();
    const anchor = selection?.anchorNode;
    if (!anchor || !selection?.isCollapsed || !editingElement.contains(anchor)) return;
    if (anchor.nodeType !== Node.TEXT_NODE) return;
    const anchorElement = anchor.parentElement;
    if (!anchorElement || anchorElement.closest(`[${PPTX_ATTRS.run}]`)) return;

    const template = runTemplateFor(nodeId);
    if (!template) return;

    const span = template.cloneNode(false) as HTMLElement;
    const offset = selection.anchorOffset;
    // Reparent the bare text node into a styled span; the node identity is
    // preserved so the caret offset stays valid.
    anchor.parentNode?.insertBefore(span, anchor);
    span.appendChild(anchor);
    removePlaceholderBreak(span, anchor.textContent ?? "");
    setCaret(anchor, Math.min(offset, (anchor.textContent ?? "").length));
    debugLog("repaired run styling at caret", { spanHtml: span.outerHTML.slice(0, 200) });
  }

  function placeCaretAtEnd(textElement: HTMLElement): void {
    try {
      if (snapCaretIntoRun(textElement)) return;
      const selection = window.getSelection();
      if (!selection) return;
      // No run spans: prefer the end of the last paragraph div over the
      // container itself so typing lands on the paragraph's line rather
      // than a new line below it.
      const paras = textElement.querySelectorAll<HTMLElement>(`[${PPTX_ATTRS.paragraph}]`);
      const target = paras[paras.length - 1] ?? textElement;
      const range = document.createRange();
      range.selectNodeContents(target);
      range.collapse(false);
      selection.removeAllRanges();
      selection.addRange(range);
    } catch {
      // Ignore caret placement failures.
    }
  }

  function placeCaretAtPoint(clientX: number, clientY: number): void {
    try {
      const selection = window.getSelection();
      if (!selection) return;
      const position = document.caretPositionFromPoint(clientX, clientY);
      if (position) {
        const range = document.createRange();
        range.setStart(position.offsetNode, position.offset);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
        fixupCaretAnchor(selection);
      }
    } catch {
      // Ignore caret placement failures.
    }
  }

  /**
   * If a point-placed caret anchored on an element (not a text node inside a
   * run span), snap it into the nearest run so typing inherits run styling.
   */
  function fixupCaretAnchor(selection: Selection): void {
    const anchor = selection.anchorNode;
    if (!(anchor instanceof HTMLElement)) return;
    // Prefer a run in the paragraph under the caret, else anywhere in scope.
    const paragraphDiv = anchor.closest<HTMLElement>(`[${PPTX_ATTRS.paragraph}]`);
    snapCaretIntoRun(paragraphDiv ?? anchor);
  }

  /** Tear down contentEditable and commit the edited text if it changed. */
  function commitTextEdits(current: Extract<InternalState, { mode: "text" }>): void {
    const { nodeId, editingElement } = current;

    editingElement.contentEditable = "false";
    editingElement.style.cursor = "";

    // Un-hide the placeholder prompt hidden by enterTextMode. When the edit
    // commits, the re-render rebuilds the shape anyway; when nothing changed
    // (no re-render), the prompt must come back by hand.
    const shapeElement = getShapeElement(nodeId);
    const placeholderPrompt = shapeElement ? placeholderPromptOf(shapeElement) : null;
    if (placeholderPrompt) placeholderPrompt.style.display = "";

    const node = slide!.nodes.find((node) => node.id === nodeId);
    if (!node) return;
    const readBack = readBackTextBody(editingElement);
    debugLog("commitTextEdits", {
      nodeId,
      editingElementConnected: editingElement.isConnected,
      editingElementHtml: editingElement.innerHTML.slice(0, 200),
      readBack: JSON.stringify(readBack).slice(0, 300),
      changed: textBodyChanged(node, readBack),
    });
    if (!textBodyChanged(node, readBack)) return;

    commitEdit(
      () =>
        store.edit({
          type: "setTextBody",
          slideId: slideId!,
          nodeId,
          paragraphs: readBack,
        }),
      undefined,
      () => {
        // The commit re-renders the slide, which steals focus from the
        // overlay. Refocus so keyboard shortcuts (undo/redo) keep working.
        rootRef.current?.focus({ preventScroll: true });
        onTextChange?.(nodeId);
      },
      (error) => {
        rootRef.current?.focus({ preventScroll: true });
        onTextChange?.(nodeId, error);
      },
    );
  }

  function doExitTextMode(
    current: Extract<InternalState, { mode: "text" }>,
    nextNodeId?: string | null,
  ): void {
    commitTextEdits(current);

    // Transition to the appropriate next state.
    if (nextNodeId === null) {
      setState({ mode: "idle" });
    } else {
      setState({ mode: "selected", nodeIds: [nextNodeId ?? current.nodeId] });
    }
    rootRef.current?.focus({ preventScroll: true });
  }

  // When in text mode the overlay has pointerEvents: none, so we must listen
  // on the document for clicks-outside and Escape.
  React.useEffect(() => {
    if (!isTextMode) return;

    function onDocPointerDown(event: PointerEvent): void {
      const currentState = stateRef.current;
      if (currentState.mode !== "text") return;

      // Click inside the editing element → let contentEditable handle it.
      if (currentState.editingElement.contains(event.target as Node)) return;

      // Clicks on the overlay's own children (border move strips) are
      // handled by their own handlers.
      if (rootRef.current?.contains(event.target as Node)) return;

      // Click on another shape?
      const wrapper = rootRef.current?.parentElement?.parentElement;
      if (!wrapper) return;
      const target = event.target as HTMLElement;
      const nodeElement = target.closest?.(`[${PPTX_ATTRS.nodeId}]`) as HTMLElement | null;
      const hitNodeId =
        nodeElement && wrapper.contains(nodeElement)
          ? nodeElement.getAttribute(PPTX_ATTRS.nodeId)
          : null;

      if (hitNodeId && hitNodeId !== currentState.nodeId) {
        // Clicked a different shape: exit text and start interacting with it
        // in the same gesture (PowerPoint: you can immediately drag another
        // shape while editing). preventDefault stops the browser's default
        // mousedown focus change from blurring the overlay; without it,
        // typing right after this click would go nowhere.
        event.preventDefault();
        commitTextEdits(currentState);
        rootRef.current?.setPointerCapture(event.pointerId);
        rootRef.current?.focus({ preventScroll: true });
        setState({
          mode: "move",
          nodeIds: [hitNodeId],
          primaryId: hitNodeId,
          startX: event.clientX,
          startY: event.clientY,
          dx: 0,
          dy: 0,
          moved: false,
        });
      } else if (!hitNodeId) {
        // Clicked empty area: exit text, deselect. preventDefault keeps the
        // overlay focused so undo/redo shortcuts still work afterwards.
        event.preventDefault();
        doExitTextMode(currentState, null);
      }
      // Else clicked inside the same shape but outside the text container
      // (e.g. shape padding area); stay in text mode.
    }

    function onDocKeyDown(event: KeyboardEvent): void {
      const currentState = stateRef.current;
      if (currentState.mode !== "text") return;

      if (event.key === "Escape") {
        event.preventDefault();
        // Escape → select the shape (PowerPoint behavior).
        doExitTextMode(currentState);
        return;
      }
    }

    function onDocBeforeInput(event: Event): void {
      const currentState = stateRef.current;
      if (currentState.mode !== "text") return;
      interceptTextInsertion(event as InputEvent, currentState.editingElement, currentState.nodeId);
    }

    function onDocInput(event: Event): void {
      const currentState = stateRef.current;
      if (currentState.mode !== "text") return;
      repairRunStyling(currentState.editingElement, currentState.nodeId);
      const target = event.target as HTMLElement;
      const selection = window.getSelection();
      const anchorNode = selection?.anchorNode;
      const anchorElement =
        anchorNode instanceof Element
          ? (anchorNode as HTMLElement)
          : (anchorNode?.parentElement ?? null);
      const rect = anchorElement?.getBoundingClientRect();
      const computedStyle = anchorElement ? getComputedStyle(anchorElement) : null;
      const elementAtCaret =
        rect && rect.width + rect.height > 0
          ? document.elementFromPoint(rect.x + rect.width / 2, rect.y + rect.height / 2)
          : null;
      debugLog("input event", {
        typedContent: currentState.editingElement.textContent?.slice(0, 80),
        anchorTag: anchorElement?.tagName,
        anchorIsRunSpan: anchorElement?.dataset?.[PPTX_DATASET.run] !== undefined,
        anchorRect: rect
          ? `${Math.round(rect.x)},${Math.round(rect.y)} ${Math.round(rect.width)}x${Math.round(rect.height)}`
          : null,
        color: computedStyle?.color,
        fontSize: computedStyle?.fontSize,
        opacity: computedStyle?.opacity,
        visibility: computedStyle?.visibility,
        elementAtCaret: elementAtCaret
          ? `${elementAtCaret.tagName} ${(elementAtCaret as HTMLElement).dataset?.[PPTX_DATASET.nodeId] ?? ""}`
          : null,
        anchorHtml: anchorElement?.outerHTML.slice(0, 200),
        targetIsEditingElement: target === currentState.editingElement,
        editingElementConnected: currentState.editingElement.isConnected,
      });
    }

    document.addEventListener("pointerdown", onDocPointerDown, true);
    document.addEventListener("keydown", onDocKeyDown, true);
    document.addEventListener("beforeinput", onDocBeforeInput, true);
    document.addEventListener("input", onDocInput, true);
    return () => {
      document.removeEventListener("pointerdown", onDocPointerDown, true);
      document.removeEventListener("keydown", onDocKeyDown, true);
      document.removeEventListener("beforeinput", onDocBeforeInput, true);
      document.removeEventListener("input", onDocInput, true);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- stateRef is stable
  }, [isTextMode]);

  // A revision bump makes SlideImpl replace the slide DOM in its effect,
  // detaching our contentEditable element: typed text would go into the
  // detached tree and never appear on screen. SlideImpl is a parent, so its
  // effect runs after this one; the rAF fires after the whole effects flush,
  // when the fresh DOM is in place, and re-attaches editing to it.
  React.useEffect(() => {
    debugLog("revision effect", { slideRevision, mode: stateRef.current.mode });
    if (stateRef.current.mode !== "text") return;
    const frameId = requestAnimationFrame(() => {
      const currentState = stateRef.current;
      if (currentState.mode !== "text") return;
      debugLog("repair check", {
        editingElementConnected: currentState.editingElement.isConnected,
      });
      if (currentState.editingElement.isConnected) return;
      debugLog("repairing: re-entering text mode", { nodeId: currentState.nodeId });
      enterTextMode(currentState.nodeId);
    });
    return () => cancelAnimationFrame(frameId);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- repair keyed on revision only
  }, [slideRevision]);

  function onPointerDown(event: React.PointerEvent<HTMLDivElement>): void {
    if (event.button !== 0 || isTextMode) return;

    const nodeId = hitTest(event.clientX, event.clientY);
    const selectedShapeElement = selectedNode ? getShapeElement(selectedNode.id) : null;
    debugLog("root pointerdown", {
      hit: nodeId,
      mode: state.mode,
      x: event.clientX,
      y: event.clientY,
      target: (event.target as HTMLElement)?.tagName,
      targetHandle: (event.target as HTMLElement)?.dataset?.resizeHandle,
      selectedRect: selectedNode ? getNodeRect(selectedNode) : null,
      selectedDomRect: selectedShapeElement
        ? (() => {
            const bounds = selectedShapeElement.getBoundingClientRect();
            return `${Math.round(bounds.x)},${Math.round(bounds.y)} ${Math.round(bounds.width)}x${Math.round(bounds.height)}`;
          })()
        : null,
    });
    // Prevent the browser's default mousedown behavior. Without this, a drag
    // silently extends a text selection across the slide, and the NEXT press
    // inside that selection starts a native drag-and-drop of the selection;
    // the pointer stream is cancelled (not-allowed cursor, no pointerup) and
    // the gesture goes dead.
    event.preventDefault();
    clearDocumentSelection();
    rootRef.current?.setPointerCapture(event.pointerId);
    rootRef.current?.focus();

    if (!nodeId) {
      // Empty canvas: start a marquee (rubber-band) selection. A no-drag
      // click resolves to deselect on pointer-up. Shift/Ctrl keeps the
      // current selection so the band adds to it (PowerPoint).
      setState({
        mode: "marquee",
        startX: event.clientX,
        startY: event.clientY,
        curX: event.clientX,
        curY: event.clientY,
        baseIds: isMultiSelectEvent(event) ? selectedIds : [],
      });
      return;
    }

    // Ctrl/Cmd+click toggles the shape in and out of the selection, resolved
    // here rather than on pointer-up: PowerPoint reserves Ctrl+drag for
    // duplicating a shape, which needs a core edit operation that can add a
    // node (there is none yet), so this press must not become a move.
    if (event.ctrlKey || event.metaKey) {
      const nextIds = toggleSelection(selectedIds, nodeId);
      setState(nextIds.length > 0 ? { mode: "selected", nodeIds: nextIds } : { mode: "idle" });
      return;
    }

    // Pressing a shape that is already part of the selection drags the whole
    // selection; anything else starts a fresh single-shape gesture. On
    // pointer-up we check if the user actually dragged; if not (a click),
    // text shapes enter text mode and non-text shapes get selected.
    //
    // Shift leaves both Shift+click (toggle) and Shift+drag (move along one
    // axis) open, so the gesture starts as a move and pointer-up decides.
    // Resolving the toggle up front would consume the press and leave the drag
    // unable to move anything.
    const additive = event.shiftKey;
    setState({
      mode: "move",
      nodeIds: additive
        ? mergeSelection(selectedIds, [nodeId])
        : selectedIds.includes(nodeId)
          ? selectedIds
          : [nodeId],
      primaryId: nodeId,
      startX: event.clientX,
      startY: event.clientY,
      dx: 0,
      dy: 0,
      moved: false,
      // Precomputed against the selection as it stands now: once the gesture
      // starts, `nodeIds` has already absorbed this shape and can no longer
      // tell whether the click should add it or take it back out.
      clickIds: additive ? toggleSelection(selectedIds, nodeId) : undefined,
    });
  }

  function onGripPointerDown(
    event: React.PointerEvent<HTMLDivElement>,
    grip: GripDirection,
    nodeId: string,
  ): void {
    debugLog("grip pointerdown", { grip, nodeId, mode: state.mode, button: event.button });
    if (event.button !== 0 || !selectedIds.includes(nodeId) || isTextMode) return;
    event.stopPropagation();
    // See onPointerDown: stop selection extension / native drag initiation.
    event.preventDefault();
    clearDocumentSelection();
    rootRef.current?.setPointerCapture(event.pointerId);
    setState({
      mode: "resize",
      nodeIds: selectedIds,
      gripNodeId: nodeId,
      grip,
      startX: event.clientX,
      startY: event.clientY,
      dx: 0,
      dy: 0,
      lockAspect: event.shiftKey,
    });
  }

  /**
   * Dragging the border strips while editing text (PowerPoint: grab the
   * frame to move the box without leaving your edits behind). Commits the
   * text first, then starts a move drag.
   */
  function onBorderPointerDown(event: React.PointerEvent<HTMLDivElement>): void {
    if (event.button !== 0 || state.mode !== "text") return;
    event.preventDefault();
    event.stopPropagation();

    const { nodeId } = state;
    commitTextEdits(state);

    rootRef.current?.setPointerCapture(event.pointerId);
    rootRef.current?.focus({ preventScroll: true });
    setState({
      mode: "move",
      nodeIds: [nodeId],
      primaryId: nodeId,
      startX: event.clientX,
      startY: event.clientY,
      dx: 0,
      dy: 0,
      moved: false,
      resumeText: true,
    });
  }

  function onPointerMove(event: React.PointerEvent<HTMLDivElement>): void {
    // No button held: not a drag. Guard against stale-state moves that can
    // occur between setState({ mode: "selected" }) in onPointerUp and the
    // React re-render that creates a fresh closure.
    if (event.buttons === 0) return;

    if (state.mode === "marquee") {
      setState({ ...state, curX: event.clientX, curY: event.clientY });
      return;
    }

    if (state.mode !== "move" && state.mode !== "resize") return;
    const dx = (event.clientX - state.startX) / zoom;
    const dy = (event.clientY - state.startY) / zoom;

    if (state.mode === "move") {
      const moved =
        state.moved ||
        Math.hypot(event.clientX - state.startX, event.clientY - state.startY) > DRAG_THRESHOLD;
      // Track Shift live (as with resize) so holding or releasing it mid-drag
      // locks and unlocks the axis. The constrained deltas are what lands in
      // state, so pointer-up commits the straight move without re-deriving it.
      const delta = constrainMove(dx, dy, event.shiftKey);
      if (moved) {
        for (const node of selectedNodes) {
          const shapeElement = getShapeElement(node.id);
          if (shapeElement) {
            shapeElement.style.left = `${node.position.x + delta.x}px`;
            shapeElement.style.top = `${node.position.y + delta.y}px`;
          }
        }
      }
      setState({ ...state, dx: delta.x, dy: delta.y, moved });
    } else {
      // Track Shift live so holding/releasing it mid-drag toggles the lock.
      setState({ ...state, dx, dy, lockAspect: event.shiftKey });
    }
  }

  function onPointerUp(): void {
    if (state.mode === "marquee") {
      const { startX, startY, curX, curY, baseIds } = state;
      if (!isBandDrag(startX, startY, curX, curY)) {
        // Plain click on empty canvas → deselect. Held modifiers mean the user
        // was adding to a selection, so a stray click must not throw it away.
        setState(baseIds.length > 0 ? { mode: "selected", nodeIds: baseIds } : { mode: "idle" });
        return;
      }
      // PowerPoint selects shapes fully enclosed by the rubber band.
      const a = clientToSlide(startX, startY);
      const b = clientToSlide(curX, curY);
      const box: Rect = {
        x: Math.min(a.x, b.x),
        y: Math.min(a.y, b.y),
        w: Math.abs(a.x - b.x),
        h: Math.abs(a.y - b.y),
      };
      const contained = slide!.nodes
        .filter((node) => {
          const nodeRect = getNodeRect(node);
          return (
            nodeRect.x >= box.x &&
            nodeRect.y >= box.y &&
            nodeRect.x + nodeRect.w <= box.x + box.w &&
            nodeRect.y + nodeRect.h <= box.y + box.h
          );
        })
        .map((node) => node.id);
      const nextIds = mergeSelection(baseIds, contained);
      setState(nextIds.length > 0 ? { mode: "selected", nodeIds: nextIds } : { mode: "idle" });
    } else if (state.mode === "move") {
      const { nodeIds, primaryId, dx, dy, moved, resumeText, clickIds } = state;

      if (moved && selectedNodes.length > 0) {
        // Actual drag → commit the move (one undoable edit for the whole
        // selection), land in selected mode.
        const movingNodes = selectedNodes;
        setState({ mode: "selected", nodeIds });
        commitEdit(
          () =>
            store.edit(
              movingNodes.length === 1
                ? {
                    type: "setNodeTransform",
                    slideId: slideId!,
                    nodeId: movingNodes[0].id,
                    position: {
                      x: movingNodes[0].position.x + dx,
                      y: movingNodes[0].position.y + dy,
                    },
                  }
                : {
                    type: "batch",
                    operations: movingNodes.map((node) => ({
                      type: "setNodeTransform",
                      slideId: slideId!,
                      nodeId: node.id,
                      position: { x: node.position.x + dx, y: node.position.y + dy },
                    })),
                  },
            ),
          () => {
            for (const node of movingNodes) {
              const shapeElement = getShapeElement(node.id);
              if (shapeElement) {
                shapeElement.style.left = `${node.position.x}px`;
                shapeElement.style.top = `${node.position.y}px`;
              }
            }
          },
          () => {
            for (const id of nodeIds) onNodeTransform?.(id);
            if (resumeText) resumeTextEditing(primaryId);
          },
          (error) => {
            for (const id of nodeIds) onNodeTransform?.(id, error);
          },
        );
      } else {
        // Click (no drag). Reset any stray left/top offset.
        for (const node of selectedNodes) {
          const shapeElement = getShapeElement(node.id);
          if (shapeElement) {
            shapeElement.style.left = `${node.position.x}px`;
            shapeElement.style.top = `${node.position.y}px`;
          }
        }

        // Text boxes / placeholders → edit immediately. Everything else
        // (regular shapes, even with text) → select; they edit on
        // double-click or by typing. Clicking a member of a multi-selection
        // without dragging collapses the selection to it (PowerPoint).
        if (clickIds) {
          // Shift/Ctrl+click: toggle only, never open the caret.
          setState(
            clickIds.length > 0 ? { mode: "selected", nodeIds: clickIds } : { mode: "idle" },
          );
        } else if (resumeText) {
          // Border click without drag: back to editing (caret at end). The
          // text commit may re-render the slide, so wait for the new DOM.
          resumeTextEditing(primaryId);
        } else if (nodeEditsOnClick(primaryId)) {
          // Use the original pointer position for caret placement.
          enterTextMode(primaryId, state.startX, state.startY);
        } else {
          setState({ mode: "selected", nodeIds: [primaryId] });
        }
      }
    } else if (state.mode === "resize") {
      const { nodeIds, gripNodeId, grip, dx, dy, lockAspect } = state;
      const resizingNodes = selectedNodes;
      const grippedNode = resizingNodes.find((node) => node.id === gripNodeId) ?? null;
      setState({ mode: "selected", nodeIds });
      debugLog("resize pointerup", {
        nodeIds,
        gripNodeId,
        grip,
        dx,
        dy,
        lockAspect,
        nodeCount: resizingNodes.length,
        baseRect: grippedNode ? getNodeRect(grippedNode) : null,
      });
      if (grippedNode && (dx !== 0 || dy !== 0)) {
        // Dragging one shape's handle scales the whole selection by the factors
        // that shape was scaled by, each shape anchored in place (PowerPoint).
        const scale = gripScale(getNodeRect(grippedNode), grip, dx, dy, lockAspect);
        const transforms = resizingNodes.map((node) => ({
          nodeId: node.id,
          rect: scaleRectFromGrip(getNodeRect(node), grip, scale),
        }));
        commitEdit(
          () =>
            store.edit(
              transforms.length === 1
                ? {
                    type: "setNodeTransform",
                    slideId: slideId!,
                    nodeId: transforms[0].nodeId,
                    position: { x: transforms[0].rect.x, y: transforms[0].rect.y },
                    size: { w: transforms[0].rect.w, h: transforms[0].rect.h },
                  }
                : {
                    type: "batch",
                    operations: transforms.map(({ nodeId: id, rect }) => ({
                      type: "setNodeTransform",
                      slideId: slideId!,
                      nodeId: id,
                      position: { x: rect.x, y: rect.y },
                      size: { w: rect.w, h: rect.h },
                    })),
                  },
            ),
          undefined,
          () => {
            for (const id of nodeIds) onNodeTransform?.(id);
          },
          (error) => {
            for (const id of nodeIds) onNodeTransform?.(id, error);
          },
        );
      }
    }
  }

  /**
   * Double click on a regular shape with text → edit with caret at point.
   */
  function onDoubleClick(event: React.MouseEvent<HTMLDivElement>): void {
    debugLog("double click", { isTextMode });
    if (isTextMode) return;
    const nodeId = hitTest(event.clientX, event.clientY);
    if (nodeId && nodeHasText(nodeId)) {
      enterTextMode(nodeId, event.clientX, event.clientY);
    }
  }

  /**
   * Native drag-and-drop of slide content (images, stale text selections)
   * cancels our pointer gestures; block it while the overlay is interactive.
   */
  function onDragStart(event: React.DragEvent<HTMLDivElement>): void {
    debugLog("dragstart blocked", { target: (event.target as HTMLElement)?.tagName });
    event.preventDefault();
  }

  function onPointerCancel(): void {
    debugLog("pointercancel", { mode: state.mode });
    if (state.mode === "move") {
      for (const node of selectedNodes) {
        const shapeElement = getShapeElement(node.id);
        if (shapeElement) {
          shapeElement.style.left = `${node.position.x}px`;
          shapeElement.style.top = `${node.position.y}px`;
        }
      }
      setState({ mode: "selected", nodeIds: state.nodeIds });
    } else if (state.mode === "resize") {
      setState({ mode: "selected", nodeIds: state.nodeIds });
    } else if (state.mode === "marquee") {
      const { baseIds } = state;
      setState(baseIds.length > 0 ? { mode: "selected", nodeIds: baseIds } : { mode: "idle" });
    }
  }

  /**
   * Move every given node by `delta`, as a single undoable edit.
   */
  function nudge(nodes: SlideNode[], delta: NodePosition): void {
    if (nodes.length === 0) return;
    const ops = nodes.map((node) => ({
      type: "setNodeTransform" as const,
      slideId: slideId!,
      nodeId: node.id,
      position: { x: node.position.x + delta.x, y: node.position.y + delta.y },
    }));
    commitEdit(
      () => store.edit(ops.length === 1 ? ops[0] : { type: "batch", operations: ops }),
      undefined,
      () => {
        for (const node of nodes) onNodeTransform?.(node.id);
      },
      (error) => {
        for (const node of nodes) onNodeTransform?.(node.id, error);
      },
    );
  }

  function onKeyDown(event: React.KeyboardEvent<HTMLDivElement>): void {
    // Text mode keys are handled by the document listener; this handler
    // only fires when the overlay div has focus (selected/idle modes).
    // Undo/redo shortcuts live in the outer Selection wrapper (document
    // level) so they survive the slide-change remount of this component.
    if (isTextMode) return;

    const mod = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();

    // Ctrl/Cmd+A selects every shape on the slide (PowerPoint).
    if (mod && key === "a") {
      event.preventDefault();
      const allIds = slide!.nodes.map((n) => n.id);
      if (allIds.length > 0) setState({ mode: "selected", nodeIds: allIds });
      return;
    }

    if (selectedNodes.length === 0) return;

    // F2 or Enter on a single selected text shape enters text mode.
    if (
      (event.key === "F2" || event.key === "Enter") &&
      isSoloSelection &&
      nodeHasText(selectedNodes[0].id)
    ) {
      event.preventDefault();
      enterTextMode(selectedNodes[0].id);
      return;
    }

    // PowerPoint: typing while a single shape is selected starts text editing
    // with that keystroke as the first character (caret at end of existing text).
    if (
      event.key.length === 1 &&
      !mod &&
      !event.altKey &&
      isSoloSelection &&
      nodeHasText(selectedNodes[0].id)
    ) {
      event.preventDefault();
      enterTextMode(selectedNodes[0].id, undefined, undefined, event.key);
      return;
    }

    if (event.key === "Escape") {
      setState({ mode: "idle" });
      return;
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      const nodeIds = selectedNodes.map((n) => n.id);
      commitEdit(
        () =>
          store.edit(
            nodeIds.length === 1
              ? { type: "deleteNode", slideId: slideId!, nodeId: nodeIds[0] }
              : {
                  type: "batch",
                  operations: nodeIds.map((nodeId) => ({
                    type: "deleteNode" as const,
                    slideId: slideId!,
                    nodeId,
                  })),
                },
          ),
        undefined,
        () => {
          setState({ mode: "idle" });
          for (const nodeId of nodeIds) onNodeDelete?.(nodeId);
        },
        (error) => {
          for (const nodeId of nodeIds) onNodeDelete?.(nodeId, error);
        },
      );
      return;
    }

    const step = event.shiftKey ? 10 : 1;
    const arrows: Record<string, NodePosition> = {
      ArrowLeft: { x: -step, y: 0 },
      ArrowRight: { x: step, y: 0 },
      ArrowUp: { x: 0, y: -step },
      ArrowDown: { x: 0, y: step },
    };
    const delta = arrows[event.key];
    if (delta) {
      event.preventDefault();
      nudge(selectedNodes, delta);
    }
  }

  // Bail out below every hook, never above one: returning early before the
  // effects would change the hook count between renders, which React rejects.
  if (!presentation?.sourcePackage || !slide || !slideId) return null;

  return renderElement(
    "div",
    { render },
    {
      state: publicState,
      ref: mergeRefs(rootRef, forwardedRef),
      props: [
        {
          role: "application",
          "data-pptx-selection": "",
          "data-mode": state.mode,
          tabIndex: 0,
          onPointerDown,
          onPointerMove,
          onPointerUp,
          onPointerCancel,
          onDoubleClick,
          onKeyDown,
          onDragStart,
          children: (
            <>
              {/*
               * Slide-origin frame: the root extends into the pasteboard, so
               * selection boxes (positioned in slide coordinates) need a
               * positioning context whose origin is the slide's top-left.
               */}
              <div
                style={{
                  position: "absolute",
                  left: pasteboard.left * zoom,
                  top: pasteboard.top * zoom,
                  width: presentation.width * zoom,
                  height: presentation.height * zoom,
                  pointerEvents: "none",
                }}
              >
                {selectedNodes.map((node) => (
                  <SelectionBox
                    key={node.id}
                    node={node}
                    state={state}
                    zoom={zoom}
                    resizeScale={resizeScale}
                    // Grips are unreachable mid-band (the pointer is captured)
                    // and would sit under the rubber band, so only the outline
                    // of the base selection shows while it is being extended.
                    // Until the band starts, the press is still just a click and
                    // must not blink the handles of what is already selected.
                    showResizeGrips={!isBanding}
                    onGripPointerDown={onGripPointerDown}
                    onBorderPointerDown={onBorderPointerDown}
                  />
                ))}
              </div>
              {state.mode === "marquee" && (
                <MarqueeBox state={state} rootElement={rootRef.current} />
              )}
            </>
          ),
          style: {
            position: "absolute",
            // Negative insets extend the event surface over the pasteboard
            // wherever shapes overflow the slide (see `pasteboard` above).
            left: -pasteboard.left * zoom,
            top: -pasteboard.top * zoom,
            right: -pasteboard.right * zoom,
            bottom: -pasteboard.bottom * zoom,
            // In text mode let clicks through to the contentEditable element.
            pointerEvents: isTextMode ? "none" : "auto",
            outline: "none",
            touchAction: "none",
          },
        },
        selectionProps,
      ],
    },
  );
});

/**
 * Thin outer wrapper that keys the inner component by `slideId`. When the
 * slide changes, React unmounts and remounts `SelectionImpl`, naturally
 * resetting its local state (selection, text-editing refs). Shape IDs are
 * reused across slides, so a stale selection would otherwise bleed onto the
 * new slide; the key swap is the React-idiomatic way to scope local state
 * to the current context without effect-driven resets.
 *
 * Undo/redo shortcuts are handled here, at document level, rather than in
 * `SelectionImpl`: they are store operations that don't need selection state,
 * and they must survive the remount when an undo navigates to another slide.
 * PowerPoint likewise accepts Ctrl+Z regardless of what part of the app has
 * focus.
 */
export const Selection = React.forwardRef<HTMLDivElement, SelectionProps>(function Selection(
  { onUndo, onRedo, ...props },
  forwardedRef,
) {
  const store = useStoreContext(SELECTION_NAME);
  const { presentation } = usePresentation();
  const { slideId } = useSlide();

  const onUndoRef = useLatestRef(onUndo);
  const onRedoRef = useLatestRef(onRedo);

  const editable = presentation?.sourcePackage != null;

  React.useEffect(() => {
    if (!editable) return;

    function onDocKeyDown(event: KeyboardEvent): void {
      if (!(event.ctrlKey || event.metaKey)) return;
      if (getIsNativeUndoTarget(event.target)) return;
      const key = event.key.toLowerCase();

      if (key === "z" && !event.shiftKey) {
        event.preventDefault();
        const success = store.undo();
        onUndoRef.current?.(success ? "success" : "empty");
      } else if (key === "y" || (key === "z" && event.shiftKey)) {
        event.preventDefault();
        store.redo().then(
          (success) => onRedoRef.current?.(success ? "success" : "empty"),
          (error) => onRedoRef.current?.("empty", error),
        );
      }
    }

    document.addEventListener("keydown", onDocKeyDown);
    return () => document.removeEventListener("keydown", onDocKeyDown);
  }, [editable, store]);

  if (!slideId) return null;

  return <SelectionImpl key={slideId} ref={forwardedRef} {...props} />;
});

interface SelectionBoxProps {
  node: SlideNode;
  state: InternalState;
  zoom: number;
  /** Scale taken from the dragged shape, applied to every selected shape. */
  resizeScale: ResizeScale | null;
  /** False mid-band, when the grips are unreachable anyway. */
  showResizeGrips: boolean;
  onGripPointerDown: (
    event: React.PointerEvent<HTMLDivElement>,
    grip: GripDirection,
    nodeId: string,
  ) => void;
  onBorderPointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
}

function SelectionBox({
  node,
  state,
  zoom,
  resizeScale,
  showResizeGrips,
  onGripPointerDown,
  onBorderPointerDown,
}: SelectionBoxProps) {
  let rect = getNodeRect(node);
  if (state.mode === "move" && state.moved) {
    rect = { ...rect, x: rect.x + state.dx, y: rect.y + state.dy };
  } else if (state.mode === "resize" && resizeScale) {
    rect = scaleRectFromGrip(rect, state.grip, resizeScale);
  }

  const isTextMode = state.mode === "text";
  // Handles hide once a drag is actually under way, not merely on pointer-down:
  // a press that only toggles the selection (Shift+click) would otherwise blink
  // them off and back on.
  const isDragging = state.mode === "move" && state.moved;
  const showGrips = showResizeGrips && node.rotation === 0 && !isDragging && !isTextMode;

  return (
    <div
      data-selection-box=""
      style={{
        position: "absolute",
        left: rect.x * zoom,
        top: rect.y * zoom,
        width: rect.w * zoom,
        height: rect.h * zoom,
        transform: node.rotation !== 0 ? `rotate(${node.rotation}deg)` : undefined,
        boxShadow: `0 0 0 ${isTextMode ? "2" : "1.5"}px var(--presentation-selection, #2563eb)`,
        cursor: isTextMode ? "text" : "move",
        pointerEvents: "none",
      }}
    >
      {showGrips && (
        <ResizeGrips onGripPointerDown={(event, grip) => onGripPointerDown(event, grip, node.id)} />
      )}
      {isTextMode && <FrameGrabHitboxes onPointerDown={onBorderPointerDown} />}
    </div>
  );
}

const GRIP_POSITIONS: Record<GripDirection, React.CSSProperties> = {
  nw: { left: 0, top: 0 },
  n: { left: "50%", top: 0 },
  ne: { left: "100%", top: 0 },
  e: { left: "100%", top: "50%" },
  se: { left: "100%", top: "100%" },
  s: { left: "50%", top: "100%" },
  sw: { left: 0, top: "100%" },
  w: { left: 0, top: "50%" },
};

interface ResizeGripsProps {
  onGripPointerDown: (event: React.PointerEvent<HTMLDivElement>, grip: GripDirection) => void;
}

/** The eight handles, positioned against whichever box encloses them. */
function ResizeGrips({ onGripPointerDown }: ResizeGripsProps) {
  return GRIP_DIRECTIONS.map((direction) => (
    <div
      key={direction}
      aria-hidden="true"
      data-resize-grip={direction}
      onPointerDown={(event) => onGripPointerDown(event, direction)}
      style={{
        position: "absolute",
        ...GRIP_POSITIONS[direction],
        width: 9,
        height: 9,
        marginLeft: -4.5,
        marginTop: -4.5,
        background: "#fff",
        border: "1.5px solid var(--presentation-selection, #2563eb)",
        borderRadius: 2,
        cursor: GRIP_CURSORS[direction],
        pointerEvents: "auto",
        touchAction: "none",
      }}
    />
  ));
}

interface MarqueeBoxProps {
  state: Extract<InternalState, { mode: "marquee" }>;
  rootElement: HTMLElement | null;
}

/**
 * Rubber-band rectangle drawn while drag-selecting on empty canvas.
 */
function MarqueeBox({ state, rootElement }: MarqueeBoxProps) {
  // State coords are viewport-relative; the overlay is our positioning context.
  const origin = rootElement?.getBoundingClientRect();
  if (!origin) return null;
  const left = Math.min(state.startX, state.curX) - origin.left;
  const top = Math.min(state.startY, state.curY) - origin.top;
  const width = Math.abs(state.curX - state.startX);
  const height = Math.abs(state.curY - state.startY);
  if (!isBandDrag(state.startX, state.startY, state.curX, state.curY)) return null;

  return (
    <div
      data-marquee=""
      style={{
        position: "absolute",
        left,
        top,
        width,
        height,
        border: "1px solid var(--presentation-selection, #2563eb)",
        background: "color-mix(in srgb, var(--presentation-selection, #2563eb) 10%, transparent)",
        pointerEvents: "none",
      }}
    />
  );
}

interface FrameGrabHitboxesProps {
  onPointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
}

/**
 * Invisible grab hitboxes straddling the selection frame while editing text.
 * Dragging them moves the shape (PowerPoint: grab the frame to move).
 * Each hitbox extends both inward and outward from the border line.
 */
function FrameGrabHitboxes({ onPointerDown }: FrameGrabHitboxesProps) {
  return (
    <>
      {FRAME_GRAB_HITBOXES.map((style, index) => (
        <div
          key={index}
          aria-hidden="true"
          data-border-move=""
          onPointerDown={onPointerDown}
          style={{
            position: "absolute",
            ...style,
            cursor: "move",
            pointerEvents: "auto",
            touchAction: "none",
          }}
        />
      ))}
    </>
  );
}

export namespace Selection {
  export type State = SelectionState;
  export type Props = SelectionProps;
}
