import * as React from "react";

import type {
  EditOperation,
  NodePosition,
  ParagraphStyle,
  SetTextBodyParagraph,
  ShapeNodeData,
  SlideNode,
  TextRange,
  TextStyle,
} from "@diceui/pptx-core";
import { PPTX_ATTRS, PPTX_DATASET } from "@diceui/pptx-core";

import {
  RootContext,
  usePresentation,
  useSlide,
  useSlideRevision,
  useStoreContext,
  useZoom,
} from "./context";
import { useLatestRef } from "./hooks";
import type { PrimitiveProps, StatefulPrimitiveProps } from "./render";
import { mergeRefs, renderElement } from "./render";
import { getIsKeepSelectionTarget, getSelectionChannel } from "./selection-channel";
import type { SnapGuide as SnapGuideData } from "./snap";
import { getUnionRect, SNAP_THRESHOLD, snapMove, snapResize } from "./snap";
import type { ActiveTextStyle } from "./text-range";
import {
  cleanText,
  expandToWord,
  getActiveTextStyle,
  getDomPoint,
  getParagraphElements,
  getParagraphText,
  getRangeRuns,
  getTextRange,
} from "./text-range";

export { cleanText };

const SELECTION_NAME = "Presentation.Selection";

const ENABLE_DEBUG_LOG = false;

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

const CORNER_GRIPS: ReadonlySet<GripDirection> = new Set(["nw", "ne", "se", "sw"]);

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

function debugLog(...args: unknown[]): void {
  if (!ENABLE_DEBUG_LOG) return;
  console.debug("[pptx-selection]", ...args);
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
 * Whether a document-level keystroke belongs to the presentation.
 *
 * `boundary` is `Root`'s element when there is one, so a keystroke aimed at the
 * thumbnail strip still counts; it falls back to the overlay for trees that
 * skip `Root`. A target outside it belongs to the host page, which must keep
 * its own undo. An unfocused page (`body`) has no other claimant, so the deck
 * takes the keystroke: this is what makes the shortcut work before anything in
 * the presentation has been clicked.
 */
function getIsPresentationTarget(
  target: EventTarget | null,
  boundary: HTMLElement | null,
): boolean {
  if (!(target instanceof Node)) return true;
  if (target === document || target === document.body || target === document.documentElement) {
    return true;
  }
  // An editor toolbar is part of the deck's UI: undo after pressing Bold
  // undoes the bold.
  if (getIsKeepSelectionTarget(target)) return true;
  return boundary?.contains(target) ?? false;
}

function getIsUndoEvent(event: KeyboardEvent): boolean {
  if (!(event.ctrlKey || event.metaKey)) return false;
  return !event.shiftKey && event.key.toLowerCase() === "z";
}

function getIsRedoEvent(event: KeyboardEvent): boolean {
  if (!(event.ctrlKey || event.metaKey)) return false;
  const key = event.key.toLowerCase();
  if (key === "y") return !event.shiftKey;
  return key === "z" && event.shiftKey;
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
 * Walks the edited contentEditable container and produces a `setTextBody`
 * paragraphs payload. Uses the paragraph/run data attributes (see
 * `PPTX_ATTRS`) to map back to source indices for style inheritance.
 */
export function readBackTextBody(container: HTMLElement): SetTextBodyParagraph[] {
  const paragraphs: SetTextBodyParagraph[] = [];
  let lastSourceParagraph = 0;

  for (const paragraphDiv of getParagraphElements(container)) {
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
      /** Smart guides for what the drag snapped to. */
      guides?: SnapGuideData[];
    }
  | {
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
      nodeIds: string[];
      /** The node whose handle is being dragged; it sets the scale factors. */
      gripNodeId: string;
      grip: GripDirection;
      startX: number;
      startY: number;
      dx: number;
      dy: number;
      lockAspect: boolean;
      /** Smart guides for what the resized edges snapped to. */
      guides?: SnapGuideData[];
    }
  | {
      mode: "text";
      nodeId: string;
      editingElement: HTMLElement;
    };

/** Text range to put back after a formatting edit re-renders the text being edited. */
interface TextRestore {
  nodeId: string;
  range: TextRange;
  /**
   * Whether to move focus back into the text. Off while the toolbar is typing
   * into a field of its own (a font size box), which must keep focus.
   */
  focus: boolean;
}

/** The selection as reported to `onSelectionChange`. */
export interface SelectionChangeEvent {
  /** Ids of every selected node, in selection order. Empty when nothing is selected. */
  nodeIds: string[];

  /** The selected nodes themselves, matching `nodeIds`. */
  nodes: SlideNode[];
}

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

export interface SelectionProps extends PrimitiveProps<"div", SelectionState> {
  /**
   * When `true`, binds undo and redo to the deck: `Ctrl/Cmd+Z` and
   * `Ctrl/Cmd+Shift+Z` (or `Ctrl/Cmd+Y`).
   *
   * Shortcuts only apply within `Presentation.Root` and are ignored in text
   * fields. Off by default to avoid conflicts with the surrounding app.
   *
   * @default false
   *
   * ```tsx
   * <Presentation.Selection undoRedoShortcuts />
   * ```
   */
  undoRedoShortcuts?: boolean;

  /** Event handler called after an undo triggers. */
  onUndo?: (status: "success" | "empty", error?: unknown) => void;

  /** Event handler called after a redo triggers. */
  onRedo?: (status: "success" | "empty", error?: unknown) => void;

  /** Event handler called after a node delete is attempted (keyboard or pointer). */
  onNodeDelete?: (nodeId: string, error?: unknown) => void;

  /** Event handler called after a node move or resize is attempted. */
  onNodeTransform?: (nodeId: string, error?: unknown) => void;

  /** Event handler called after inline text editing is committed (or errors). */
  onTextChange?: (nodeId: string, error?: unknown) => void;

  /**
   * Event handler called whenever the set of selected nodes changes, so UI
   * outside the overlay (a formatting toolbar, a properties panel) can
   * follow along.
   *
   * Also called with an empty selection when the active slide changes, since
   * navigating away drops the selection. Treat the handler as idempotent.
   *
   * ```tsx
   * onSelectionChange={({ nodes }) => setInspectedShape(nodes[0] ?? null)}
   * ```
   */
  onSelectionChange?: (event: SelectionChangeEvent) => void;

  /**
   * Whether moving and resizing snap to the slide's edges and center and to
   * other shapes' edges and centers, with guides showing what lined up. Hold
   * Alt (Option) while dragging to place freely.
   *
   * @default true
   */
  snapping?: boolean;

  /**
   * How close, in screen pixels, an edge or center has to come to snap.
   *
   * @default 6
   */
  snapThreshold?: number;

  /**
   * Event handler called when the interaction mode changes, e.g. entering
   * text editing or starting a drag. Useful for disabling toolbars
   * mid-gesture.
   */
  onModeChange?: (mode: SelectionState["mode"], previousMode: SelectionState["mode"]) => void;
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
 * - Shift/Ctrl/Cmd+click toggles shapes in and out of the selection.
 * - Shift+drag a shape → move it along one axis only.
 * - Ctrl/Cmd+drag a shape duplicates it in PowerPoint; unsupported here, so the
 *   press only toggles the selection.
 * - Drag on empty canvas → marquee-select fully enclosed shapes.
 * - Shift/Ctrl/Cmd+drag on empty canvas → add the enclosed shapes to the selection.
 * - Every selected shape keeps its own handles, and dragging any one of them
 *   scales the whole selection, each shape in place.
 * - Ctrl/Cmd+A selects every shape on the slide.
 * - Arrow keys nudge, Delete removes: applied to the whole selection.
 */
const SelectionImpl = React.forwardRef<HTMLDivElement, SelectionProps>(function SelectionImpl(
  {
    render,
    onNodeDelete,
    onNodeTransform,
    onTextChange,
    onSelectionChange,
    onModeChange,
    snapping = true,
    snapThreshold = SNAP_THRESHOLD,
    children,
    ...selectionImplProps
  },
  forwardedRef,
) {
  const store = useStoreContext(SELECTION_NAME);
  const { presentation } = usePresentation();
  const { slide, slideId } = useSlide();
  const { zoom } = useZoom();

  const rootRef = React.useRef<HTMLDivElement>(null);
  const frameRef = React.useRef<HTMLDivElement>(null);
  const [state, setState] = React.useState<InternalState>({ mode: "idle" });

  const stateRef = useLatestRef(state);

  // Shallow clone of a styled run span, captured per shape on entering text
  // mode. When the browser destroys all run spans (select-all + delete),
  // typed text is re-wrapped with this template so it keeps the run's
  // styling.
  const runTemplateRef = React.useRef<{ nodeId: string; span: HTMLElement } | null>(null);

  // The last text range selected in text mode. Pressing a toolbar moves focus
  // (and often the document selection) out of the text, so formatting reads
  // the range from here once the live selection has left.
  const savedRangeRef = React.useRef<Range | null>(null);

  // Set by a formatting edit in text mode, consumed when the re-rendered text
  // is re-entered.
  const pendingRestoreRef = React.useRef<TextRestore | null>(null);

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
  const selectedNode = selectedNodes[0] ?? null;
  const isSoloSelection = selectedNodes.length === 1;

  const isTextMode = state.mode === "text";
  const publicState: SelectionState = { mode: state.mode, selectedNode, selectedNodes };

  // Selection transitions happen across ~25 pointer/keyboard paths, several of
  // them triggered from document-level listeners, so notifying from each handler
  // would mean threading the callback through all of them. Diffing a single
  // derived key keeps one notification site and guarantees it can't drift out
  // of sync with the rendered selection.
  const onSelectionChangeRef = useLatestRef(onSelectionChange);
  const onModeChangeRef = useLatestRef(onModeChange);
  const selectedIdsRef = useLatestRef(selectedIds);
  const selectedNodesRef = useLatestRef(selectedNodes);
  const selectionKey = selectedIds.join("\u0000");
  const mode = state.mode;
  const previousModeRef = React.useRef(mode);

  React.useEffect(() => {
    onSelectionChangeRef.current?.({
      nodeIds: selectedIdsRef.current,
      nodes: selectedNodesRef.current,
    });
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- the joined key is the change signal; payload is read from refs
  }, [selectionKey, onSelectionChangeRef, selectedIdsRef, selectedNodesRef]);

  React.useEffect(() => {
    const previousMode = previousModeRef.current;
    if (previousMode === mode) return;
    previousModeRef.current = mode;
    onModeChangeRef.current?.(mode, previousMode);
  }, [mode, onModeChangeRef]);

  const isBanding =
    state.mode === "marquee" && isBandDrag(state.startX, state.startY, state.curX, state.curY);

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
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- nodes mutate in place; revision is the change signal
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
   * Drops any document text selection left behind by earlier gestures.
   * A stale selection under the pointer turns the next press-and-drag into a
   * native drag-and-drop, cancelling our pointer stream mid-resize/move.
   */
  function clearDocumentSelection(): void {
    const selection = document.getSelection();
    if (selection && !selection.isCollapsed) selection.removeAllRanges();
  }

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
    restore?: TextRestore | null,
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

    savedRangeRef.current = null;
    if (restore?.focus !== false) textElement.focus({ preventScroll: true });

    if (restore) {
      restoreTextRange(textElement, restore);
    } else if (clientX !== undefined && clientY !== undefined) {
      placeCaretAtPoint(clientX, clientY);
    } else {
      placeCaretAtEnd(textElement);
    }

    // PowerPoint: typing while a shape is selected starts editing with that
    // keystroke as the first character.
    if (insertText && !restore && !insertTextNatively(textElement, insertText)) {
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
   * Places the caret inside the last styled run span of `scope`. Typed text
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

  function runTemplateFor(nodeId: string): HTMLElement | null {
    const entry = runTemplateRef.current;
    return entry && entry.nodeId === nodeId ? entry.span : null;
  }

  /**
   * Inserts the keystroke that started text editing the way typing does, so it
   * lands on the browser's undo stack: Cmd/Ctrl+Z in text mode is native undo,
   * and a character inserted by mutating the DOM would be invisible to it.
   *
   * Only when the caret is inside a run's text, where native insertion keeps
   * the run's styling. Returns false otherwise (or without `execCommand`), and
   * the caller falls back to `insertTextAtCaret`.
   */
  function insertTextNatively(editingElement: HTMLElement, data: string): boolean {
    const anchor = window.getSelection()?.anchorNode;
    if (
      !anchor ||
      anchor.nodeType !== Node.TEXT_NODE ||
      !anchor.textContent ||
      !editingElement.contains(anchor) ||
      !anchor.parentElement?.closest(`[${PPTX_ATTRS.run}]`) ||
      typeof document.execCommand !== "function"
    ) {
      return false;
    }
    return document.execCommand("insertText", false, data);
  }

  /**
   * Inserts `data` at the current caret, keeping it inside a styled run span
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
        const textNode = anchor as Text;
        const offset = Math.min(selection.anchorOffset, textNode.length);
        textNode.insertData(offset, data);
        setCaret(textNode, offset + data.length);
      } else {
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

  /**
   * Intercept text insertion when the browser would drop it outside a styled
   * run span. Chrome cannot keep the caret inside an empty text node (the
   * keystroke escapes into the parent div), and destructive edits like
   * select-all + delete remove the run spans entirely; in both cases typed
   * text would render with unstyled defaults (e.g. near-white → invisible).
   * Instead of the default insertion we place the text into the run span at
   * the caret, or into a clone of the span captured on entering text mode.
   */
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

    if (nextNodeId === null) {
      setState({ mode: "idle" });
    } else {
      setState({ mode: "selected", nodeIds: [nextNodeId ?? current.nodeId] });
    }
    rootRef.current?.focus({ preventScroll: true });
  }

  /** Puts a model text range back into freshly rendered text. */
  function restoreTextRange(textElement: HTMLElement, restore: TextRestore): void {
    const start = getDomPoint(textElement, restore.range.start);
    const end = getDomPoint(textElement, restore.range.end);
    if (!start || !end) return;
    const range = document.createRange();
    range.setStart(...start);
    range.setEnd(...end);
    savedRangeRef.current = range;
    if (!restore.focus) return;
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range.cloneRange());
  }

  /** The text container being edited, re-resolved if a re-render replaced it. */
  function getEditingContainer(
    current: Extract<InternalState, { mode: "text" }>,
  ): HTMLElement | null {
    if (current.editingElement.isConnected) return current.editingElement;
    const shapeElement = getShapeElement(current.nodeId);
    return shapeElement ? textContainerOf(shapeElement) : null;
  }

  /** The text range to format: the live selection in the text, else the saved one. */
  function getEditingDomRange(container: HTMLElement): Range | null {
    const selection = window.getSelection();
    const live = selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
    for (const range of [live, savedRangeRef.current]) {
      if (
        range &&
        container.contains(range.startContainer) &&
        container.contains(range.endContainer)
      ) {
        return range;
      }
    }
    return null;
  }

  function readTextStyle(): ActiveTextStyle | null {
    const current = stateRef.current;
    if (current.mode === "text") {
      const container = getEditingContainer(current);
      if (!container) return null;
      const range = getEditingDomRange(container);
      if (range) {
        const { runs, paragraphs } = getRangeRuns(container, range);
        return getActiveTextStyle(runs, paragraphs);
      }
      return getActiveTextStyle(
        Array.from(container.querySelectorAll<HTMLElement>(`[${PPTX_ATTRS.run}]`)),
        getParagraphElements(container),
      );
    }

    const runs: HTMLElement[] = [];
    const paragraphs: HTMLElement[] = [];
    for (const nodeId of selectedIdsRef.current) {
      const shapeElement = getShapeElement(nodeId);
      const container = shapeElement ? textContainerOf(shapeElement) : null;
      if (!container) continue;
      const containerParagraphs = getParagraphElements(container);
      const containerRuns = container.querySelectorAll<HTMLElement>(`[${PPTX_ATTRS.run}]`);
      paragraphs.push(...containerParagraphs);
      runs.push(...(containerRuns.length > 0 ? containerRuns : containerParagraphs));
    }
    return getActiveTextStyle(runs, paragraphs);
  }

  const selectionChannel = getSelectionChannel(store);

  function publishSelection(): void {
    selectionChannel.publish({
      slideId,
      mode: stateRef.current.mode,
      nodeIds: selectedIdsRef.current,
      textStyle: readTextStyle(),
    });
  }

  /** `useSelection().select()`: selects nodes on this slide, ending text editing. */
  function selectNodes(nodeIds: string[]): void {
    const ids = nodeIds.filter((id) => slide?.nodes.some((node) => node.id === id));
    const current = stateRef.current;
    if (current.mode === "text") commitTextEdits(current);
    setState(ids.length > 0 ? { mode: "selected", nodeIds: ids } : { mode: "idle" });
  }

  /**
   * `useSelection().setTextStyle()` and `setParagraphStyle()`.
   *
   * In text mode, uncommitted typing is committed in the same batch, so the
   * formatting lands on the text as shown and the whole thing is one undo
   * step. The slide re-renders afterwards, and the repair effect puts the
   * caret or selection back.
   */
  async function applyFormat(
    kind: "text" | "paragraph",
    style: TextStyle | ParagraphStyle,
  ): Promise<void> {
    if (!slide || !slideId) return;
    const toOperation = (nodeId: string, range?: TextRange): EditOperation =>
      kind === "text"
        ? { type: "setTextStyle", slideId, nodeId, style: style as TextStyle, range }
        : { type: "setParagraphStyle", slideId, nodeId, style: style as ParagraphStyle, range };

    const current = stateRef.current;
    if (current.mode !== "text") {
      const operations = selectedNodesRef.current
        .filter((node) => nodeHasText(node.id))
        .map((node) => toOperation(node.id));
      if (operations.length === 0) return;
      await store.edit(operations.length === 1 ? operations[0] : { type: "batch", operations });
      return;
    }

    const { nodeId } = current;
    const container = getEditingContainer(current);
    const node = slide.nodes.find((n) => n.id === nodeId);
    if (!container || !node) return;

    const readBack = readBackTextBody(container);
    const changed = textBodyChanged(node, readBack);
    // Text read back from editing drops line breaks; untouched text keeps them.
    const countBreaks = !changed;
    const domRange = getEditingDomRange(container);
    const caretRange = domRange ? getTextRange(container, domRange, countBreaks) : null;

    let range = caretRange ?? undefined;
    if (
      range &&
      kind === "text" &&
      range.start.paragraphIndex === range.end.paragraphIndex &&
      range.start.offset === range.end.offset
    ) {
      // A caret formats the word it touches, like PowerPoint. In an empty
      // paragraph it formats the paragraph mark, so typed text picks it up.
      const { paragraphIndex, offset } = range.start;
      const paragraph = getParagraphElements(container)[paragraphIndex];
      const text = paragraph ? getParagraphText(paragraph, countBreaks) : "";
      if (text.length > 0) {
        const word = expandToWord(text, offset);
        if (!word) return;
        range = {
          start: { paragraphIndex, offset: word[0] },
          end: { paragraphIndex, offset: word[1] },
        };
      }
    }

    const operations: EditOperation[] = [];
    if (changed) operations.push({ type: "setTextBody", slideId, nodeId, paragraphs: readBack });
    operations.push(toOperation(nodeId, range));

    pendingRestoreRef.current = caretRange
      ? {
          nodeId,
          range: caretRange,
          focus:
            !getIsNativeUndoTarget(document.activeElement) ||
            container.contains(document.activeElement),
        }
      : null;
    try {
      await store.edit(operations.length === 1 ? operations[0] : { type: "batch", operations });
    } catch (error) {
      pendingRestoreRef.current = null;
      throw error;
    }
  }

  const enterTextModeRef = useLatestRef(enterTextMode);
  const commitTextEditsRef = useLatestRef(commitTextEdits);
  const doExitTextModeRef = useLatestRef(doExitTextMode);
  const publishSelectionRef = useLatestRef(publishSelection);
  const selectNodesRef = useLatestRef(selectNodes);
  const applyFormatRef = useLatestRef(applyFormat);

  React.useEffect(
    () =>
      selectionChannel.register({
        select: (nodeIds) => selectNodesRef.current(nodeIds),
        setTextStyle: (style) => applyFormatRef.current("text", style),
        setParagraphStyle: (style) => applyFormatRef.current("paragraph", style),
      }),
    [selectionChannel, selectNodesRef, applyFormatRef],
  );

  // Publish to `useSelection` when the selection, the mode, or the slide's
  // content changes. An edit replaces the slide DOM in SlideImpl's effect,
  // which runs after this one (it is a parent), so the text style is read
  // again a frame later from the fresh DOM.
  React.useEffect(() => {
    publishSelectionRef.current();
    const frameId = requestAnimationFrame(() => publishSelectionRef.current());
    return () => cancelAnimationFrame(frameId);
  }, [selectionKey, mode, slideRevision, publishSelectionRef]);

  // In text mode, follow the caret: remember the range for toolbar presses,
  // and republish so the toolbar's toggles track the text under the caret.
  React.useEffect(() => {
    if (!isTextMode) return;

    function onSelectionChange(): void {
      const currentState = stateRef.current;
      if (currentState.mode !== "text") return;
      const selection = window.getSelection();
      if (selection && selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        const container = currentState.editingElement;
        if (container.contains(range.startContainer) && container.contains(range.endContainer)) {
          savedRangeRef.current = range.cloneRange();
        }
      }
      publishSelectionRef.current();
    }

    document.addEventListener("selectionchange", onSelectionChange);
    return () => document.removeEventListener("selectionchange", onSelectionChange);
  }, [isTextMode, stateRef, publishSelectionRef]);

  // When in text mode the overlay has pointerEvents: none, so we must listen
  // on the document for clicks-outside and Escape.
  React.useEffect(() => {
    if (!isTextMode) return;

    function onDocPointerDown(event: PointerEvent): void {
      const currentState = stateRef.current;
      if (currentState.mode !== "text") return;

      if (currentState.editingElement.contains(event.target as Node)) return;

      // Clicks on the overlay's own children (border move strips) are
      // handled by their own handlers.
      if (rootRef.current?.contains(event.target as Node)) return;

      // A toolbar formats the text being edited, so pressing it must not end
      // editing. The text range is remembered (see `savedRangeRef`).
      if (getIsKeepSelectionTarget(event.target)) return;

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
        commitTextEditsRef.current(currentState);
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
        doExitTextModeRef.current(currentState, null);
      }
      // Else clicked inside the same shape but outside the text container
      // (e.g. shape padding area); stay in text mode.
    }

    function onDocKeyDown(event: KeyboardEvent): void {
      const currentState = stateRef.current;
      if (currentState.mode !== "text") return;

      // Escape inside a toolbar closes its popover, not the text editing.
      if (event.key === "Escape" && !getIsKeepSelectionTarget(event.target)) {
        event.preventDefault();
        doExitTextModeRef.current(currentState);
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
  }, [isTextMode, stateRef, commitTextEditsRef, doExitTextModeRef]);

  /**
   * A press outside the overlay deselects, the way clicking off the canvas
   * does in PowerPoint. The overlay covers only the slide and its pasteboard,
   * so without this the handles survive a press on the viewport margin, and
   * survive it out of reach: focus leaves the overlay and `onKeyDown` is a
   * React handler on it, so Escape stops arriving.
   *
   * Presses inside UI marked `data-pptx-keep-selection` (editor toolbars and
   * their popovers) don't count as outside.
   *
   * Deselecting on release rather than on press is what keeps touch usable.
   * A scroll begins as a pointerdown on whatever is under the finger, so
   * acting there would drop the selection every time the deck was scrolled.
   * The browser triggers pointercancel when it takes the gesture over for
   * scrolling, which distinguishes a scroll from a tap without guessing at a
   * movement threshold.
   *
   * Text mode has its own listener above, which commits edits first.
   */
  React.useEffect(() => {
    if (isTextMode) return;

    let pending = false;

    function onDocPointerDown(event: PointerEvent): void {
      pending = false;
      // Primary button only, as in `onPointerDown`; touch and pen report 0.
      if (event.button !== 0) return;
      if (stateRef.current.mode === "idle") return;
      // Drags capture the pointer onto the overlay, so this skips those too.
      if (rootRef.current?.contains(event.target as Node)) return;
      // Toolbars act on the selection, so pressing them must keep it.
      if (getIsKeepSelectionTarget(event.target)) return;
      pending = true;
    }

    function onDocPointerUp(): void {
      if (!pending) return;
      pending = false;
      setState({ mode: "idle" });
    }

    function onDocPointerCancel(): void {
      pending = false;
    }

    document.addEventListener("pointerdown", onDocPointerDown, true);
    document.addEventListener("pointerup", onDocPointerUp, true);
    document.addEventListener("pointercancel", onDocPointerCancel, true);
    return () => {
      document.removeEventListener("pointerdown", onDocPointerDown, true);
      document.removeEventListener("pointerup", onDocPointerUp, true);
      document.removeEventListener("pointercancel", onDocPointerCancel, true);
    };
  }, [isTextMode, stateRef]);

  // A revision bump makes SlideImpl replace the slide DOM in its effect,
  // detaching our contentEditable element: typed text would go into the
  // detached tree and never appear on screen. SlideImpl is a parent, so its
  // effect runs after this one; the rAF triggers after the whole effects flush,
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
      const restore =
        pendingRestoreRef.current?.nodeId === currentState.nodeId
          ? pendingRestoreRef.current
          : null;
      pendingRestoreRef.current = null;
      enterTextModeRef.current(currentState.nodeId, undefined, undefined, undefined, restore);
      // Re-read the toolbar state from the fresh text, which moves no caret
      // (and fires no selectionchange) when focus stayed in the toolbar.
      publishSelectionRef.current();
    });
    return () => cancelAnimationFrame(frameId);
  }, [slideRevision, stateRef, enterTextModeRef, publishSelectionRef]);

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

  /** What a drag can snap to: the slide, and every shape not being dragged. */
  function getSnapTargets(draggedIds: string[]): Rect[] {
    const targets: Rect[] = [{ x: 0, y: 0, w: presentation!.width, h: presentation!.height }];
    for (const node of slide!.nodes) {
      if (!draggedIds.includes(node.id)) targets.push(getNodeRect(node));
    }
    return targets;
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
    // Alt places freely, as in PowerPoint. The threshold is in screen pixels.
    const isSnapping = snapping && !event.altKey;
    const threshold = snapThreshold / zoom;

    if (state.mode === "move") {
      const moved =
        state.moved ||
        Math.hypot(event.clientX - state.startX, event.clientY - state.startY) > DRAG_THRESHOLD;
      // Track Shift live (as with resize) so holding or releasing it mid-drag
      // locks and unlocks the axis. The constrained (and snapped) deltas are
      // what lands in state, so pointer-up commits them without re-deriving.
      const delta = constrainMove(dx, dy, event.shiftKey);
      const bounds = getUnionRect(selectedNodes.map(getNodeRect));
      const next =
        moved && isSnapping && bounds
          ? snapMove(bounds, delta.x, delta.y, getSnapTargets(state.nodeIds), threshold, {
              // An axis Shift locked stays put rather than snapping off it.
              x: !(event.shiftKey && delta.x === 0),
              y: !(event.shiftKey && delta.y === 0),
            })
          : { dx: delta.x, dy: delta.y, guides: [] };
      if (moved) {
        for (const node of selectedNodes) {
          const shapeElement = getShapeElement(node.id);
          if (shapeElement) {
            shapeElement.style.left = `${node.position.x + next.dx}px`;
            shapeElement.style.top = `${node.position.y + next.dy}px`;
          }
        }
      }
      setState({ ...state, dx: next.dx, dy: next.dy, moved, guides: next.guides });
    } else {
      // Track Shift live so holding/releasing it mid-drag toggles the lock.
      const lockAspect = event.shiftKey;
      const grippedNode = selectedNodes.find((node) => node.id === state.gripNodeId);
      const next =
        isSnapping && grippedNode
          ? snapResize(
              getNodeRect(grippedNode),
              state.grip,
              dx,
              dy,
              getSnapTargets(state.nodeIds),
              threshold,
              lockAspect,
            )
          : { dx, dy, guides: [] };
      setState({ ...state, dx: next.dx, dy: next.dy, lockAspect, guides: next.guides });
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
    // only triggers when the overlay div has focus (selected/idle modes).
    // Undo/redo shortcuts live in the outer Selection wrapper (document
    // level) so they survive the slide-change remount of this component.
    if (isTextMode) return;

    const mod = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();

    if (mod && key === "a") {
      event.preventDefault();
      const allIds = slide!.nodes.map((n) => n.id);
      if (allIds.length > 0) setState({ mode: "selected", nodeIds: allIds });
      return;
    }

    if (selectedNodes.length === 0) return;

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
                ref={frameRef}
                style={{
                  position: "absolute",
                  left: pasteboard.left * zoom,
                  top: pasteboard.top * zoom,
                  width: presentation.width * zoom,
                  height: presentation.height * zoom,
                  pointerEvents: "none",
                }}
              >
                <SelectionPartsContext.Provider
                  value={{
                    state,
                    zoom,
                    selectedNodes,
                    resizeScale,
                    showResizeHandles: !isBanding,
                    frameRef,
                    onGripPointerDown,
                    onBorderPointerDown,
                  }}
                >
                  {children ?? DEFAULT_PARTS}
                </SelectionPartsContext.Provider>
              </div>
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
        selectionImplProps,
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
 * Being on the document does not make them the page's: they only act on
 * keystrokes aimed at the presentation (see `getIsPresentationTarget`), so an
 * app that owns Ctrl/Cmd+Z elsewhere keeps it.
 */
export const Selection = React.forwardRef<HTMLDivElement, SelectionProps>(function Selection(
  { undoRedoShortcuts = false, onUndo, onRedo, ...selectionProps },
  forwardedRef,
) {
  const store = useStoreContext(SELECTION_NAME);
  const { presentation } = usePresentation();
  const { slideId } = useSlide();
  const rootRef = React.useContext(RootContext);

  // Focus boundary for trees without a `Root`: the overlay is all there is.
  const overlayRef = React.useRef<HTMLDivElement | null>(null);

  const onUndoRef = useLatestRef(onUndo);
  const onRedoRef = useLatestRef(onRedo);

  const editable = presentation?.sourcePackage != null;

  React.useEffect(() => {
    if (!editable || !undoRedoShortcuts) return;

    function onDocKeyDown(event: KeyboardEvent): void {
      if (getIsNativeUndoTarget(event.target)) return;
      const boundary = rootRef?.current ?? overlayRef.current;
      if (!getIsPresentationTarget(event.target, boundary)) return;

      if (getIsUndoEvent(event)) {
        event.preventDefault();
        const success = store.undo();
        onUndoRef.current?.(success ? "success" : "empty");
      } else if (getIsRedoEvent(event)) {
        event.preventDefault();
        store.redo().then(
          (success) => onRedoRef.current?.(success ? "success" : "empty"),
          (error) => onRedoRef.current?.("empty", error),
        );
      }
    }

    document.addEventListener("keydown", onDocKeyDown);
    return () => document.removeEventListener("keydown", onDocKeyDown);
  }, [editable, undoRedoShortcuts, store, rootRef, onUndoRef, onRedoRef]);

  if (!slideId) return null;

  return (
    <SelectionImpl key={slideId} ref={mergeRefs(overlayRef, forwardedRef)} {...selectionProps} />
  );
});

// ---------------------------------------------------------------------------
// Overlay parts
// ---------------------------------------------------------------------------

/**
 * What the overlay hands its parts. Internal: each part exposes only its own
 * state, so the overlay's interaction state stays free to change.
 */
interface SelectionPartsContextValue {
  state: InternalState;
  zoom: number;
  selectedNodes: SlideNode[];
  resizeScale: ResizeScale | null;
  /**
   * False mid-band: the handles are unreachable then (the pointer is captured)
   * and would sit under the rubber band. Until the band starts, the press is
   * still a click and must not blink the handles of what is already selected.
   */
  showResizeHandles: boolean;
  /** The slide-origin frame the parts are positioned in. */
  frameRef: React.RefObject<HTMLDivElement | null>;
  onGripPointerDown: (
    event: React.PointerEvent<HTMLDivElement>,
    grip: GripDirection,
    nodeId: string,
  ) => void;
  onBorderPointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
}

const SelectionPartsContext = React.createContext<SelectionPartsContextValue | null>(null);

function useSelectionParts(partName: string): SelectionPartsContextValue {
  const context = React.useContext(SelectionPartsContext);
  if (!context) {
    throw new Error(`\`${partName}\` must be used within \`Presentation.Selection\`.`);
  }
  return context;
}

interface SelectionBoxContextValue {
  node: SlideNode;
  showHandles: boolean;
}

const SelectionBoxContext = React.createContext<SelectionBoxContextValue | null>(null);

const SELECTION_BOX_NAME = "Presentation.SelectionBox";
const RESIZE_HANDLE_NAME = "Presentation.ResizeHandle";
const SNAP_GUIDE_NAME = "Presentation.SnapGuide";
const MARQUEE_NAME = "Presentation.Marquee";

/** The current state of one selection outline. */
export interface SelectionBoxState {
  /** The selected shape the outline belongs to. */
  node: SlideNode;

  /** Interaction mode of the overlay. `"text"` while the shape's text is edited. */
  mode: SelectionState["mode"];

  /** Whether the selection is being dragged. */
  dragging: boolean;
}

export interface SelectionBoxProps extends StatefulPrimitiveProps<"div", SelectionBoxState> {}

/**
 * The outline around each selected shape. Rendered once per selected shape,
 * following it through moves, resizes, and rotation.
 *
 * Place inside `Presentation.Selection`, with `Presentation.ResizeHandle`
 * inside it for handles.
 *
 * Carries `data-selection-box`, `data-mode`, and `data-dragging`.
 */
export const SelectionBox = React.forwardRef<HTMLDivElement, SelectionBoxProps>(
  function SelectionBox(
    { render, className, style, children, ...selectionBoxProps },
    forwardedRef,
  ) {
    const { state, zoom, selectedNodes, resizeScale, showResizeHandles, onBorderPointerDown } =
      useSelectionParts(SELECTION_BOX_NAME);

    const isTextMode = state.mode === "text";
    // Handles hide once a drag is actually under way, not merely on pointer-down:
    // a press that only toggles the selection (Shift+click) would otherwise blink
    // them off and back on.
    const isDragging = state.mode === "move" && state.moved;

    return selectedNodes.map((node) => {
      let rect = getNodeRect(node);
      if (state.mode === "move" && state.moved) {
        rect = { ...rect, x: rect.x + state.dx, y: rect.y + state.dy };
      } else if (state.mode === "resize" && resizeScale) {
        rect = scaleRectFromGrip(rect, state.grip, resizeScale);
      }
      const showHandles = showResizeHandles && node.rotation === 0 && !isDragging && !isTextMode;
      const boxState: SelectionBoxState = { node, mode: state.mode, dragging: isDragging };

      return (
        <SelectionBoxContext.Provider key={node.id} value={{ node, showHandles }}>
          {renderElement(
            "div",
            { render, className, style },
            {
              state: boxState,
              ref: forwardedRef,
              props: [
                {
                  "data-selection-box": "",
                  "data-mode": state.mode,
                  "data-dragging": isDragging ? "" : undefined,
                  children: (
                    <>
                      {children}
                      {isTextMode && <FrameGrabHitboxes onPointerDown={onBorderPointerDown} />}
                    </>
                  ),
                  style: {
                    position: "absolute",
                    left: rect.x * zoom,
                    top: rect.y * zoom,
                    width: rect.w * zoom,
                    height: rect.h * zoom,
                    transform: node.rotation !== 0 ? `rotate(${node.rotation}deg)` : undefined,
                    pointerEvents: "none",
                  },
                },
                selectionBoxProps,
              ],
            },
          )}
        </SelectionBoxContext.Provider>
      );
    });
  },
);

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

/** The current state of one resize handle. */
export interface ResizeHandleState {
  /** Which edge or corner the handle drags. */
  direction: GripDirection;
}

export interface ResizeHandleProps extends StatefulPrimitiveProps<"div", ResizeHandleState> {}

/**
 * The resize handles of a selected shape: one per edge and corner, centered
 * on the outline. Hidden while dragging, editing text, and on rotated shapes.
 *
 * Unstyled when composed: give it a size, or it is a zero-size target.
 *
 * Place inside `Presentation.SelectionBox`. Carries `data-resize-grip`, set
 * to its direction.
 */
export const ResizeHandle = React.forwardRef<HTMLDivElement, ResizeHandleProps>(
  function ResizeHandle({ render, className, style, ...resizeHandleProps }, forwardedRef) {
    const { onGripPointerDown } = useSelectionParts(RESIZE_HANDLE_NAME);
    const box = React.useContext(SelectionBoxContext);
    if (!box) {
      throw new Error(`\`${RESIZE_HANDLE_NAME}\` must be used within \`${SELECTION_BOX_NAME}\`.`);
    }
    if (!box.showHandles) return null;

    return GRIP_DIRECTIONS.map((direction) => (
      <React.Fragment key={direction}>
        {renderElement(
          "div",
          { render, className, style },
          {
            state: { direction },
            ref: forwardedRef,
            props: [
              {
                "aria-hidden": true,
                "data-resize-grip": direction,
                onPointerDown: (event: React.PointerEvent<HTMLDivElement>) =>
                  onGripPointerDown(event, direction, box.node.id),
                style: {
                  position: "absolute",
                  ...GRIP_POSITIONS[direction],
                  transform: "translate(-50%, -50%)",
                  cursor: GRIP_CURSORS[direction],
                  pointerEvents: "auto",
                  touchAction: "none",
                },
              },
              resizeHandleProps,
            ],
          },
        )}
      </React.Fragment>
    ));
  },
);

/** The current state of one snap guide. */
export interface SnapGuideState {
  /** `"x"` for a vertical line, `"y"` for a horizontal one. */
  axis: SnapGuideData["axis"];

  /** Where the line sits on its axis, in slide pixels. */
  position: number;
}

export interface SnapGuideProps extends StatefulPrimitiveProps<"div", SnapGuideState> {}

/**
 * The guides shown while a move or resize snaps: one line per edge or center
 * that lined up, spanning everything aligned on it.
 *
 * The element is the line itself, `--presentation-snap-guide-width` thick
 * (1px by default) and centered on the snapped position, so a background
 * color draws it.
 *
 * Place inside `Presentation.Selection`. Carries `data-snap-guide`, set to
 * its axis.
 */
export const SnapGuide = React.forwardRef<HTMLDivElement, SnapGuideProps>(function SnapGuide(
  { render, className, style, ...snapGuideProps },
  forwardedRef,
) {
  const { state, zoom } = useSelectionParts(SNAP_GUIDE_NAME);
  if (state.mode !== "move" && state.mode !== "resize") return null;

  return (state.guides ?? []).map((guide) => {
    const isVertical = guide.axis === "x";
    const length = (guide.end - guide.start) * zoom;
    const thickness = "var(--presentation-snap-guide-width, 1px)";
    return (
      <React.Fragment key={`${guide.axis}:${guide.position}`}>
        {renderElement(
          "div",
          { render, className, style },
          {
            state: { axis: guide.axis, position: guide.position },
            ref: forwardedRef,
            props: [
              {
                "aria-hidden": true,
                "data-snap-guide": guide.axis,
                style: {
                  position: "absolute",
                  left: isVertical ? guide.position * zoom : guide.start * zoom,
                  top: isVertical ? guide.start * zoom : guide.position * zoom,
                  width: isVertical ? thickness : length,
                  height: isVertical ? length : thickness,
                  transform: isVertical ? "translateX(-50%)" : "translateY(-50%)",
                  pointerEvents: "none",
                },
              },
              snapGuideProps,
            ],
          },
        )}
      </React.Fragment>
    );
  });
});

export interface MarqueeProps extends StatefulPrimitiveProps<"div", Record<string, never>> {}

/**
 * The rubber band drawn while dragging across empty canvas to select.
 *
 * Place inside `Presentation.Selection`. Carries `data-marquee`.
 */
export const Marquee = React.forwardRef<HTMLDivElement, MarqueeProps>(function Marquee(
  { render, className, style, ...marqueeProps },
  forwardedRef,
) {
  const { state, frameRef } = useSelectionParts(MARQUEE_NAME);
  if (state.mode !== "marquee") return null;
  if (!isBandDrag(state.startX, state.startY, state.curX, state.curY)) return null;
  // State coordinates are viewport-relative; the frame is the positioning context.
  const origin = frameRef.current?.getBoundingClientRect();
  if (!origin) return null;

  return renderElement(
    "div",
    { render, className, style },
    {
      state: {},
      ref: forwardedRef,
      props: [
        {
          "data-marquee": "",
          style: {
            position: "absolute",
            left: Math.min(state.startX, state.curX) - origin.left,
            top: Math.min(state.startY, state.curY) - origin.top,
            width: Math.abs(state.curX - state.startX),
            height: Math.abs(state.curY - state.startY),
            pointerEvents: "none",
          },
        },
        marqueeProps,
      ],
    },
  );
});

/**
 * The parts `Presentation.Selection` renders without children: the default
 * look, built from the same parts an app composes, themed through the
 * `--presentation-selection` and `--presentation-snap-guide` variables.
 */
const DEFAULT_PARTS = (
  <>
    <SelectionBox
      style={({ mode }) => ({
        boxShadow: `0 0 0 ${mode === "text" ? "2" : "1.5"}px var(--presentation-selection, #2563eb)`,
      })}
    >
      <ResizeHandle
        style={{
          width: 9,
          height: 9,
          background: "#fff",
          border: "1.5px solid var(--presentation-selection, #2563eb)",
          borderRadius: 2,
        }}
      />
    </SelectionBox>
    <SnapGuide style={{ background: "var(--presentation-snap-guide, #f43f5e)" }} />
    <Marquee
      style={{
        border: "1px solid var(--presentation-selection, #2563eb)",
        background: "color-mix(in srgb, var(--presentation-selection, #2563eb) 10%, transparent)",
      }}
    />
  </>
);

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

export namespace SelectionBox {
  export type State = SelectionBoxState;
  export type Props = SelectionBoxProps;
}

export namespace ResizeHandle {
  export type State = ResizeHandleState;
  export type Props = ResizeHandleProps;
}

export namespace SnapGuide {
  export type State = SnapGuideState;
  export type Props = SnapGuideProps;
}

export namespace Marquee {
  export type Props = MarqueeProps;
}
