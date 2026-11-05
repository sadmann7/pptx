import * as React from "react";

import type {
  EditOperation,
  EditResult,
  NodeOrder,
  NodePosition,
  ParagraphStyle,
  ShapeNodeData,
  SlideNode,
  TextRange,
  TextStyle,
} from "@diceui/pptx-core";
import { copyNodes, PPTX_ATTRS, PPTX_DATASET } from "@diceui/pptx-core";

import {
  RootContext,
  usePresentation,
  useSlide,
  useSlideRevision,
  useStoreContext,
  useZoom,
} from "../context";
import { useLatestRef } from "../hooks";
import type { PrimitiveProps } from "../render";
import { mergeRefs, renderElement } from "../render";
import type { DrawTool } from "./channel";
import { getIsKeepSelectionTarget, getSelectionChannel } from "./channel";
import { readClipboardToken, stashNodes, takeStash, writeClipboardToken } from "./clipboard";
import { SelectionPartsContext } from "./context";
import type { GripDirection, Rect } from "./geometry";
import {
  constrainMove,
  DRAG_THRESHOLD,
  getGripCursor,
  getNodeBounds,
  getResizedRects,
  getNodeRect,
  getPasteboardOverhang,
  getPointerAngle,
  isBandDrag,
  mergeSelection,
  normalizeAngle,
  ROTATION_STEP,
  toggleSelection,
} from "./geometry";
import { Marquee } from "./marquee";
import { ResizeHandle } from "./resize-handle";
import { ROTATE_CURSOR, RotateHandle } from "./rotate-handle";
import { SelectionBox } from "./selection-box";
import type { SnapGuide as SnapGuideData } from "./snap";
import { getUnionRect, SNAP_THRESHOLD, snapMove, snapResize } from "./snap";
import { SnapGuide } from "./snap-guide";
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
  placeholderPromptOf,
  readBackTextBody,
  textBodyChanged,
  textContainerOf,
} from "./text-range";

const SELECTION_NAME = "Presentation.Selection";

const ENABLE_DEBUG_LOG = false;

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

const VISUALLY_HIDDEN: React.CSSProperties = {
  position: "absolute",
  width: 1,
  height: 1,
  margin: -1,
  padding: 0,
  overflow: "hidden",
  clip: "rect(0 0 0 0)",
  whiteSpace: "nowrap",
  border: 0,
};

const PASTABLE_IMAGE_TYPES = new Set(["image/png", "image/jpeg", "image/gif"]);

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

export type InternalState =
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
      /** The pressed shape was already the only one selected: a click edits its text. */
      wasSelected?: boolean;
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
      /** Dragging out a new text box or shape with an armed draw tool. */
      mode: "draw";
      tool: DrawTool;
      startX: number;
      startY: number;
      curX: number;
      curY: number;
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
      mode: "rotate";
      nodeIds: string[];
      /** The node whose handle is being dragged; its angle is what Shift steps. */
      gripNodeId: string;
      /** Center of the gripped node, in client pixels: the pivot. */
      centerX: number;
      centerY: number;
      /** Pointer angle around the center at pointer-down, in degrees. */
      startAngle: number;
      /** Degrees added to every selected node's rotation. */
      delta: number;
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
  mode: "idle" | "selected" | "move" | "resize" | "rotate" | "text" | "marquee" | "draw";

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
 * Interaction model mirrors PowerPoint and Google Slides:
 * - Click a shape → select it. Click it again, double-click, or type → edit its text.
 * - Drag any shape → move it (a multi-selection moves as a unit).
 * - Escape from text editing → select the shape (handles appear).
 * - Escape from selection → deselect.
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
  const rootContext = React.useContext(RootContext);
  const frameRef = React.useRef<HTMLDivElement>(null);
  const [state, setState] = React.useState<InternalState>({ mode: "idle" });
  const [drawTool, setDrawTool] = React.useState<DrawTool | null>(null);

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
    state.mode === "selected" ||
    state.mode === "move" ||
    state.mode === "resize" ||
    state.mode === "rotate"
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
  const announcement = isTextMode
    ? `Editing text in ${selectedNode?.name || "shape"}`
    : selectedNodes.length === 1
      ? `${selectedNode?.name || "Shape"} selected`
      : selectedNodes.length > 1
        ? `${selectedNodes.length} shapes selected`
        : "";
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
  const resizeRects =
    state.mode === "resize" && grippedNode
      ? getResizedRects(
          selectedNodes,
          grippedNode,
          state.grip,
          state.dx,
          state.dy,
          state.lockAspect,
        )
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

  function commitEdit<T>(
    action: () => Promise<T>,
    onRollback?: () => void,
    onSuccess?: (result: T) => void,
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
    if (insertText && !restore && !insertTextNatively(textElement, nodeId, insertText)) {
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
   * When the caret isn't inside a run's text (an empty placeholder), a styled
   * span holding a zero-width space goes in first, so the native insertion
   * lands inside it and keeps the run's styling. Returns false without
   * `execCommand`, and the caller falls back to `insertTextAtCaret`.
   */
  function insertTextNatively(editingElement: HTMLElement, nodeId: string, data: string): boolean {
    const selection = window.getSelection();
    const anchor = selection?.anchorNode;
    if (
      !selection?.isCollapsed ||
      !anchor ||
      !editingElement.contains(anchor) ||
      typeof document.execCommand !== "function"
    ) {
      return false;
    }

    const isInRunText =
      anchor.nodeType === Node.TEXT_NODE &&
      Boolean(anchor.textContent) &&
      Boolean(anchor.parentElement?.closest(`[${PPTX_ATTRS.run}]`));
    if (isInRunText) return document.execCommand("insertText", false, data);

    const host = insertRunHost(editingElement, nodeId, "\u200B");
    setCaret(host, host.length);
    if (!document.execCommand("insertText", false, data)) {
      host.parentElement?.remove();
      return false;
    }
    removePlaceholderBreak(host, data);
    return true;
  }

  /**
   * Puts a styled run span holding `text` at the caret: inside the empty run
   * span the caret sits in, or a clone of the shape's template span (else a
   * plain span inheriting the paragraph's styling) in the caret's paragraph.
   * Returns the span's text node.
   */
  function insertRunHost(editingElement: HTMLElement, nodeId: string, text: string): Text {
    const selection = window.getSelection()!;
    const anchor = selection.anchorNode!;
    const anchorElement =
      anchor instanceof Element ? (anchor as HTMLElement) : anchor.parentElement;
    const textNode = document.createTextNode(text);

    const runSpan = anchorElement?.closest<HTMLElement>(`[${PPTX_ATTRS.run}]`);
    if (runSpan) {
      runSpan.appendChild(textNode);
      return textNode;
    }

    const template = runTemplateFor(nodeId);
    const span = template
      ? (template.cloneNode(false) as HTMLElement)
      : document.createElement("span");
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
    return textNode;
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
    // any, e.g. a plain rectangle you just started typing into): wrap the
    // insertion in a run span. Read-back emits it without a sourceRun and the
    // edit falls back to paragraph defaults.
    const textNode = insertRunHost(editingElement, nodeId, data);
    removePlaceholderBreak(textNode, data);
    setCaret(textNode, textNode.length);
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

    // Insert through the browser's editing pipeline so the keystroke stays
    // on its undo stack; fall back to placing it by hand.
    if (
      insertTextNatively(editingElement, nodeId, event.data) ||
      insertTextAtCaret(editingElement, nodeId, event.data)
    ) {
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

  /**
   * An empty text box goes away once editing moves elsewhere, as in
   * PowerPoint and Google Slides. Returns whether it was removed.
   */
  function removeIfEmptyTextBox(current: Extract<InternalState, { mode: "text" }>): boolean {
    const { nodeId, editingElement } = current;
    const node = slide?.nodes.find((n) => n.id === nodeId);
    if (!slideId || node?.nodeType !== "shape" || !(node as ShapeNodeData).isTextBox) return false;
    const readBack = readBackTextBody(editingElement);
    if (readBack.some((paragraph) => paragraph.runs.some((run) => run.text.length > 0))) {
      return false;
    }
    editingElement.contentEditable = "false";
    editingElement.style.cursor = "";
    commitEdit(
      () => store.edit({ type: "deleteNode", slideId, nodeId }),
      undefined,
      () => onNodeDelete?.(nodeId),
      (error) => onNodeDelete?.(nodeId, error),
    );
    return true;
  }

  function doExitTextMode(
    current: Extract<InternalState, { mode: "text" }>,
    nextNodeId?: string | null,
  ): void {
    if (nextNodeId === undefined || !removeIfEmptyTextBox(current)) commitTextEdits(current);

    if (nextNodeId === null) {
      setState({ mode: "idle" });
    } else {
      setState({ mode: "selected", nodeIds: [nextNodeId ?? current.nodeId] });
    }
    rootRef.current?.focus({ preventScroll: true });
  }

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
      // Ids of nodes still on the slide: an undo can remove a selected copy.
      nodeIds: selectedNodesRef.current.map((node) => node.id),
      textStyle: readTextStyle(),
      drawTool,
    });
  }

  function startDrawing(tool: DrawTool | null): void {
    const current = stateRef.current;
    if (current.mode === "text") commitTextEdits(current);
    setState({ mode: "idle" });
    setDrawTool(tool);
    rootRef.current?.focus({ preventScroll: true });
  }

  /** Places what a draw gesture drew: dragged out to `rect`, or clicked at `at`. */
  async function insertDrawn(tool: DrawTool, rect: Rect | null, at: NodePosition): Promise<void> {
    if (!slideId) return;
    const placement = rect
      ? { position: { x: rect.x, y: rect.y }, size: { w: rect.w, h: rect.h } }
      : { position: at };
    const { createdNodeIds } = await store.edit(
      tool.type === "textBox"
        ? { type: "addTextBox", slideId, ...placement }
        : { type: "addShape", slideId, geometry: tool.geometry, ...placement },
    );
    const nodeId = createdNodeIds?.[0];
    if (!nodeId) return;
    if (tool.type === "textBox") editTextRef.current(nodeId);
    else selectNodesRef.current([nodeId], true);
  }

  function editText(nodeId: string): void {
    const current = stateRef.current;
    if (current.mode === "text") commitTextEdits(current);
    setState({ mode: "selected", nodeIds: [nodeId] });
    resumeTextEditing(nodeId);
  }

  function selectNodes(nodeIds: string[], focus = false): void {
    const ids = nodeIds.filter((id) => slide?.nodes.some((node) => node.id === id));
    const current = stateRef.current;
    if (current.mode === "text") commitTextEdits(current);
    setState(ids.length > 0 ? { mode: "selected", nodeIds: ids } : { mode: "idle" });
    if (focus) rootRef.current?.focus({ preventScroll: true });
  }

  /**
   * In text mode, uncommitted typing is committed in the same batch, so the
   * formatting applies to the text as shown and undoes in one step.
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
    const domRange = getEditingDomRange(container);
    const caretRange = domRange ? getTextRange(container, domRange) : null;

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
      const text = paragraph ? getParagraphText(paragraph) : "";
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
  const removeIfEmptyTextBoxRef = useLatestRef(removeIfEmptyTextBox);
  const publishSelectionRef = useLatestRef(publishSelection);
  const selectNodesRef = useLatestRef(selectNodes);
  const editTextRef = useLatestRef(editText);
  const startDrawingRef = useLatestRef(startDrawing);
  const duplicateSelectionRef = useLatestRef(duplicateSelection);
  const deleteSelectionRef = useLatestRef(deleteSelection);
  const reorderSelectionRef = useLatestRef(reorderSelection);
  const applyFormatRef = useLatestRef(applyFormat);

  React.useEffect(
    () =>
      selectionChannel.register({
        select: (nodeIds, focus) => selectNodesRef.current(nodeIds, focus),
        editText: (nodeId) => editTextRef.current(nodeId),
        startDrawing: (tool) => startDrawingRef.current(tool),
        duplicate: () => duplicateSelectionRef.current(),
        remove: () => deleteSelectionRef.current(),
        reorder: (order) => reorderSelectionRef.current(order),
        setTextStyle: (style) => applyFormatRef.current("text", style),
        setParagraphStyle: (style) => applyFormatRef.current("paragraph", style),
      }),
    [
      selectionChannel,
      selectNodesRef,
      editTextRef,
      startDrawingRef,
      duplicateSelectionRef,
      deleteSelectionRef,
      reorderSelectionRef,
      applyFormatRef,
    ],
  );

  // Publish to `useSelection` when the selection, the mode, or the slide's
  // content changes. An edit replaces the slide DOM in SlideImpl's effect,
  // which runs after this one (it is a parent), so the text style is read
  // again a frame later from the fresh DOM.
  React.useEffect(() => {
    publishSelectionRef.current();
    const frameId = requestAnimationFrame(() => publishSelectionRef.current());
    return () => cancelAnimationFrame(frameId);
  }, [selectionKey, mode, slideRevision, drawTool, publishSelectionRef]);

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
        if (!removeIfEmptyTextBoxRef.current(currentState)) {
          commitTextEditsRef.current(currentState);
        }
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
  }, [isTextMode, stateRef, commitTextEditsRef, doExitTextModeRef, removeIfEmptyTextBoxRef]);

  const copySelectionRef = useLatestRef(copySelection);
  const pasteClipboardRef = useLatestRef(pasteClipboard);

  // Text mode keeps the browser's own clipboard for the text being edited.
  React.useEffect(() => {
    if (isTextMode) return;
    const onCopy = (event: ClipboardEvent) => copySelectionRef.current(event, false);
    const onCut = (event: ClipboardEvent) => copySelectionRef.current(event, true);
    const onPaste = (event: ClipboardEvent) => pasteClipboardRef.current(event);
    document.addEventListener("copy", onCopy);
    document.addEventListener("cut", onCut);
    document.addEventListener("paste", onPaste);
    return () => {
      document.removeEventListener("copy", onCopy);
      document.removeEventListener("cut", onCut);
      document.removeEventListener("paste", onPaste);
    };
  }, [isTextMode, copySelectionRef, pasteClipboardRef]);

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
    if (drawTool && !isTextMode) {
      if (event.button !== 0) {
        setDrawTool(null);
        return;
      }
      event.preventDefault();
      clearDocumentSelection();
      rootRef.current?.setPointerCapture(event.pointerId);
      rootRef.current?.focus();
      setState({
        mode: "draw",
        tool: drawTool,
        startX: event.clientX,
        startY: event.clientY,
        curX: event.clientX,
        curY: event.clientY,
      });
      return;
    }
    if (event.button === 2 && !isTextMode) {
      // PowerPoint: right-clicking a shape outside the selection selects it, so
      // a context menu acts on what was clicked. Empty canvas clears it.
      const nodeId = hitTest(event.clientX, event.clientY);
      if (!nodeId) setState({ mode: "idle" });
      else if (!selectedIds.includes(nodeId)) setState({ mode: "selected", nodeIds: [nodeId] });
      return;
    }
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
    // duplicating a shape, which isn't supported, so this press must not
    // become a move.
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
      wasSelected: selectedIds.length === 1 && selectedIds[0] === nodeId,
    });
  }

  function onGripPointerDown(
    event: React.PointerEvent<HTMLDivElement>,
    grip: GripDirection,
    nodeId: string,
  ): void {
    debugLog("grip pointerdown", { grip, nodeId, mode: state.mode, button: event.button });
    if (event.button !== 0 || !selectedIds.includes(nodeId)) return;
    event.stopPropagation();
    // See onPointerDown: stop selection extension / native drag initiation.
    event.preventDefault();
    // Handles show while editing text, as in Google Slides: pressing one ends it.
    if (state.mode === "text") commitTextEdits(state);
    clearDocumentSelection();
    rootRef.current?.setPointerCapture(event.pointerId);
    rootRef.current?.focus({ preventScroll: true });
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
   * Shape transforms captured when a rotate drag starts. The drag previews by
   * rotating the shape elements directly; cancelling puts these back.
   */
  const rotateOriginRef = React.useRef<Map<string, string>>(new Map());

  function onRotatePointerDown(event: React.PointerEvent<HTMLDivElement>, nodeId: string): void {
    if (event.button !== 0 || !selectedIds.includes(nodeId)) return;
    const node = selectedNodes.find((selected) => selected.id === nodeId);
    const frame = frameRef.current?.getBoundingClientRect();
    if (!node || !frame) return;
    event.stopPropagation();
    // See onPointerDown: stop selection extension / native drag initiation.
    event.preventDefault();
    if (state.mode === "text") commitTextEdits(state);
    clearDocumentSelection();
    rootRef.current?.setPointerCapture(event.pointerId);
    rootRef.current?.focus({ preventScroll: true });

    rotateOriginRef.current = new Map(
      selectedNodes.map((selected) => [
        selected.id,
        getShapeElement(selected.id)?.style.transform ?? "",
      ]),
    );
    const rect = getNodeRect(node);
    const centerX = frame.left + (rect.x + rect.w / 2) * zoom;
    const centerY = frame.top + (rect.y + rect.h / 2) * zoom;
    setState({
      mode: "rotate",
      nodeIds: selectedIds,
      gripNodeId: nodeId,
      centerX,
      centerY,
      startAngle: getPointerAngle(centerX, centerY, event.clientX, event.clientY),
      delta: 0,
    });
  }

  /** Previews a rotation on the shape elements, keeping their flips. */
  function previewRotation(delta: number): void {
    for (const node of selectedNodes) {
      const shapeElement = getShapeElement(node.id);
      if (!shapeElement) continue;
      const original = rotateOriginRef.current.get(node.id) ?? "";
      const flips = original.replace(/rotate\([^)]*\)\s*/, "");
      shapeElement.style.transform =
        `rotate(${normalizeAngle(node.rotation + delta)}deg) ${flips}`.trim();
    }
  }

  function restoreRotation(): void {
    for (const [nodeId, transform] of rotateOriginRef.current) {
      const shapeElement = getShapeElement(nodeId);
      if (shapeElement) shapeElement.style.transform = transform;
    }
  }

  /**
   * Pressing the frame of a shape being edited ends the editing, like Google
   * Slides: a click leaves the shape selected, and a drag moves it.
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
    });
  }

  /** What a drag can snap to: the slide, and every shape not being dragged. */
  function getSnapTargets(draggedIds: string[]): Rect[] {
    const targets: Rect[] = [{ x: 0, y: 0, w: presentation!.width, h: presentation!.height }];
    for (const node of slide!.nodes) {
      if (!draggedIds.includes(node.id)) targets.push(getNodeBounds(node));
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

    if (state.mode === "draw") {
      let curX = event.clientX;
      let curY = event.clientY;
      if (event.shiftKey) {
        const side = Math.max(Math.abs(curX - state.startX), Math.abs(curY - state.startY));
        curX = state.startX + Math.sign(curX - state.startX || 1) * side;
        curY = state.startY + Math.sign(curY - state.startY || 1) * side;
      }
      setState({ ...state, curX, curY });
      return;
    }

    if (state.mode === "rotate") {
      const grippedNode = selectedNodes.find((node) => node.id === state.gripNodeId);
      if (!grippedNode) return;
      const turned =
        getPointerAngle(state.centerX, state.centerY, event.clientX, event.clientY) -
        state.startAngle;
      // Shift steps the gripped shape's angle by 15°, as in PowerPoint.
      let angle = grippedNode.rotation + turned;
      if (event.shiftKey) angle = Math.round(angle / ROTATION_STEP) * ROTATION_STEP;
      const delta = angle - grippedNode.rotation;
      previewRotation(delta);
      setState({ ...state, delta });
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
      const bounds = getUnionRect(selectedNodes.map(getNodeBounds));
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
        // Snapping lines up slide-axis edges, which a rotated shape doesn't have.
        isSnapping && grippedNode && grippedNode.rotation === 0
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
    if (state.mode === "draw") {
      const { tool, startX, startY, curX, curY } = state;
      setDrawTool(null);
      setState({ mode: "idle" });
      const a = clientToSlide(startX, startY);
      const b = clientToSlide(curX, curY);
      const rect: Rect | null = isBandDrag(startX, startY, curX, curY)
        ? {
            x: Math.min(a.x, b.x),
            y: Math.min(a.y, b.y),
            w: Math.abs(a.x - b.x),
            h: Math.abs(a.y - b.y),
          }
        : null;
      runShortcut(() => insertDrawn(tool, rect, a));
      return;
    }
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
          const nodeRect = getNodeBounds(node);
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
      const { nodeIds, primaryId, dx, dy, moved, clickIds, wasSelected } = state;

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

        // Like Google Slides, a click selects; clicking the selected shape
        // again edits its text (as do a double-click and typing). Clicking a
        // member of a multi-selection without dragging collapses the
        // selection to it.
        if (clickIds) {
          // Shift/Ctrl+click: toggle only, never open the caret.
          setState(
            clickIds.length > 0 ? { mode: "selected", nodeIds: clickIds } : { mode: "idle" },
          );
        } else if (wasSelected && nodeHasText(primaryId)) {
          enterTextMode(primaryId, state.startX, state.startY);
        } else {
          setState({ mode: "selected", nodeIds: [primaryId] });
        }
      }
    } else if (state.mode === "rotate") {
      const { nodeIds, delta } = state;
      const rotatingNodes = selectedNodes;
      setState({ mode: "selected", nodeIds });
      if (normalizeAngle(delta) === 0) {
        restoreRotation();
        return;
      }
      const operations = rotatingNodes.map((node) => ({
        type: "setNodeTransform" as const,
        slideId: slideId!,
        nodeId: node.id,
        rotation: normalizeAngle(node.rotation + delta),
      }));
      commitEdit(
        () => store.edit(operations.length === 1 ? operations[0] : { type: "batch", operations }),
        restoreRotation,
        () => {
          for (const id of nodeIds) onNodeTransform?.(id);
        },
        (error) => {
          for (const id of nodeIds) onNodeTransform?.(id, error);
        },
      );
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
        const rects = getResizedRects(resizingNodes, grippedNode, grip, dx, dy, lockAspect);
        const transforms = [...rects].map(([nodeId, rect]) => ({ nodeId, rect }));
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
    } else if (state.mode === "rotate") {
      restoreRotation();
      setState({ mode: "selected", nodeIds: state.nodeIds });
    } else if (state.mode === "draw") {
      setState({ mode: "idle" });
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

  /** Runs an action from a shortcut, where nobody awaits it. */
  function runShortcut(action: () => Promise<void>): void {
    action().catch((err) => console.warn("[pptx] edit failed:", err));
  }

  async function deleteSelection(): Promise<void> {
    const nodeIds = selectedNodesRef.current.map((n) => n.id);
    if (nodeIds.length === 0 || !slideId) return;
    try {
      await store.edit(
        nodeIds.length === 1
          ? { type: "deleteNode", slideId, nodeId: nodeIds[0] }
          : {
              type: "batch",
              operations: nodeIds.map((nodeId) => ({
                type: "deleteNode" as const,
                slideId,
                nodeId,
              })),
            },
      );
    } catch (error) {
      for (const nodeId of nodeIds) onNodeDelete?.(nodeId, error);
      throw error;
    }
    setState({ mode: "idle" });
    for (const nodeId of nodeIds) onNodeDelete?.(nodeId);
  }

  /**
   * Whether a clipboard event is the deck's: aimed at the presentation or an
   * editor toolbar, and not at a field with its own clipboard.
   */
  function getIsDeckClipboardEvent(event: ClipboardEvent): boolean {
    if (getIsNativeUndoTarget(event.target)) return false;
    return getIsPresentationTarget(event.target, rootContext?.current ?? rootRef.current);
  }

  function copySelection(event: ClipboardEvent, cut: boolean): void {
    const nodes = selectedNodesRef.current;
    if (!presentation || !slideId || nodes.length === 0 || !event.clipboardData) return;
    if (!getIsDeckClipboardEvent(event)) return;
    event.preventDefault();

    const token = stashNodes(
      presentation,
      slideId,
      copyNodes(
        presentation,
        slideId,
        nodes.map((node) => node.id),
      ),
    );
    const text = nodes
      .flatMap((node) =>
        node.nodeType === "shape"
          ? ((node as ShapeNodeData).textBody?.paragraphs.map((paragraph) =>
              paragraph.runs.map((run) => run.text).join(""),
            ) ?? [])
          : [],
      )
      .join("\n");
    writeClipboardToken(event.clipboardData, token, text);
    if (cut) runShortcut(deleteSelection);
  }

  /** Pastes shapes copied in this page, or else an image file from the system clipboard. */
  function pasteClipboard(event: ClipboardEvent): void {
    const data = event.clipboardData;
    if (!presentation || !slideId || !data || !getIsDeckClipboardEvent(event)) return;
    const selectCreated = (result: EditResult) => {
      if (result.createdNodeIds?.length) {
        setState({ mode: "selected", nodeIds: result.createdNodeIds });
      }
    };

    const token = readClipboardToken(data);
    const stashed = token ? takeStash(token, presentation, slideId) : null;
    if (stashed) {
      event.preventDefault();
      commitEdit(
        async () =>
          store.edit({
            type: "pasteNodes",
            slideId,
            clipboard: await stashed.clipboard,
            offset: { x: stashed.offset, y: stashed.offset },
          }),
        undefined,
        selectCreated,
      );
      return;
    }

    const image = Array.from(data.files).find((file) => PASTABLE_IMAGE_TYPES.has(file.type));
    if (!image) return;
    event.preventDefault();
    commitEdit(
      async () =>
        store.edit({
          type: "addImage",
          slideId,
          data: new Uint8Array(await image.arrayBuffer()),
          contentType: image.type,
          name: image.name || undefined,
        }),
      undefined,
      selectCreated,
    );
  }

  /** PowerPoint's Ctrl+D: copies land just above their originals, selected. */
  async function duplicateSelection(): Promise<void> {
    if (!slideId) return;
    const operations = selectedNodesRef.current.map((node) => ({
      type: "duplicateNode" as const,
      slideId,
      nodeId: node.id,
    }));
    if (operations.length === 0) return;
    const result = await store.edit(
      operations.length === 1 ? operations[0] : { type: "batch", operations },
    );
    if (result.createdNodeIds?.length) {
      setState({ mode: "selected", nodeIds: result.createdNodeIds });
    }
  }

  async function reorderSelection(order: NodeOrder): Promise<void> {
    if (!slide || !slideId) return;
    // Move the selected nodes one at a time from the end they travel toward,
    // so they keep their order among themselves.
    const stacked = [...selectedNodesRef.current].sort(
      (a, b) => slide.nodes.indexOf(a) - slide.nodes.indexOf(b),
    );
    if (order === "back" || order === "forward") stacked.reverse();
    const operations = stacked.map((node) => ({
      type: "reorderNode" as const,
      slideId,
      nodeId: node.id,
      order,
    }));
    if (operations.length === 0) return;
    await store.edit(operations.length === 1 ? operations[0] : { type: "batch", operations });
  }

  function onKeyDown(event: React.KeyboardEvent<HTMLDivElement>): void {
    // Text mode keys are handled by the document listener; this handler
    // only triggers when the overlay div has focus (selected/idle modes).
    // Undo/redo shortcuts live in the outer Selection wrapper (document
    // level) so they survive the slide-change remount of this component.
    if (isTextMode) return;

    if (drawTool && event.key === "Escape") {
      event.preventDefault();
      setDrawTool(null);
      return;
    }

    const mod = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();

    // PowerPoint: Tab and Shift+Tab move through the slide's shapes. Past the
    // last one, focus moves on as usual, so the keyboard is never trapped.
    if (event.key === "Tab" && !mod && !event.altKey && slide && slide.nodes.length > 0) {
      const nodes = slide.nodes;
      const step = event.shiftKey ? -1 : 1;
      const current = isSoloSelection ? nodes.indexOf(selectedNodes[0]) : -1;
      const next = current === -1 ? (event.shiftKey ? nodes.length - 1 : 0) : current + step;
      if (next < 0 || next >= nodes.length) {
        setState({ mode: "idle" });
        return;
      }
      event.preventDefault();
      setState({ mode: "selected", nodeIds: [nodes[next].id] });
      return;
    }

    if (mod && key === "a") {
      event.preventDefault();
      const allIds = slide!.nodes.map((n) => n.id);
      if (allIds.length > 0) setState({ mode: "selected", nodeIds: allIds });
      return;
    }

    if (selectedNodes.length === 0) return;

    if (mod && key === "d") {
      event.preventDefault();
      runShortcut(duplicateSelection);
      return;
    }

    // Ctrl/Cmd+] and [ step through the stack; with Shift, to the front or back.
    if (mod && (event.code === "BracketRight" || event.code === "BracketLeft")) {
      event.preventDefault();
      const up = event.code === "BracketRight";
      const order = event.shiftKey ? (up ? "front" : "back") : up ? "forward" : "backward";
      runShortcut(() => reorderSelection(order));
      return;
    }

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
      runShortcut(deleteSelection);
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
          "aria-label": "Slide editor",
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
                    resizeRects,
                    showResizeHandles: !isBanding,
                    frameRef,
                    slideWidth: presentation.width,
                    slideHeight: presentation.height,
                    onGripPointerDown,
                    onBorderPointerDown,
                    onRotatePointerDown,
                  }}
                >
                  {children ?? DEFAULT_PARTS}
                </SelectionPartsContext.Provider>
              </div>
              <div role="status" aria-live="polite" style={VISUALLY_HIDDEN}>
                {announcement}
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
            // A drag captures the pointer onto the overlay, so the handle's
            // cursor has to come along or it reverts mid-drag.
            cursor:
              drawTool || state.mode === "draw"
                ? "crosshair"
                : state.mode === "rotate"
                  ? ROTATE_CURSOR
                  : state.mode === "resize"
                    ? getGripCursor(state.grip, grippedNode?.rotation ?? 0)
                    : undefined,
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

/**
 * The parts `Presentation.Selection` renders without children: the default
 * look, built from the same parts an app composes, themed through the
 * `--presentation-selection` and `--presentation-snap-guide` variables.
 */
const DEFAULT_PARTS = (
  <>
    <SelectionBox
      style={{
        // Centered on the edge, over the shape's own outline.
        outline: "2px solid var(--presentation-selection, #1a73e8)",
        outlineOffset: -1,
      }}
    >
      <ResizeHandle
        style={{
          width: 9,
          height: 9,
          background: "#fff",
          border: "1.5px solid var(--presentation-selection, #1a73e8)",
          borderRadius: 2,
        }}
      />
      <RotateHandle
        style={{
          width: 10,
          height: 10,
          background: "#fff",
          border: "1.5px solid var(--presentation-selection, #1a73e8)",
          borderRadius: "50%",
        }}
      />
    </SelectionBox>
    <SnapGuide style={{ background: "var(--presentation-snap-guide, #f43f5e)" }} />
    <Marquee
      style={{
        border: "1px solid var(--presentation-selection, #1a73e8)",
        background: "color-mix(in srgb, var(--presentation-selection, #1a73e8) 10%, transparent)",
      }}
    />
  </>
);

export namespace Selection {
  export type State = SelectionState;
  export type Props = SelectionProps;
}
