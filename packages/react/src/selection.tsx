import * as React from "react";

import type { Position, SetTextBodyParagraph, ShapeNodeData, SlideNode } from "@diceui/pptx-parser";

import { usePresentation, usePresentationStore, useSlide, useZoom } from "./context";
import type { RenderProp } from "./render";
import { mergeRefs, renderElement } from "./render";

const SELECTION_NAME = "PresentationSelection";

/** Minimum shape size (slide px) a resize can shrink to. */
const MIN_SIZE = 8;
/** Screen-px movement before a pointer-down becomes a drag instead of a click. */
const DRAG_THRESHOLD = 3;

const HANDLE_DIRECTIONS = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;
type HandleDirection = (typeof HANDLE_DIRECTIONS)[number];

const HANDLE_CURSORS: Record<HandleDirection, string> = {
  nw: "nwse-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  e: "ew-resize",
  se: "nwse-resize",
  s: "ns-resize",
  sw: "nesw-resize",
  w: "ew-resize",
};

// Internal discriminated union — not part of the public API.
type InternalState =
  | { mode: "idle" }
  | { mode: "selected"; nodeId: string }
  | {
      mode: "move";
      nodeId: string;
      startX: number;
      startY: number;
      dx: number;
      dy: number;
      /** True once movement exceeded the drag threshold; gates the commit. */
      moved: boolean;
      /** Move started from text-mode border drag: return to text mode after. */
      resumeText?: boolean;
    }
  | {
      mode: "resize";
      nodeId: string;
      handle: HandleDirection;
      startX: number;
      startY: number;
      dx: number;
      dy: number;
    }
  | {
      mode: "text";
      nodeId: string;
      /** The text container element that has contentEditable. */
      editingEl: HTMLElement;
    };

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Apply a resize drag (slide-px deltas) to the original rect, clamped to MIN_SIZE. */
function resizeRect(origin: Rect, handle: HandleDirection, dx: number, dy: number): Rect {
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

function nodeRect(node: SlideNode): Rect {
  return { x: node.position.x, y: node.position.y, w: node.size.w, h: node.size.h };
}

// ---------------------------------------------------------------------------
// DOM read-back: contentEditable → setTextBody payload
// ---------------------------------------------------------------------------

/**
 * Walk the edited contentEditable container and produce a `setTextBody`
 * paragraphs payload. Uses `data-pptx-p` / `data-pptx-r` to map back to
 * source paragraph and run indices for style inheritance.
 */
function readBackTextBody(container: HTMLElement): SetTextBodyParagraph[] {
  const paragraphs: SetTextBodyParagraph[] = [];

  const paraDivs = Array.from(container.children).filter(
    (el) => el instanceof HTMLElement,
  ) as HTMLElement[];

  const effectiveDivs = paraDivs.length > 0 ? paraDivs : [container];
  let lastSourceP = 0;

  for (const paraDiv of effectiveDivs) {
    const srcPStr = paraDiv.dataset?.pptxP;
    const sourceParagraphIndex = srcPStr !== undefined ? Number(srcPStr) : lastSourceP;
    lastSourceP = sourceParagraphIndex;

    const runs = readRunsFromParagraphDiv(paraDiv, sourceParagraphIndex);
    paragraphs.push({ sourceParagraphIndex, runs });
  }

  return paragraphs;
}

function readRunsFromParagraphDiv(
  paraDiv: HTMLElement,
  defaultSourceP: number,
): SetTextBodyParagraph["runs"] {
  const runs: SetTextBodyParagraph["runs"] = [];
  let lastSourceR: [number, number] | undefined;

  for (const child of Array.from(paraDiv.childNodes)) {
    if (child instanceof HTMLElement && child.dataset.pptxBullet !== undefined) {
      continue;
    }

    if (child instanceof HTMLElement && child.dataset.pptxR !== undefined) {
      const runIdx = Number(child.dataset.pptxR);
      const sourceRun: [number, number] = [defaultSourceP, runIdx];
      const text = child.textContent ?? "";
      if (text.length > 0) {
        runs.push({ text, sourceRun });
        lastSourceR = sourceRun;
      }
    } else if (child instanceof HTMLBRElement) {
      // Browsers insert <br> for empty paragraphs; skip.
    } else if (child.nodeType === Node.TEXT_NODE) {
      const text = child.textContent ?? "";
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceR });
      }
    } else if (child instanceof HTMLElement) {
      const text = child.textContent ?? "";
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceR });
      }
    }
  }

  if (runs.length === 0) {
    runs.push({ text: "", sourceRun: lastSourceR });
  }

  return runs;
}

/** Compare the read-back paragraphs to the model to detect changes. */
function textBodyChanged(node: SlideNode, readBack: SetTextBodyParagraph[]): boolean {
  if (node.nodeType !== "shape") return false;
  const shape = node as ShapeNodeData;
  const paragraphs = shape.textBody?.paragraphs;
  if (!paragraphs) return false;
  if (paragraphs.length !== readBack.length) return true;
  for (let i = 0; i < paragraphs.length; i++) {
    const origRuns = paragraphs[i].runs;
    const newRuns = readBack[i].runs;
    if (origRuns.length !== newRuns.length) return true;
    for (let j = 0; j < origRuns.length; j++) {
      if (origRuns[j].text !== newRuns[j].text) return true;
    }
  }
  return false;
}

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** The current interaction state of the edit layer. */
export interface SelectionState {
  /** Interaction mode of the selection. */
  mode: "idle" | "selected" | "move" | "resize" | "text";
  /** The slide node currently selected, or `null` when nothing is selected. */
  selectedNode: SlideNode | null;
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
 * - Drag any shape → move it.
 * - Escape from text editing → select the shape (handles appear).
 * - Escape from selection → deselect.
 * - Click a non-text shape → select it.
 * - Arrow keys nudge, Delete removes (while shape is selected, not editing text).
 */
export const Selection = React.forwardRef<HTMLDivElement, SelectionProps>(function Selection(
  { render, onUndo, onRedo, onNodeDelete, onNodeTransform, onTextChange, ...selectionProps },
  forwardedRef,
) {
  const store = usePresentationStore(SELECTION_NAME);
  const { presentation } = usePresentation();
  const { slide, slideId } = useSlide();
  const { zoom } = useZoom();

  const rootRef = React.useRef<HTMLDivElement>(null);
  const [state, setState] = React.useState<InternalState>({ mode: "idle" });

  // Stable refs for document-level listeners (avoids stale closures).
  const stateRef = React.useRef(state);
  stateRef.current = state;

  const selectedNode =
    state.mode !== "idle" && slide
      ? (slide.nodes.find((n) => n.id === state.nodeId) ?? null)
      : null;

  const isTextMode = state.mode === "text";
  const publicState: SelectionState = { mode: state.mode, selectedNode };

  if (!presentation?.pkg || !slide || !slideId) return null;

  // --- DOM helpers ---

  function slideWrapper(): HTMLElement | null {
    return rootRef.current?.parentElement?.parentElement ?? null;
  }

  function shapeElement(nodeId: string): HTMLElement | null {
    return (
      slideWrapper()?.querySelector<HTMLElement>(`[data-pptx-node-id="${CSS.escape(nodeId)}"]`) ??
      null
    );
  }

  function textContainerOf(shapeEl: HTMLElement): HTMLElement | null {
    const firstPara = shapeEl.querySelector("[data-pptx-p]");
    return (firstPara?.parentElement as HTMLElement) ?? null;
  }

  function hitTest(clientX: number, clientY: number): string | null {
    const root = rootRef.current;
    const wrapper = slideWrapper();
    if (!root || !wrapper) return null;
    for (const el of document.elementsFromPoint(clientX, clientY)) {
      if (root.contains(el)) continue;
      if (!wrapper.contains(el)) continue;
      const nodeEl = (el as HTMLElement).closest<HTMLElement>("[data-pptx-node-id]");
      if (nodeEl) return nodeEl.getAttribute("data-pptx-node-id");
    }
    return null;
  }

  function nodeHasText(nodeId: string): boolean {
    const node = slide!.nodes.find((n) => n.id === nodeId);
    return node?.nodeType === "shape" && Boolean((node as ShapeNodeData).textBody);
  }

  function commitEdit(
    action: () => Promise<unknown>,
    onRollback?: () => void,
    onSuccess?: () => void,
    onFailure?: (error: unknown) => void,
  ): void {
    action().then(onSuccess, (err) => {
      onRollback?.();
      if (onFailure) {
        onFailure(err);
      } else {
        console.warn("[pptx] edit failed:", err);
      }
    });
  }

  // --- Text editing helpers ---

  function enterTextMode(nodeId: string, clientX?: number, clientY?: number): void {
    const shapeEl = shapeElement(nodeId);
    if (!shapeEl) return;
    const textEl = textContainerOf(shapeEl);
    if (!textEl) return;

    textEl.contentEditable = "plaintext-only";
    if (!textEl.isContentEditable) {
      textEl.contentEditable = "true";
    }
    textEl.style.cursor = "text";
    textEl.style.outline = "none";

    // The caret-from-point APIs hit-test the DOM, and the overlay still
    // covers the text at this moment (React hasn't re-rendered with
    // pointerEvents: none yet). Disable it directly so the point resolves
    // into the text; the state render below keeps it disabled.
    if (rootRef.current) rootRef.current.style.pointerEvents = "none";

    textEl.focus({ preventScroll: true });

    if (clientX !== undefined && clientY !== undefined) {
      placeCaretAtPoint(clientX, clientY);
    } else {
      placeCaretAtEnd(textEl);
    }

    setState({ mode: "text", nodeId, editingEl: textEl });
  }

  /**
   * Re-enter text mode after a commit re-rendered the slide. The store
   * notify flushes React state at microtask end and the new slide DOM is
   * committed before the next paint, so a double rAF (not a timeout) is the
   * earliest reliable moment the replacement shape element exists.
   */
  function resumeTextEditing(nodeId: string): void {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        enterTextMode(nodeId);
      });
    });
  }

  function placeCaretAtEnd(textEl: HTMLElement): void {
    try {
      const sel = window.getSelection();
      if (!sel) return;
      const range = document.createRange();
      range.selectNodeContents(textEl);
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    } catch {
      // Ignore caret placement failures.
    }
  }

  function placeCaretAtPoint(clientX: number, clientY: number): void {
    try {
      const sel = window.getSelection();
      if (!sel) return;
      // Standard API (Firefox, newer Chrome/Safari).
      const docWithCaret = document as Document & {
        caretPositionFromPoint?: (
          x: number,
          y: number,
        ) => { offsetNode: Node; offset: number } | null;
      };
      if (docWithCaret.caretPositionFromPoint) {
        const pos = docWithCaret.caretPositionFromPoint(clientX, clientY);
        if (pos) {
          const range = document.createRange();
          range.setStart(pos.offsetNode, pos.offset);
          range.collapse(true);
          sel.removeAllRanges();
          sel.addRange(range);
          return;
        }
      }
      // Legacy WebKit API.
      if (document.caretRangeFromPoint) {
        const range = document.caretRangeFromPoint(clientX, clientY);
        if (range) {
          sel.removeAllRanges();
          sel.addRange(range);
        }
      }
    } catch {
      // Ignore caret placement failures.
    }
  }

  /** Tear down contentEditable and commit the edited text if it changed. */
  function commitTextEdits(current: Extract<InternalState, { mode: "text" }>): void {
    const { nodeId, editingEl } = current;

    editingEl.contentEditable = "false";
    editingEl.style.cursor = "";

    const node = slide!.nodes.find((n) => n.id === nodeId);
    if (!node) return;
    const readBack = readBackTextBody(editingEl);
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
      setState({ mode: "selected", nodeId: nextNodeId ?? current.nodeId });
    }
    rootRef.current?.focus({ preventScroll: true });
  }

  // --- Document-level listeners for text mode ---
  // When in text mode the overlay has pointerEvents: none, so we must listen
  // on the document for clicks-outside and Escape.
  React.useEffect(() => {
    if (!isTextMode) return;

    function onDocPointerDown(e: PointerEvent): void {
      const cur = stateRef.current;
      if (cur.mode !== "text") return;

      // Click inside the editing element → let contentEditable handle it.
      if (cur.editingEl.contains(e.target as Node)) return;

      // Clicks on the overlay's own children (border move strips) are
      // handled by their own handlers.
      if (rootRef.current?.contains(e.target as Node)) return;

      // Click on another shape?
      const wrapper = rootRef.current?.parentElement?.parentElement;
      if (!wrapper) return;
      const target = e.target as HTMLElement;
      const nodeEl = target.closest?.("[data-pptx-node-id]") as HTMLElement | null;
      const hitNodeId =
        nodeEl && wrapper.contains(nodeEl) ? nodeEl.getAttribute("data-pptx-node-id") : null;

      if (hitNodeId && hitNodeId !== cur.nodeId) {
        // Clicked a different shape: exit text, select the new one.
        doExitTextMode(cur, hitNodeId);
      } else if (!hitNodeId) {
        // Clicked empty area: exit text, deselect.
        doExitTextMode(cur, null);
      }
      // Else clicked inside the same shape but outside the text container
      // (e.g. shape padding area) — stay in text mode.
    }

    function onDocKeyDown(e: KeyboardEvent): void {
      const cur = stateRef.current;
      if (cur.mode !== "text") return;

      if (e.key === "Escape") {
        e.preventDefault();
        // Escape → select the shape (PowerPoint behavior).
        doExitTextMode(cur);
        return;
      }
    }

    document.addEventListener("pointerdown", onDocPointerDown, true);
    document.addEventListener("keydown", onDocKeyDown, true);
    return () => {
      document.removeEventListener("pointerdown", onDocPointerDown, true);
      document.removeEventListener("keydown", onDocKeyDown, true);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- stateRef is stable
  }, [isTextMode]);

  // --- Overlay pointer events (non-text modes) ---

  function onPointerDown(event: React.PointerEvent<HTMLDivElement>): void {
    if (event.button !== 0 || isTextMode) return;

    const nodeId = hitTest(event.clientX, event.clientY);
    if (!nodeId) {
      setState({ mode: "idle" });
      return;
    }

    // All shapes start in move mode on pointer-down. On pointer-up we check
    // if the user actually dragged. If they didn't (it was a click), text
    // shapes enter text mode and non-text shapes enter selected mode.
    rootRef.current?.setPointerCapture(event.pointerId);
    rootRef.current?.focus();
    setState({
      mode: "move",
      nodeId,
      startX: event.clientX,
      startY: event.clientY,
      dx: 0,
      dy: 0,
      moved: false,
    });
  }

  function onHandlePointerDown(
    event: React.PointerEvent<HTMLDivElement>,
    handle: HandleDirection,
  ): void {
    if (event.button !== 0 || state.mode === "idle" || isTextMode) return;
    event.stopPropagation();
    rootRef.current?.setPointerCapture(event.pointerId);
    setState({
      mode: "resize",
      nodeId: state.nodeId,
      handle,
      startX: event.clientX,
      startY: event.clientY,
      dx: 0,
      dy: 0,
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
      nodeId,
      startX: event.clientX,
      startY: event.clientY,
      dx: 0,
      dy: 0,
      moved: false,
      resumeText: true,
    });
  }

  function onPointerMove(event: React.PointerEvent<HTMLDivElement>): void {
    if (state.mode !== "move" && state.mode !== "resize") return;
    const dx = (event.clientX - state.startX) / zoom;
    const dy = (event.clientY - state.startY) / zoom;

    if (state.mode === "move") {
      const moved =
        state.moved ||
        Math.hypot(event.clientX - state.startX, event.clientY - state.startY) > DRAG_THRESHOLD;
      if (moved) {
        const el = shapeElement(state.nodeId);
        if (el) el.style.translate = `${dx}px ${dy}px`;
      }
      setState({ ...state, dx, dy, moved });
    } else {
      setState({ ...state, dx, dy });
    }
  }

  function onPointerUp(): void {
    if (state.mode === "move") {
      const { nodeId, dx, dy, moved, resumeText } = state;

      if (moved && selectedNode) {
        // Actual drag → commit the move, land in selected mode.
        setState({ mode: "selected", nodeId });
        commitEdit(
          () =>
            store.edit({
              type: "setNodeTransform",
              slideId: slideId!,
              nodeId,
              position: { x: selectedNode.position.x + dx, y: selectedNode.position.y + dy },
            }),
          () => {
            const el = shapeElement(nodeId);
            if (el) el.style.translate = "";
          },
          () => {
            onNodeTransform?.(nodeId);
            // Border drag from text mode: return to editing so the user can
            // keep typing right where they left off (PowerPoint behavior).
            if (resumeText) resumeTextEditing(nodeId);
          },
          (error) => onNodeTransform?.(nodeId, error),
        );
      } else {
        // Click (no drag). Clear any stray preview translate.
        const el = shapeElement(nodeId);
        if (el) el.style.translate = "";

        // Text shapes → edit immediately. Non-text shapes → select.
        if (resumeText) {
          // Border click without drag: back to editing (caret at end). The
          // text commit may re-render the slide, so wait for the new DOM.
          resumeTextEditing(nodeId);
        } else if (nodeHasText(nodeId)) {
          // Use the original pointer position for caret placement.
          enterTextMode(nodeId, state.startX, state.startY);
        } else {
          setState({ mode: "selected", nodeId });
        }
      }
    } else if (state.mode === "resize") {
      const { nodeId, handle, dx, dy } = state;
      setState({ mode: "selected", nodeId });
      if (selectedNode && (dx !== 0 || dy !== 0)) {
        const next = resizeRect(nodeRect(selectedNode), handle, dx, dy);
        commitEdit(
          () =>
            store.edit({
              type: "setNodeTransform",
              slideId: slideId!,
              nodeId,
              position: { x: next.x, y: next.y },
              size: { w: next.w, h: next.h },
            }),
          undefined,
          () => onNodeTransform?.(nodeId),
          (error) => onNodeTransform?.(nodeId, error),
        );
      }
    }
  }

  function onPointerCancel(): void {
    if (state.mode === "move" || state.mode === "resize") {
      const el = shapeElement(state.nodeId);
      if (el) el.style.translate = "";
      setState({ mode: "selected", nodeId: state.nodeId });
    }
  }

  function nudge(node: SlideNode, delta: Position): void {
    commitEdit(
      () =>
        store.edit({
          type: "setNodeTransform",
          slideId: slideId!,
          nodeId: node.id,
          position: { x: node.position.x + delta.x, y: node.position.y + delta.y },
        }),
      undefined,
      () => onNodeTransform?.(node.id),
      (error) => onNodeTransform?.(node.id, error),
    );
  }

  function onKeyDown(event: React.KeyboardEvent<HTMLDivElement>): void {
    // Text mode keys are handled by the document listener; this handler
    // only fires when the overlay div has focus (selected/idle modes).
    if (isTextMode) return;

    const mod = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();

    // Undo / redo — active whenever the overlay is focused.
    if (mod && key === "z" && !event.shiftKey) {
      event.preventDefault();
      const success = store.undo();
      rootRef.current?.focus({ preventScroll: true });
      onUndo?.(success ? "success" : "empty");
      return;
    }
    if (mod && (key === "y" || (key === "z" && event.shiftKey))) {
      event.preventDefault();
      store
        .redo()
        .then((success) => {
          rootRef.current?.focus({ preventScroll: true });
          onRedo?.(success ? "success" : "empty");
        })
        .catch((error) => {
          rootRef.current?.focus({ preventScroll: true });
          onRedo?.("empty", error);
        });
      return;
    }

    if (!selectedNode) return;

    // F2 or Enter on a selected text shape enters text mode.
    if ((event.key === "F2" || event.key === "Enter") && nodeHasText(selectedNode.id)) {
      event.preventDefault();
      enterTextMode(selectedNode.id);
      return;
    }

    if (event.key === "Escape") {
      setState({ mode: "idle" });
      return;
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      const nodeId = selectedNode.id;
      commitEdit(
        () => store.edit({ type: "deleteNode", slideId: slideId!, nodeId }),
        undefined,
        () => onNodeDelete?.(nodeId),
        (error) => onNodeDelete?.(nodeId, error),
      );
      return;
    }

    const step = event.shiftKey ? 10 : 1;
    const arrows: Record<string, Position> = {
      ArrowLeft: { x: -step, y: 0 },
      ArrowRight: { x: step, y: 0 },
      ArrowUp: { x: 0, y: -step },
      ArrowDown: { x: 0, y: step },
    };
    const delta = arrows[event.key];
    if (delta) {
      event.preventDefault();
      nudge(selectedNode, delta);
    }
  }

  return renderElement(
    "div",
    { render },
    {
      state: publicState,
      ref: mergeRefs(rootRef, forwardedRef),
      props: [
        {
          "data-pptx-selection": "",
          "data-mode": state.mode,
          tabIndex: -1,
          onPointerDown,
          onPointerMove,
          onPointerUp,
          onPointerCancel,
          onKeyDown,
          children: selectedNode ? (
            <SelectionBox
              node={selectedNode}
              state={state}
              zoom={zoom}
              onHandlePointerDown={onHandlePointerDown}
              onBorderPointerDown={onBorderPointerDown}
            />
          ) : null,
          style: {
            position: "absolute",
            inset: 0,
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

// ---------------------------------------------------------------------------
// SelectionBox (internal)
// ---------------------------------------------------------------------------

/** Hitbox width (screen px) of the border move strips shown while editing text. */
const BORDER_GRAB_SIZE = 10;

interface SelectionBoxProps {
  node: SlideNode;
  state: InternalState;
  zoom: number;
  onHandlePointerDown: (event: React.PointerEvent<HTMLDivElement>, handle: HandleDirection) => void;
  onBorderPointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
}

function SelectionBox({
  node,
  state,
  zoom,
  onHandlePointerDown,
  onBorderPointerDown,
}: SelectionBoxProps) {
  let rect = nodeRect(node);
  if (state.mode === "move" && state.moved) {
    rect = { ...rect, x: rect.x + state.dx, y: rect.y + state.dy };
  } else if (state.mode === "resize") {
    rect = resizeRect(rect, state.handle, state.dx, state.dy);
  }

  const isTextMode = state.mode === "text";
  const showHandles = node.rotation === 0 && state.mode !== "move" && !isTextMode;

  const handlePositions: Record<HandleDirection, React.CSSProperties> = {
    nw: { left: 0, top: 0 },
    n: { left: "50%", top: 0 },
    ne: { left: "100%", top: 0 },
    e: { left: "100%", top: "50%" },
    se: { left: "100%", top: "100%" },
    s: { left: "50%", top: "100%" },
    sw: { left: 0, top: "100%" },
    w: { left: 0, top: "50%" },
  };

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
        boxShadow: `0 0 0 ${isTextMode ? "2" : "1.5"}px var(--pptx-selection, #2563eb)`,
        cursor: isTextMode ? "text" : "move",
        pointerEvents: "none",
      }}
    >
      {showHandles &&
        HANDLE_DIRECTIONS.map((direction) => (
          <div
            key={direction}
            data-resize-handle={direction}
            onPointerDown={(event) => onHandlePointerDown(event, direction)}
            style={{
              position: "absolute",
              ...handlePositions[direction],
              width: 9,
              height: 9,
              marginLeft: -4.5,
              marginTop: -4.5,
              background: "#fff",
              border: "1.5px solid var(--pptx-selection, #2563eb)",
              borderRadius: 2,
              cursor: HANDLE_CURSORS[direction],
              pointerEvents: "auto",
            }}
          />
        ))}
      {isTextMode && <BorderMoveStrips onPointerDown={onBorderPointerDown} />}
    </div>
  );
}

/**
 * Invisible grab strips straddling the shape border while editing text.
 * Dragging them moves the shape (PowerPoint: grab the frame to move).
 * The hitbox extends both inward and outward from the border line.
 */
function BorderMoveStrips({
  onPointerDown,
}: {
  onPointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
}) {
  const half = BORDER_GRAB_SIZE / 2;
  const strips: React.CSSProperties[] = [
    { left: -half, right: -half, top: -half, height: BORDER_GRAB_SIZE },
    { left: -half, right: -half, bottom: -half, height: BORDER_GRAB_SIZE },
    { left: -half, top: half, bottom: half, width: BORDER_GRAB_SIZE },
    { right: -half, top: half, bottom: half, width: BORDER_GRAB_SIZE },
  ];
  return (
    <>
      {strips.map((style, i) => (
        <div
          key={i}
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
