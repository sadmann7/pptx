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

  // The text container's children are paragraph divs (data-pptx-p) or
  // browser-inserted divs (from pressing Enter).
  const paraDivs = Array.from(container.children).filter(
    (el) => el instanceof HTMLElement,
  ) as HTMLElement[];

  // If the container has no child divs (e.g. all text was deleted and the
  // user typed directly into the container) treat the whole container as one
  // paragraph.
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
    // Skip bullet spans.
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
      // Browsers insert <br> for empty paragraphs; skip at end.
    } else if (child.nodeType === Node.TEXT_NODE) {
      const text = child.textContent ?? "";
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceR });
      }
    } else if (child instanceof HTMLElement) {
      // Browser-wrapped content (e.g. <span> without data-pptx-r).
      const text = child.textContent ?? "";
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceR });
      }
    }
  }

  // Ensure at least one empty run so the paragraph is not dropped.
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
 * - Click a shape to select it (topmost shape wins, like PowerPoint).
 * - Double-click (or F2/Enter) a text shape to edit text inline.
 * - Drag to move; drag the handles to resize (non-rotated shapes).
 * - Arrow keys nudge (Shift for 10px steps), Delete removes, Escape deselects.
 *
 * Every committed gesture goes through `store.edit()`, so it participates in
 * undo/redo and is persisted by `store.save()`.
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

  // Selection is derived: a stale nodeId (deleted node, slide change) simply
  // resolves to null and the overlay disappears without effect-driven resets.
  const selectedNode =
    state.mode !== "idle" && slide
      ? (slide.nodes.find((n) => n.id === state.nodeId) ?? null)
      : null;

  const isTextMode = state.mode === "text";
  const publicState: SelectionState = { mode: state.mode, selectedNode };

  if (!presentation?.pkg || !slide || !slideId) return null;

  /** The wrapper div that contains both the rendered slide and this overlay. */
  function slideWrapper(): HTMLElement | null {
    return rootRef.current?.parentElement?.parentElement ?? null;
  }

  function shapeElement(nodeId: string): HTMLElement | null {
    return (
      slideWrapper()?.querySelector<HTMLElement>(`[data-pptx-node-id="${CSS.escape(nodeId)}"]`) ??
      null
    );
  }

  /** Find the text container div inside a shape element. */
  function textContainerOf(shapeEl: HTMLElement): HTMLElement | null {
    // The text container is the flex-column div that holds the paragraph divs.
    // It's the first descendant with data-pptx-p children.
    const firstPara = shapeEl.querySelector("[data-pptx-p]");
    return (firstPara?.parentElement as HTMLElement) ?? null;
  }

  /** Topmost slide node under the pointer, using DOM paint order. */
  function hitTest(clientX: number, clientY: number): string | null {
    const root = rootRef.current;
    const wrapper = slideWrapper();
    if (!root || !wrapper) return null;
    for (const el of document.elementsFromPoint(clientX, clientY)) {
      if (root.contains(el)) continue; // skip the overlay itself
      if (!wrapper.contains(el)) continue;
      const nodeEl = (el as HTMLElement).closest<HTMLElement>("[data-pptx-node-id]");
      if (nodeEl) return nodeEl.getAttribute("data-pptx-node-id");
    }
    return null;
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

  function nodeHasText(nodeId: string): boolean {
    const node = slide!.nodes.find((n) => n.id === nodeId);
    return node?.nodeType === "shape" && Boolean((node as ShapeNodeData).textBody);
  }

  function enterTextMode(nodeId: string, clientX?: number, clientY?: number): void {
    const shapeEl = shapeElement(nodeId);
    if (!shapeEl) return;
    const textEl = textContainerOf(shapeEl);
    if (!textEl) return;

    textEl.contentEditable = "plaintext-only";
    // Fallback: some browsers (Firefox) don't support plaintext-only.
    if (!textEl.isContentEditable) {
      textEl.contentEditable = "true";
    }
    textEl.style.cursor = "text";
    textEl.style.outline = "none";
    textEl.focus();

    // Place caret at click position if possible.
    if (clientX !== undefined && clientY !== undefined) {
      try {
        if (document.caretRangeFromPoint) {
          const range = document.caretRangeFromPoint(clientX, clientY);
          if (range) {
            const sel = window.getSelection();
            sel?.removeAllRanges();
            sel?.addRange(range);
          }
        }
      } catch {
        // Ignore caret placement failures.
      }
    }

    setState({ mode: "text", nodeId, editingEl: textEl });
  }

  function exitTextMode(): void {
    if (state.mode !== "text") return;
    const { nodeId, editingEl } = state;

    editingEl.contentEditable = "false";
    editingEl.style.cursor = "";

    // Read back the DOM and commit if changed.
    const node = slide!.nodes.find((n) => n.id === nodeId);
    if (node) {
      const readBack = readBackTextBody(editingEl);
      if (textBodyChanged(node, readBack)) {
        commitEdit(
          () =>
            store.edit({
              type: "setTextBody",
              slideId: slideId!,
              nodeId,
              paragraphs: readBack,
            }),
          undefined,
          () => onTextChange?.(nodeId),
          (error) => onTextChange?.(nodeId, error),
        );
      }
    }

    setState({ mode: "selected", nodeId });
    rootRef.current?.focus({ preventScroll: true });
  }

  // --- Pointer events ---

  function onPointerDown(event: React.PointerEvent<HTMLDivElement>): void {
    if (event.button !== 0) return;

    // In text mode, clicks outside the editing shape exit text mode.
    if (isTextMode) {
      const nodeId = hitTest(event.clientX, event.clientY);
      if (nodeId !== state.nodeId) {
        exitTextMode();
        if (!nodeId) {
          setState({ mode: "idle" });
        }
      }
      // Let the click propagate to the contentEditable element.
      return;
    }

    const nodeId = hitTest(event.clientX, event.clientY);
    if (!nodeId) {
      setState({ mode: "idle" });
      return;
    }

    // Double-click on a text shape enters text mode immediately.
    if (event.detail >= 2 && nodeHasText(nodeId)) {
      enterTextMode(nodeId, event.clientX, event.clientY);
      return;
    }

    // Clicking an already-selected text shape enters text mode (like
    // PowerPoint: first click selects, second click starts editing).
    if (state.mode === "selected" && state.nodeId === nodeId && nodeHasText(nodeId)) {
      enterTextMode(nodeId, event.clientX, event.clientY);
      return;
    }

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
      const { nodeId, dx, dy, moved } = state;
      setState({ mode: "selected", nodeId });
      if (moved && selectedNode) {
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
          () => onNodeTransform?.(nodeId),
          (error) => onNodeTransform?.(nodeId, error),
        );
      } else {
        const el = shapeElement(nodeId);
        if (el) el.style.translate = "";
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
    // In text mode, only Escape exits; everything else goes to contentEditable.
    if (isTextMode) {
      if (event.key === "Escape") {
        event.preventDefault();
        exitTextMode();
      }
      return;
    }

    const mod = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();

    // Undo / redo
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

interface SelectionBoxProps {
  node: SlideNode;
  state: InternalState;
  zoom: number;
  onHandlePointerDown: (event: React.PointerEvent<HTMLDivElement>, handle: HandleDirection) => void;
}

function SelectionBox({ node, state, zoom, onHandlePointerDown }: SelectionBoxProps) {
  let rect = nodeRect(node);
  if (state.mode === "move" && state.moved) {
    rect = { ...rect, x: rect.x + state.dx, y: rect.y + state.dy };
  } else if (state.mode === "resize") {
    rect = resizeRect(rect, state.handle, state.dx, state.dy);
  }

  const isTextMode = state.mode === "text";
  // Resize math assumes axis-aligned bounds; rotated shapes get move-only.
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
    </div>
  );
}

export namespace Selection {
  export type State = SelectionState;
  export type Props = SelectionProps;
}
