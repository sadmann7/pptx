import * as React from "react";

import type { Position, SlideNode } from "@diceui/pptx-parser";

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
// Public types
// ---------------------------------------------------------------------------

/** The current interaction state of the edit layer. */
export interface SelectionState {
  /** Interaction mode of the selection. */
  mode: "idle" | "selected" | "move" | "resize";
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
  /**
   * Called after `Ctrl+Z` fires.
   * `status` is `"empty"` when the undo stack has nothing to revert.
   * `error` is set if the operation threw.
   *
   * ```ts
   * onUndo={(status, error) => {
   *   if (error) toast.error("Undo failed");
   *   else if (status === "empty") toast("Nothing to undo");
   *   else toast("Undone");
   * }}
   * ```
   */
  onUndo?: (status: "success" | "empty", error?: unknown) => void;
  /**
   * Called after `Ctrl+Shift+Z` / `Ctrl+Y` fires.
   * `status` is `"empty"` when the redo stack has nothing to replay.
   * `error` is set if the operation threw.
   *
   * ```ts
   * onRedo={(status, error) => {
   *   if (error) toast.error("Redo failed");
   *   else if (status === "empty") toast("Nothing to redo");
   *   else toast("Redone");
   * }}
   * ```
   */
  onRedo?: (status: "success" | "empty", error?: unknown) => void;
  /**
   * Called after a node delete is attempted (keyboard or pointer).
   * `error` is set if the operation threw.
   *
   * ```ts
   * onNodeDelete={(nodeId, error) => {
   *   if (error) toast.error("Delete failed");
   *   else toast(`Deleted ${nodeId}`);
   * }}
   * ```
   */
  onNodeDelete?: (nodeId: string, error?: unknown) => void;
  /**
   * Called after a node move or resize is attempted.
   * `error` is set if the operation threw.
   *
   * ```ts
   * onNodeTransform={(nodeId, error) => {
   *   if (error) toast.error("Transform failed");
   * }}
   * ```
   */
  onNodeTransform?: (nodeId: string, error?: unknown) => void;
}

/**
 * PowerPoint-style shape selection and manipulation overlay.
 *
 * Place inside `<Presentation.Slide>`. Renders nothing unless the loaded
 * presentation was opened with `readOnly={false}`.
 *
 * - Click a shape to select it (topmost shape wins, like PowerPoint).
 * - Drag to move; drag the handles to resize (non-rotated shapes).
 * - Arrow keys nudge (Shift for 10px steps), Delete removes, Escape deselects.
 *
 * Every committed gesture goes through `store.edit()`, so it participates in
 * undo/redo and is persisted by `store.save()`.
 */
export const Selection = React.forwardRef<HTMLDivElement, SelectionProps>(function Selection(
  { render, onUndo, onRedo, onNodeDelete, onNodeTransform, ...selectionProps },
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

  function onPointerDown(event: React.PointerEvent<HTMLDivElement>): void {
    if (event.button !== 0) return;
    const nodeId = hitTest(event.clientX, event.clientY);
    if (!nodeId) {
      setState({ mode: "idle" });
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
    if (event.button !== 0 || state.mode === "idle") return;
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
      // Live-preview by translating the rendered shape; the real position is
      // committed (and re-rendered) on pointer-up.
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
        // The preview translate stays on the old DOM; the commit re-renders
        // the slide, replacing it. Cleared only if the edit fails.
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
        // Unmoved click: clear any stray preview translate.
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
    const mod = event.ctrlKey || event.metaKey;

    // Undo / redo — active whenever the overlay is focused, no shape needed.
    // After each operation the slide re-renders (revision bump → SlideImpl
    // replaces its DOM), which steals focus from the overlay. Refocus
    // immediately so subsequent shortcuts keep working.
    const key = event.key.toLowerCase();
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
            pointerEvents: "auto",
            outline: "none",
            touchAction: "none",
          },
        },
        selectionProps,
      ],
    },
  );
});

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

  // Resize math assumes axis-aligned bounds; rotated shapes get move-only.
  const showHandles = node.rotation === 0 && state.mode !== "move";

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
        boxShadow: "0 0 0 1.5px var(--pptx-selection, #2563eb)",
        cursor: "move",
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
