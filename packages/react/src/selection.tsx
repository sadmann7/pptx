import * as React from "react";

import type { Position, SetTextBodyParagraph, ShapeNodeData, SlideNode } from "@diceui/pptx-core";

import { usePresentation, useSlide, useSlideRevision, useStoreContext, useZoom } from "./context";
import type { RenderProp } from "./render";
import { mergeRefs, renderElement } from "./render";

const SELECTION_NAME = "Presentation.Selection";

/** Temporary debug logging for inline text editing. */
const DEBUG_TEXT_EDIT = false;

function debugLog(...args: unknown[]): void {
  if (!DEBUG_TEXT_EDIT) return;
  console.debug("[pptx-selection]", ...args);
}

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

// Internal discriminated union: not part of the public API.
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
    }
  | {
      /** Rubber-band selection from a drag on empty canvas (client coords). */
      mode: "marquee";
      startX: number;
      startY: number;
      curX: number;
      curY: number;
    }
  | {
      mode: "resize";
      nodeId: string;
      handle: HandleDirection;
      startX: number;
      startY: number;
      dx: number;
      dy: number;
      /** Shift held during the drag: corner handles keep the aspect ratio. */
      lockAspect: boolean;
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

const CORNER_HANDLES: ReadonlySet<HandleDirection> = new Set(["nw", "ne", "se", "sw"]);

/**
 * Apply a resize drag (slide-px deltas) to the original rect, clamped to
 * MIN_SIZE. With `lockAspect` (Shift held), corner handles scale both
 * dimensions proportionally around the opposite corner, like PowerPoint.
 */
function resizeRect(
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
    // Follow the axis the pointer changed most (relative to the shape) and
    // derive the other from the original aspect ratio.
    let scale =
      Math.abs(w / origin.w - 1) >= Math.abs(h / origin.h - 1) ? w / origin.w : h / origin.h;
    // Keep both dimensions at or above the minimum size.
    scale = Math.max(scale, MIN_SIZE / Math.min(origin.w, origin.h));
    w = origin.w * scale;
    h = origin.h * scale;
    // Re-anchor so the opposite corner stays put.
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

  const children = Array.from(container.children).filter(
    (el) => el instanceof HTMLElement,
  ) as HTMLElement[];

  // After destructive edits the paragraph divs may be gone, leaving run
  // spans (or bare text) directly in the container; treat the container
  // itself as a single implicit paragraph in that case.
  const paraDivs = children.filter((el) => el.dataset.pptxR === undefined);
  const effectiveDivs =
    paraDivs.length > 0 && paraDivs.length === children.length ? paraDivs : [container];
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

/**
 * Strip zero-width spaces: line-height spacer spans from the renderer must
 * not reach the model.
 */
function cleanText(raw: string | null | undefined): string {
  return (raw ?? "").replace(/\u200B/g, "");
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
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun });
        lastSourceR = sourceRun;
      }
    } else if (child instanceof HTMLBRElement) {
      // Browsers insert <br> for empty paragraphs; skip.
    } else if (child.nodeType === Node.TEXT_NODE) {
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceR });
      }
    } else if (child instanceof HTMLElement) {
      const text = cleanText(child.textContent);
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
    // Compare concatenated text, not run-by-run: a paragraph with no runs
    // reads back as a single empty run, which is not a real change (plain
    // shapes render an empty editable paragraph before any typing happens).
    const origText = paragraphs[i].runs.map((r) => r.text ?? "").join("");
    const newText = readBack[i].runs.map((r) => r.text).join("");
    if (origText !== newText) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

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
 * - Drag on empty canvas → marquee-select fully enclosed shapes.
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
  const stateRef = React.useRef(state);
  stateRef.current = state;

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
    state.mode === "selected" || state.mode === "move"
      ? state.nodeIds
      : state.mode === "resize" || state.mode === "text"
        ? [state.nodeId]
        : [];
  const selectedNodes: SlideNode[] = slide
    ? selectedIds
        .map((id) => slide.nodes.find((n) => n.id === id))
        .filter((n): n is SlideNode => n !== undefined)
    : [];
  /** The single selected node; with a multi-selection, the first one. */
  const selectedNode = selectedNodes[0] ?? null;
  const isSoloSelection = selectedNodes.length === 1;

  const isTextMode = state.mode === "text";
  const publicState: SelectionState = { mode: state.mode, selectedNode, selectedNodes };

  if (!presentation?.sourcePackage || !slide || !slideId) return null;

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
    // Empty placeholders also render a prompt overlay ("Click to add text")
    // whose paragraphs carry data-pptx-p; skip it: only the real text
    // container is editable.
    for (const para of Array.from(shapeEl.querySelectorAll<HTMLElement>("[data-pptx-p]"))) {
      if (para.closest("[data-pptx-placeholder-prompt]")) continue;
      return para.parentElement;
    }
    return null;
  }

  function placeholderPromptOf(shapeEl: HTMLElement): HTMLElement | null {
    return shapeEl.querySelector<HTMLElement>("[data-pptx-placeholder-prompt]");
  }

  /**
   * Drop any document text selection left behind by earlier gestures.
   * A stale selection under the pointer turns the next press-and-drag into a
   * native drag-and-drop, cancelling our pointer stream mid-resize/move.
   */
  function clearDocumentSelection(): void {
    const sel = document.getSelection();
    if (sel && !sel.isCollapsed) sel.removeAllRanges();
  }

  /** Convert viewport (client) coordinates to slide-space px. */
  function clientToSlide(clientX: number, clientY: number): Position {
    const wrapperRect = slideWrapper()?.getBoundingClientRect();
    if (!wrapperRect) return { x: 0, y: 0 };
    return { x: (clientX - wrapperRect.left) / zoom, y: (clientY - wrapperRect.top) / zoom };
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

  /**
   * PowerPoint: a single click starts text editing only on dedicated text
   * boxes and placeholders. Regular shapes with text select on click and
   * need a double click (or typing) to edit.
   */
  function nodeEditsOnClick(nodeId: string): boolean {
    const node = slide!.nodes.find((n) => n.id === nodeId);
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
      onRollback?.();
      if (onFailure) {
        onFailure(err);
      } else {
        console.warn("[pptx] edit failed:", err);
      }
    });
  }

  // --- Text editing helpers ---

  function enterTextMode(
    nodeId: string,
    clientX?: number,
    clientY?: number,
    insertText?: string,
  ): void {
    const shapeEl = shapeElement(nodeId);
    const textEl = shapeEl ? textContainerOf(shapeEl) : null;
    debugLog("enterTextMode", {
      nodeId,
      shapeElFound: Boolean(shapeEl),
      textElFound: Boolean(textEl),
      textElConnected: textEl?.isConnected,
      textElHtml: textEl?.outerHTML.slice(0, 200),
    });
    if (!textEl) {
      // No editable text container (e.g. decorative shape): fall back to
      // selection instead of leaving the previous mode (move) active.
      setState({ mode: "selected", nodeIds: [nodeId] });
      return;
    }

    textEl.contentEditable = "plaintext-only";
    if (!textEl.isContentEditable) {
      textEl.contentEditable = "true";
    }
    textEl.style.cursor = "text";
    textEl.style.outline = "none";

    // Capture a styling template before the browser can mutate the DOM.
    // When the DOM has no run spans left (e.g. re-entering a box that was
    // cleared without a re-render), keep the template captured earlier for
    // this shape instead of discarding it.
    const templateSource = textEl.querySelector<HTMLElement>("[data-pptx-r]");
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
    const prompt = shapeEl ? placeholderPromptOf(shapeEl) : null;
    if (prompt) prompt.style.display = "none";

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

    // PowerPoint: typing while a shape is selected starts editing with that
    // keystroke as the first character.
    if (insertText) {
      insertTextAtCaret(textEl, nodeId, insertText);
    }

    const sel = window.getSelection();
    debugLog("enterTextMode done", {
      activeElement: document.activeElement?.tagName,
      activeIsTextEl: document.activeElement === textEl,
      selAnchorInTextEl: sel?.anchorNode ? textEl.contains(sel.anchorNode) : null,
      selAnchor: sel?.anchorNode?.nodeName,
      isContentEditable: textEl.isContentEditable,
    });

    setState({ mode: "text", nodeId, editingEl: textEl });
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
   * must land inside a `[data-pptx-r]` span to inherit the run's font and
   * color: a bare text node at the container/paragraph level renders with
   * unstyled defaults (e.g. black text on a dark slide → invisible typing).
   * Returns false when the scope has no run spans.
   */
  function snapCaretIntoRun(scope: HTMLElement): boolean {
    const runs = scope.querySelectorAll<HTMLElement>("[data-pptx-r]");
    const last = runs[runs.length - 1];
    if (!last) return false;
    const sel = window.getSelection();
    if (!sel) return false;

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
    sel.removeAllRanges();
    sel.addRange(range);
    return true;
  }

  /** Move the caret to the given text node offset. */
  function setCaret(node: Node, offset: number): void {
    const sel = window.getSelection();
    if (!sel) return;
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
  }

  /**
   * Empty paragraphs are rendered with a placeholder `<br>` to preserve their
   * line height; once real text is inserted the `<br>` would push it onto a
   * second line. Remove it, but only when the paragraph holds nothing except
   * the just-inserted text; `<br>`s in paragraphs with other content are
   * genuine line breaks.
   */
  function removePlaceholderBreak(from: Node, insertedText: string): void {
    const el = from instanceof Element ? (from as HTMLElement) : from.parentElement;
    const para = el?.closest<HTMLElement>("[data-pptx-p]");
    if (!para) return;
    if (cleanText(para.textContent) !== cleanText(insertedText)) return;
    for (const child of Array.from(para.children)) {
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
  function insertTextAtCaret(editingEl: HTMLElement, nodeId: string, data: string): boolean {
    const sel = window.getSelection();
    const anchor = sel?.anchorNode;
    if (!anchor || !sel?.isCollapsed || !editingEl.contains(anchor)) return false;

    const anchorEl = anchor instanceof Element ? (anchor as HTMLElement) : anchor.parentElement;
    const runSpan = anchorEl?.closest<HTMLElement>("[data-pptx-r]");

    if (runSpan) {
      if (anchor.nodeType === Node.TEXT_NODE && (anchor.textContent ?? "").length > 0) {
        // Insert into the existing text node at the caret offset.
        const textNode = anchor as Text;
        const offset = Math.min(sel.anchorOffset, textNode.length);
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
    if (anchorEl?.closest("[data-pptx-p]")) {
      sel.getRangeAt(0).insertNode(span);
    } else {
      // Caret sits at the container level (e.g. caret fallback when all run
      // spans were destroyed): inserting there would land the span *below*
      // the paragraph div, on its own line. Append into the last paragraph
      // instead.
      const paras = editingEl.querySelectorAll<HTMLElement>("[data-pptx-p]");
      const lastPara = paras[paras.length - 1];
      if (lastPara) lastPara.appendChild(span);
      else sel.getRangeAt(0).insertNode(span);
    }
    removePlaceholderBreak(span, data);
    setCaret(textNode, textNode.length);
    debugLog("rewrapped typing in run span", { spanHtml: span.outerHTML.slice(0, 200) });
    return true;
  }

  function interceptTextInsertion(e: InputEvent, editingEl: HTMLElement, nodeId: string): void {
    if (e.inputType !== "insertText" || !e.data) return;
    const sel = window.getSelection();
    const anchor = sel?.anchorNode;
    if (!anchor || !sel.isCollapsed || !editingEl.contains(anchor)) return;

    const anchorEl = anchor instanceof Element ? (anchor as HTMLElement) : anchor.parentElement;

    // Caret in a non-empty text node inside a styled element (run span,
    // hyperlink run, or a fallback span from a previous insertion): the
    // default insertion behaves correctly, don't interfere.
    if (
      anchor.nodeType === Node.TEXT_NODE &&
      (anchor.textContent ?? "").length > 0 &&
      anchorEl &&
      (anchorEl.closest("[data-pptx-r]") || anchorEl.tagName === "SPAN")
    ) {
      return;
    }

    if (insertTextAtCaret(editingEl, nodeId, e.data)) {
      e.preventDefault();
    }
  }

  /**
   * Post-input safety net: if typed text still ended up as a bare text node
   * outside any run span (paths not covered by beforeinput, e.g. IME
   * composition or paste), reparent it into a clone of the template span.
   */
  function repairRunStyling(editingEl: HTMLElement, nodeId: string): void {
    const sel = window.getSelection();
    const anchor = sel?.anchorNode;
    if (!anchor || !sel?.isCollapsed || !editingEl.contains(anchor)) return;
    if (anchor.nodeType !== Node.TEXT_NODE) return;
    const anchorEl = anchor.parentElement;
    if (!anchorEl || anchorEl.closest("[data-pptx-r]")) return;

    const template = runTemplateFor(nodeId);
    if (!template) return;

    const span = template.cloneNode(false) as HTMLElement;
    const offset = sel.anchorOffset;
    // Reparent the bare text node into a styled span; the node identity is
    // preserved so the caret offset stays valid.
    anchor.parentNode?.insertBefore(span, anchor);
    span.appendChild(anchor);
    removePlaceholderBreak(span, anchor.textContent ?? "");
    setCaret(anchor, Math.min(offset, (anchor.textContent ?? "").length));
    debugLog("repaired run styling at caret", { spanHtml: span.outerHTML.slice(0, 200) });
  }

  function placeCaretAtEnd(textEl: HTMLElement): void {
    try {
      if (snapCaretIntoRun(textEl)) return;
      const sel = window.getSelection();
      if (!sel) return;
      // No run spans: prefer the end of the last paragraph div over the
      // container itself so typing lands on the paragraph's line rather
      // than a new line below it.
      const paras = textEl.querySelectorAll<HTMLElement>("[data-pptx-p]");
      const target = paras[paras.length - 1] ?? textEl;
      const range = document.createRange();
      range.selectNodeContents(target);
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
          fixupCaretAnchor(sel);
          return;
        }
      }
      // Legacy WebKit API.
      if (document.caretRangeFromPoint) {
        const range = document.caretRangeFromPoint(clientX, clientY);
        if (range) {
          sel.removeAllRanges();
          sel.addRange(range);
          fixupCaretAnchor(sel);
        }
      }
    } catch {
      // Ignore caret placement failures.
    }
  }

  /**
   * If a point-placed caret anchored on an element (not a text node inside a
   * run span), snap it into the nearest run so typing inherits run styling.
   */
  function fixupCaretAnchor(sel: Selection): void {
    const anchor = sel.anchorNode;
    if (!anchor) return;
    if (anchor.nodeType === Node.TEXT_NODE) return;
    const el = anchor as HTMLElement;
    // Prefer a run in the paragraph under the caret, else anywhere in scope.
    const paraDiv = el.closest?.("[data-pptx-p]") as HTMLElement | null;
    snapCaretIntoRun(paraDiv ?? el);
  }

  /** Tear down contentEditable and commit the edited text if it changed. */
  function commitTextEdits(current: Extract<InternalState, { mode: "text" }>): void {
    const { nodeId, editingEl } = current;

    editingEl.contentEditable = "false";
    editingEl.style.cursor = "";

    // Un-hide the placeholder prompt hidden by enterTextMode. When the edit
    // commits, the re-render rebuilds the shape anyway; when nothing changed
    // (no re-render), the prompt must come back by hand.
    const shapeEl = shapeElement(nodeId);
    const prompt = shapeEl ? placeholderPromptOf(shapeEl) : null;
    if (prompt) prompt.style.display = "";

    const node = slide!.nodes.find((n) => n.id === nodeId);
    if (!node) return;
    const readBack = readBackTextBody(editingEl);
    debugLog("commitTextEdits", {
      nodeId,
      editingElConnected: editingEl.isConnected,
      editingElHtml: editingEl.innerHTML.slice(0, 200),
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
        // Clicked a different shape: exit text and start interacting with it
        // in the same gesture (PowerPoint: you can immediately drag another
        // shape while editing). preventDefault stops the browser's default
        // mousedown focus change from blurring the overlay; without it,
        // typing right after this click would go nowhere.
        e.preventDefault();
        commitTextEdits(cur);
        rootRef.current?.setPointerCapture(e.pointerId);
        rootRef.current?.focus({ preventScroll: true });
        setState({
          mode: "move",
          nodeIds: [hitNodeId],
          primaryId: hitNodeId,
          startX: e.clientX,
          startY: e.clientY,
          dx: 0,
          dy: 0,
          moved: false,
        });
      } else if (!hitNodeId) {
        // Clicked empty area: exit text, deselect. preventDefault keeps the
        // overlay focused so undo/redo shortcuts still work afterwards.
        e.preventDefault();
        doExitTextMode(cur, null);
      }
      // Else clicked inside the same shape but outside the text container
      // (e.g. shape padding area); stay in text mode.
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

    function onDocBeforeInput(e: Event): void {
      const cur = stateRef.current;
      if (cur.mode !== "text") return;
      interceptTextInsertion(e as InputEvent, cur.editingEl, cur.nodeId);
    }

    function onDocInput(e: Event): void {
      const cur = stateRef.current;
      if (cur.mode !== "text") return;
      repairRunStyling(cur.editingEl, cur.nodeId);
      const target = e.target as HTMLElement;
      const sel = window.getSelection();
      const anchorNode = sel?.anchorNode;
      const anchorEl =
        anchorNode instanceof Element
          ? (anchorNode as HTMLElement)
          : (anchorNode?.parentElement ?? null);
      const rect = anchorEl?.getBoundingClientRect();
      const cs = anchorEl ? getComputedStyle(anchorEl) : null;
      const onTop =
        rect && rect.width + rect.height > 0
          ? document.elementFromPoint(rect.x + rect.width / 2, rect.y + rect.height / 2)
          : null;
      debugLog("input event", {
        typedContent: cur.editingEl.textContent?.slice(0, 80),
        anchorTag: anchorEl?.tagName,
        anchorIsRunSpan: anchorEl?.dataset?.pptxR !== undefined,
        anchorRect: rect
          ? `${Math.round(rect.x)},${Math.round(rect.y)} ${Math.round(rect.width)}x${Math.round(rect.height)}`
          : null,
        color: cs?.color,
        fontSize: cs?.fontSize,
        opacity: cs?.opacity,
        visibility: cs?.visibility,
        onTopAtAnchor: onTop
          ? `${onTop.tagName} ${(onTop as HTMLElement).dataset?.pptxNodeId ?? ""}`
          : null,
        anchorHtml: anchorEl?.outerHTML.slice(0, 200),
        targetIsEditingEl: target === cur.editingEl,
        editingElConnected: cur.editingEl.isConnected,
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

  // --- Text-mode repair after slide re-renders ---
  // A revision bump makes SlideImpl replace the slide DOM in its effect,
  // detaching our contentEditable element: typed text would go into the
  // detached tree and never appear on screen. SlideImpl is a parent, so its
  // effect runs after this one; the rAF fires after the whole effects flush,
  // when the fresh DOM is in place, and re-attaches editing to it.
  React.useEffect(() => {
    debugLog("revision effect", { slideRevision, mode: stateRef.current.mode });
    if (stateRef.current.mode !== "text") return;
    const raf = requestAnimationFrame(() => {
      const cur = stateRef.current;
      if (cur.mode !== "text") return;
      debugLog("repair check", { editingElConnected: cur.editingEl.isConnected });
      if (cur.editingEl.isConnected) return;
      debugLog("repairing: re-entering text mode", { nodeId: cur.nodeId });
      enterTextMode(cur.nodeId);
    });
    return () => cancelAnimationFrame(raf);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- repair keyed on revision only
  }, [slideRevision]);

  // --- Overlay pointer events (non-text modes) ---

  function onPointerDown(event: React.PointerEvent<HTMLDivElement>): void {
    if (event.button !== 0 || isTextMode) return;

    const nodeId = hitTest(event.clientX, event.clientY);
    const selShapeEl = selectedNode ? shapeElement(selectedNode.id) : null;
    debugLog("root pointerdown", {
      hit: nodeId,
      mode: state.mode,
      x: event.clientX,
      y: event.clientY,
      target: (event.target as HTMLElement)?.tagName,
      targetHandle: (event.target as HTMLElement)?.dataset?.resizeHandle,
      selectedRect: selectedNode ? nodeRect(selectedNode) : null,
      selectedDomRect: selShapeEl
        ? (() => {
            const r = selShapeEl.getBoundingClientRect();
            return `${Math.round(r.x)},${Math.round(r.y)} ${Math.round(r.width)}x${Math.round(r.height)}`;
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
      // click resolves to deselect on pointer-up.
      setState({
        mode: "marquee",
        startX: event.clientX,
        startY: event.clientY,
        curX: event.clientX,
        curY: event.clientY,
      });
      return;
    }

    // Shift/Ctrl+click toggles the shape in and out of the selection
    // (PowerPoint multi-select).
    if (event.shiftKey || event.ctrlKey || event.metaKey) {
      const nextIds = selectedIds.includes(nodeId)
        ? selectedIds.filter((id) => id !== nodeId)
        : [...selectedIds, nodeId];
      setState(nextIds.length > 0 ? { mode: "selected", nodeIds: nextIds } : { mode: "idle" });
      return;
    }

    // Pressing a shape that is already part of the selection drags the whole
    // selection; anything else starts a fresh single-shape gesture. On
    // pointer-up we check if the user actually dragged; if not (a click),
    // text shapes enter text mode and non-text shapes get selected.
    const nodeIds = selectedIds.includes(nodeId) ? selectedIds : [nodeId];
    setState({
      mode: "move",
      nodeIds,
      primaryId: nodeId,
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
    debugLog("handle pointerdown", { handle, mode: state.mode, button: event.button });
    // Resize handles only render for a single selection.
    if (event.button !== 0 || !isSoloSelection || isTextMode) return;
    event.stopPropagation();
    // See onPointerDown: stop selection extension / native drag initiation.
    event.preventDefault();
    clearDocumentSelection();
    rootRef.current?.setPointerCapture(event.pointerId);
    setState({
      mode: "resize",
      nodeId: selectedIds[0],
      handle,
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
      if (moved) {
        for (const id of state.nodeIds) {
          const el = shapeElement(id);
          if (el) el.style.translate = `${dx}px ${dy}px`;
        }
      }
      setState({ ...state, dx, dy, moved });
    } else {
      // Track Shift live so holding/releasing it mid-drag toggles the lock.
      setState({ ...state, dx, dy, lockAspect: event.shiftKey });
    }
  }

  function onPointerUp(): void {
    if (state.mode === "marquee") {
      const { startX, startY, curX, curY } = state;
      const dragged = Math.hypot(curX - startX, curY - startY) > DRAG_THRESHOLD;
      if (!dragged) {
        // Plain click on empty canvas → deselect.
        setState({ mode: "idle" });
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
        .filter((n) => {
          const r = nodeRect(n);
          return (
            r.x >= box.x && r.y >= box.y && r.x + r.w <= box.x + box.w && r.y + r.h <= box.y + box.h
          );
        })
        .map((n) => n.id);
      setState(contained.length > 0 ? { mode: "selected", nodeIds: contained } : { mode: "idle" });
    } else if (state.mode === "move") {
      const { nodeIds, primaryId, dx, dy, moved, resumeText } = state;

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
            for (const id of nodeIds) {
              const el = shapeElement(id);
              if (el) el.style.translate = "";
            }
          },
          () => {
            for (const id of nodeIds) onNodeTransform?.(id);
            // Border drag from text mode: return to editing so the user can
            // keep typing right where they left off (PowerPoint behavior).
            if (resumeText) resumeTextEditing(primaryId);
          },
          (error) => {
            for (const id of nodeIds) onNodeTransform?.(id, error);
          },
        );
      } else {
        // Click (no drag). Clear any stray preview translate.
        for (const id of nodeIds) {
          const el = shapeElement(id);
          if (el) el.style.translate = "";
        }

        // Text boxes / placeholders → edit immediately. Everything else
        // (regular shapes, even with text) → select; they edit on
        // double-click or by typing. Clicking a member of a multi-selection
        // without dragging collapses the selection to it (PowerPoint).
        if (resumeText) {
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
      const { nodeId, handle, dx, dy, lockAspect } = state;
      setState({ mode: "selected", nodeIds: [nodeId] });
      debugLog("resize pointerup", {
        nodeId,
        handle,
        dx,
        dy,
        lockAspect,
        hasSelectedNode: Boolean(selectedNode),
        baseRect: selectedNode ? nodeRect(selectedNode) : null,
      });
      if (selectedNode && (dx !== 0 || dy !== 0)) {
        const next = resizeRect(nodeRect(selectedNode), handle, dx, dy, lockAspect);
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

  /** Double click on a regular shape with text → edit with caret at point. */
  function onDoubleClick(event: React.MouseEvent<HTMLDivElement>): void {
    debugLog("double click", { isTextMode });
    if (isTextMode) return;
    const nodeId = hitTest(event.clientX, event.clientY);
    if (nodeId && nodeHasText(nodeId)) {
      enterTextMode(nodeId, event.clientX, event.clientY);
    }
  }

  /** Native drag-and-drop of slide content (images, stale text selections)
   * cancels our pointer gestures; block it while the overlay is interactive. */
  function onDragStart(event: React.DragEvent<HTMLDivElement>): void {
    debugLog("dragstart blocked", { target: (event.target as HTMLElement)?.tagName });
    event.preventDefault();
  }

  function onPointerCancel(): void {
    debugLog("pointercancel", { mode: state.mode });
    if (state.mode === "move") {
      for (const id of state.nodeIds) {
        const el = shapeElement(id);
        if (el) el.style.translate = "";
      }
      setState({ mode: "selected", nodeIds: state.nodeIds });
    } else if (state.mode === "resize") {
      setState({ mode: "selected", nodeIds: [state.nodeId] });
    } else if (state.mode === "marquee") {
      setState({ mode: "idle" });
    }
  }

  /** Move every given node by `delta`, as a single undoable edit. */
  function nudge(nodes: SlideNode[], delta: Position): void {
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
    const arrows: Record<string, Position> = {
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
          onDoubleClick,
          onKeyDown,
          onDragStart,
          children: (
            <>
              {selectedNodes.map((node) => (
                <SelectionBox
                  key={node.id}
                  node={node}
                  state={state}
                  zoom={zoom}
                  showResizeHandles={isSoloSelection}
                  onHandlePointerDown={onHandlePointerDown}
                  onBorderPointerDown={onBorderPointerDown}
                />
              ))}
              {state.mode === "marquee" && <MarqueeBox state={state} rootEl={rootRef.current} />}
            </>
          ),
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

/** True when a key event originates in an element with its own undo stack;
 * our text-mode contentEditable or any host-app input. Their Ctrl+Z must
 * reach the browser's native undo, not the presentation history. */
function isNativeUndoTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  return (
    target.isContentEditable ||
    target.tagName === "INPUT" ||
    target.tagName === "TEXTAREA" ||
    target.tagName === "SELECT"
  );
}

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

  // Refs keep the effect bound once while callbacks stay fresh.
  const onUndoRef = React.useRef(onUndo);
  onUndoRef.current = onUndo;
  const onRedoRef = React.useRef(onRedo);
  onRedoRef.current = onRedo;

  const editable = presentation?.sourcePackage != null;

  React.useEffect(() => {
    if (!editable) return;

    function onDocKeyDown(e: KeyboardEvent): void {
      if (!(e.ctrlKey || e.metaKey)) return;
      if (isNativeUndoTarget(e.target)) return;
      const key = e.key.toLowerCase();

      if (key === "z" && !e.shiftKey) {
        e.preventDefault();
        const success = store.undo();
        onUndoRef.current?.(success ? "success" : "empty");
      } else if (key === "y" || (key === "z" && e.shiftKey)) {
        e.preventDefault();
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

// ---------------------------------------------------------------------------
// SelectionBox (internal)
// ---------------------------------------------------------------------------

/** Hitbox width (screen px) of the border move strips shown while editing text. */
const BORDER_GRAB_SIZE = 10;

interface SelectionBoxProps {
  node: SlideNode;
  state: InternalState;
  zoom: number;
  /** False in a multi-selection: boxes render without resize handles. */
  showResizeHandles: boolean;
  onHandlePointerDown: (event: React.PointerEvent<HTMLDivElement>, handle: HandleDirection) => void;
  onBorderPointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
}

function SelectionBox({
  node,
  state,
  zoom,
  showResizeHandles,
  onHandlePointerDown,
  onBorderPointerDown,
}: SelectionBoxProps) {
  let rect = nodeRect(node);
  if (state.mode === "move" && state.moved) {
    rect = { ...rect, x: rect.x + state.dx, y: rect.y + state.dy };
  } else if (state.mode === "resize") {
    rect = resizeRect(rect, state.handle, state.dx, state.dy, state.lockAspect);
  }

  const isTextMode = state.mode === "text";
  const showHandles =
    showResizeHandles && node.rotation === 0 && state.mode !== "move" && !isTextMode;

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
        boxShadow: `0 0 0 ${isTextMode ? "2" : "1.5"}px var(--presentation-selection, #2563eb)`,
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
              border: "1.5px solid var(--presentation-selection, #2563eb)",
              borderRadius: 2,
              cursor: HANDLE_CURSORS[direction],
              pointerEvents: "auto",
              touchAction: "none",
            }}
          />
        ))}
      {isTextMode && <BorderMoveStrips onPointerDown={onBorderPointerDown} />}
    </div>
  );
}

/** Rubber-band rectangle drawn while drag-selecting on empty canvas. */
function MarqueeBox({
  state,
  rootEl,
}: {
  state: Extract<InternalState, { mode: "marquee" }>;
  rootEl: HTMLElement | null;
}) {
  // State coords are viewport-relative; the overlay is our positioning context.
  const origin = rootEl?.getBoundingClientRect();
  if (!origin) return null;
  const left = Math.min(state.startX, state.curX) - origin.left;
  const top = Math.min(state.startY, state.curY) - origin.top;
  const width = Math.abs(state.curX - state.startX);
  const height = Math.abs(state.curY - state.startY);
  if (width + height < DRAG_THRESHOLD) return null;

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
