import * as React from "react";

import type {
  AddShapeOperation,
  AddTableOperation,
  AddTextBoxOperation,
  NodeOrder,
  ParagraphStyle,
  SlideNode,
  TableCellAddress,
  TableCellRange,
  TextStyle,
} from "@diceui/pptx-core";

import { useSlide, useSlideRevision, useStoreContext } from "../context";
import type { Store } from "../store";
import type { SelectionState } from "./selection";
import type { ActiveTextStyle } from "./text-range";

/**
 * Attribute that marks UI outside the slide as part of the editor: pressing
 * inside an element that has it (or inside one of its descendants) keeps the
 * selection and text editing alive instead of counting as a press outside.
 *
 * Put it on toolbars, and on the content of their popovers and menus, which
 * usually render in a portal outside the toolbar.
 */
export const KEEP_SELECTION_ATTRIBUTE = "data-pptx-keep-selection";

/** Whether an event target sits inside UI marked with `KEEP_SELECTION_ATTRIBUTE`. */
export function getIsKeepSelectionTarget(target: EventTarget | null): boolean {
  const element =
    target instanceof Element ? target : target instanceof Node ? target.parentElement : null;
  return element?.closest(`[${KEEP_SELECTION_ATTRIBUTE}]`) != null;
}

/** A row or column change at the table cell being edited. */
export type TableChange =
  | { type: "insertRow"; side: "above" | "below" }
  | { type: "insertColumn"; side: "left" | "right" }
  | { type: "deleteRow" }
  | { type: "deleteColumn" }
  | { type: "mergeCells" }
  | { type: "splitCell" };

/** What a click or drag on the slide places, once `startDrawing` arms it. */
export type DrawTool = { type: "textBox" } | { type: "shape"; geometry: string };

export interface SelectionSnapshot {
  /** Slide the selection belongs to. */
  slideId: string | null;
  mode: SelectionState["mode"];
  nodeIds: string[];
  cell: TableCellAddress | null;
  cellRange: TableCellRange | null;
  textStyle: ActiveTextStyle | null;
  drawTool: DrawTool | null;
}

/** Commands the mounted `Presentation.Selection` serves. */
export interface SelectionController {
  /** `focus` moves focus to the overlay, so typing starts the shapes' text. */
  select: (nodeIds: string[], focus?: boolean) => void;
  editText: (nodeId: string, cell?: TableCellAddress) => void;
  startDrawing: (tool: DrawTool | null) => void;
  duplicate: () => Promise<void>;
  remove: () => Promise<void>;
  reorder: (order: NodeOrder) => Promise<void>;
  setTextStyle: (style: TextStyle) => Promise<void>;
  setParagraphStyle: (style: ParagraphStyle) => Promise<void>;
  changeTable: (change: TableChange) => Promise<void>;
  prepareTextHistory: (target: EventTarget | null) => boolean;
  getSelectionRect: () => DOMRect | null;
  getOverlayElement: () => HTMLElement | null;
}

/**
 * A virtual element in the shape Floating UI, Base UI, and Radix accept as an
 * anchor: the selection's on-screen bounds.
 */
export interface SelectionAnchor {
  getBoundingClientRect: () => DOMRect;
  /** The overlay, so the anchor follows when an ancestor scrolls or resizes. */
  contextElement?: Element;
}

const IDLE_SNAPSHOT: SelectionSnapshot = {
  slideId: null,
  mode: "idle",
  nodeIds: [],
  cell: null,
  cellRange: null,
  textStyle: null,
  drawTool: null,
};

function getIsSameTextStyle(a: ActiveTextStyle | null, b: ActiveTextStyle | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return (Object.keys(a) as (keyof ActiveTextStyle)[]).every((key) => a[key] === b[key]);
}

function getIsSameSnapshot(a: SelectionSnapshot, b: SelectionSnapshot): boolean {
  return (
    a.slideId === b.slideId &&
    a.mode === b.mode &&
    a.nodeIds.length === b.nodeIds.length &&
    a.nodeIds.every((id, index) => id === b.nodeIds[index]) &&
    a.cell?.row === b.cell?.row &&
    a.cell?.column === b.cell?.column &&
    a.cellRange?.start.row === b.cellRange?.start.row &&
    a.cellRange?.start.column === b.cellRange?.start.column &&
    a.cellRange?.end.row === b.cellRange?.end.row &&
    a.cellRange?.end.column === b.cellRange?.end.column &&
    getIsSameTextStyle(a.textStyle, b.textStyle) &&
    a.drawTool?.type === b.drawTool?.type &&
    (a.drawTool?.type !== "shape" ||
      a.drawTool.geometry === (b.drawTool as { geometry?: string } | null)?.geometry)
  );
}

/**
 * Carries the selection from `Presentation.Selection`, which lives inside the
 * slide, to toolbars anywhere under the same store. One per store.
 */
export class SelectionChannel {
  private snapshot: SelectionSnapshot = IDLE_SNAPSHOT;
  private controller: SelectionController | null = null;
  private listeners = new Set<() => void>();

  getSnapshot = (): SelectionSnapshot => this.snapshot;

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  publish(snapshot: SelectionSnapshot): void {
    if (getIsSameSnapshot(this.snapshot, snapshot)) return;
    this.snapshot = snapshot;
    for (const listener of this.listeners) listener();
  }

  /** Registers the mounted overlay. Returns the unregister function. */
  register(controller: SelectionController): () => void {
    this.controller = controller;
    return () => {
      if (this.controller !== controller) return;
      this.controller = null;
      this.publish(IDLE_SNAPSHOT);
    };
  }

  getController(): SelectionController | null {
    return this.controller;
  }
}

const channels = new WeakMap<Store, SelectionChannel>();

export function getSelectionChannel(store: Store): SelectionChannel {
  let channel = channels.get(store);
  if (!channel) {
    channel = new SelectionChannel();
    channels.set(store, channel);
  }
  return channel;
}

export interface UseSelectionResult {
  /** Interaction mode of the selection overlay. */
  mode: SelectionState["mode"];

  /** Ids of every selected node, in selection order. In text mode, the edited node. */
  nodeIds: string[];

  /** The selected nodes, matching `nodeIds`. */
  nodes: SlideNode[];

  /** The table cell being edited, or `null` outside a table. */
  cell: TableCellAddress | null;

  /**
   * Cells selected by dragging across them or Shift+clicking, as their
   * top-left and bottom-right corners. Formatting applies to all of them.
   */
  cellRange: TableCellRange | null;

  /**
   * Formatting of the text being edited (in text mode), or of every selected
   * shape's text (otherwise). `null` when the selection has no text.
   */
  textStyle: ActiveTextStyle | null;

  /**
   * The selection's on-screen bounds, for anchoring a floating toolbar or
   * popover to it. `null` when nothing is selected. A new object whenever the
   * selection, its mode, or the slide changes, so a positioner re-measures.
   *
   * ```tsx
   * <Popover.Positioner anchor={anchor} side="top" sideOffset={8} />
   * ```
   */
  anchor: SelectionAnchor | null;

  /** Selects the given nodes on the active slide. Ends text editing. */
  select: (nodeIds: string[]) => void;

  /** Clears the selection. Ends text editing. */
  clear: () => void;

  /**
   * Starts editing a shape's text, caret at the end. For a table, edits `cell`,
   * or the first cell. Waits for a slide re-render, so it works right after an
   * edit that created the shape.
   */
  editText: (nodeId: string, cell?: TableCellAddress) => void;

  /** The armed draw tool, or `null`. */
  drawTool: DrawTool | null;

  /**
   * Arms a draw tool, like Google Slides' text box and shape tools: the slide
   * shows a crosshair, and a click places the text box or shape at its default
   * size there, or a drag draws it (Shift keeps it square). A text box then
   * starts editing and a shape is selected. Escape or `null` disarms it.
   */
  startDrawing: (tool: DrawTool | null) => void;

  /** Duplicates the selected shapes, like Ctrl/Cmd+D, and selects the copies. */
  duplicate: () => Promise<void>;

  /** Deletes the selected shapes, like Delete. */
  remove: () => Promise<void>;

  /** Moves the selected shapes in the stacking order, keeping their order among themselves. */
  reorder: (order: NodeOrder) => Promise<void>;

  /**
   * Adds a text box to the active slide, centered unless `position` is given,
   * and starts editing it. Resolves with the new node's id.
   */
  insertTextBox: (
    options?: Omit<AddTextBoxOperation, "type" | "slideId">,
  ) => Promise<string | undefined>;

  /**
   * Adds a preset shape (e.g. "rect", "ellipse") to the active slide,
   * centered unless `position` is given, and selects it: typing then starts
   * its text. Resolves with the new node's id.
   */
  insertShape: (
    geometry: string,
    options?: Omit<AddShapeOperation, "type" | "slideId" | "geometry">,
  ) => Promise<string | undefined>;

  /**
   * Adds a table to the active slide, centered unless `position` is given, and
   * starts editing its first cell. Resolves with the new node's id.
   */
  insertTable: (
    rows: number,
    columns: number,
    options?: Omit<AddTableOperation, "type" | "slideId" | "rows" | "columns">,
  ) => Promise<string | undefined>;

  /**
   * Inserts or deletes a row or column at the table cell being edited, keeping
   * the caret in that cell, merges the selected `cellRange`, or splits a merged
   * cell. Deleting a table's last row or column deletes the table. Does nothing
   * outside a table cell.
   *
   * ```tsx
   * <button onClick={() => void changeTable({ type: "insertRow", side: "below" })}>
   *   Insert row below
   * </button>
   * ```
   */
  changeTable: (change: TableChange) => Promise<void>;

  /**
   * Formats text. In text mode this applies to the selected text, or to the
   * word at the caret; otherwise to all text in the selected shapes. One undo
   * step. Resolves once the edit is applied.
   */
  setTextStyle: (style: TextStyle) => Promise<void>;

  /**
   * Formats paragraphs. In text mode this applies to the paragraphs the
   * selection touches; otherwise to every paragraph in the selected shapes.
   */
  setParagraphStyle: (style: ParagraphStyle) => Promise<void>;
}

function requireController(channel: SelectionChannel): SelectionController {
  const controller = channel.getController();
  if (!controller) {
    throw new Error("useSelection: no `Presentation.Selection` is mounted.");
  }
  return controller;
}

/**
 * The selection of the editing overlay, and formatting commands for it.
 *
 * Works anywhere under `Presentation.Root` or `Presentation.Provider`, so a
 * toolbar outside the slide can follow the selection. Mark that toolbar with
 * `data-pptx-keep-selection` so pressing it doesn't clear the selection.
 *
 * ```tsx
 * function BoldButton() {
 *   const { textStyle, setTextStyle } = useSelection();
 *   return (
 *     <button
 *       aria-pressed={textStyle?.bold ?? false}
 *       disabled={!textStyle}
 *       onClick={() => void setTextStyle({ bold: !textStyle?.bold })}
 *     >
 *       Bold
 *     </button>
 *   );
 * }
 * ```
 */
export function useSelection(): UseSelectionResult {
  const store = useStoreContext("useSelection");
  const channel = getSelectionChannel(store);
  const snapshot = React.useSyncExternalStore(
    channel.subscribe,
    channel.getSnapshot,
    () => IDLE_SNAPSHOT,
  );
  const { slide, slideId } = useSlide();
  // Edits mutate nodes in place; the revision is what tells the hook to re-read them.
  const revision = useSlideRevision(store, slideId);

  const nodeIds = snapshot.slideId === slideId ? snapshot.nodeIds : IDLE_SNAPSHOT.nodeIds;
  const nodes = React.useMemo(
    () =>
      slide
        ? nodeIds
            .map((id) => slide.nodes.find((node) => node.id === id))
            .filter((node): node is SlideNode => node !== undefined)
        : [],
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- revision is the change signal for in-place edits
    [slide, nodeIds, revision],
  );

  const mode = snapshot.slideId === slideId ? snapshot.mode : "idle";
  const anchor = React.useMemo<SelectionAnchor | null>(
    () =>
      nodeIds.length === 0
        ? null
        : {
            getBoundingClientRect: () =>
              channel.getController()?.getSelectionRect() ?? new DOMRect(),
            get contextElement() {
              return channel.getController()?.getOverlayElement() ?? undefined;
            },
          },
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- mode and revision change the bounds without changing the ids
    [channel, nodeIds, mode, revision],
  );

  return {
    mode,
    anchor,
    nodeIds,
    nodes,
    cell: snapshot.slideId === slideId ? snapshot.cell : null,
    cellRange: snapshot.slideId === slideId ? snapshot.cellRange : null,
    textStyle: snapshot.slideId === slideId ? snapshot.textStyle : null,
    drawTool: snapshot.slideId === slideId ? snapshot.drawTool : null,
    select: React.useCallback((ids: string[]) => channel.getController()?.select(ids), [channel]),
    clear: React.useCallback(() => channel.getController()?.select([]), [channel]),
    editText: React.useCallback(
      (nodeId: string, cell?: TableCellAddress) => channel.getController()?.editText(nodeId, cell),
      [channel],
    ),
    startDrawing: React.useCallback(
      (tool: DrawTool | null) => requireController(channel).startDrawing(tool),
      [channel],
    ),
    duplicate: React.useCallback(async () => requireController(channel).duplicate(), [channel]),
    remove: React.useCallback(async () => requireController(channel).remove(), [channel]),
    reorder: React.useCallback(
      async (order: NodeOrder) => requireController(channel).reorder(order),
      [channel],
    ),
    insertTextBox: React.useCallback(
      async (options?: Omit<AddTextBoxOperation, "type" | "slideId">) => {
        const controller = requireController(channel);
        const activeSlideId = store.getState().activeSlideId;
        if (!activeSlideId) return undefined;
        const { createdNodeIds } = await store.edit({
          type: "addTextBox",
          slideId: activeSlideId,
          ...options,
        });
        const nodeId = createdNodeIds?.[0];
        if (nodeId) controller.editText(nodeId);
        return nodeId;
      },
      [channel, store],
    ),
    insertShape: React.useCallback(
      async (
        geometry: string,
        options?: Omit<AddShapeOperation, "type" | "slideId" | "geometry">,
      ) => {
        const controller = requireController(channel);
        const activeSlideId = store.getState().activeSlideId;
        if (!activeSlideId) return undefined;
        const { createdNodeIds } = await store.edit({
          type: "addShape",
          slideId: activeSlideId,
          geometry,
          ...options,
        });
        const nodeId = createdNodeIds?.[0];
        if (nodeId) controller.select([nodeId], true);
        return nodeId;
      },
      [channel, store],
    ),
    insertTable: React.useCallback(
      async (
        rows: number,
        columns: number,
        options?: Omit<AddTableOperation, "type" | "slideId" | "rows" | "columns">,
      ) => {
        const controller = requireController(channel);
        const activeSlideId = store.getState().activeSlideId;
        if (!activeSlideId) return undefined;
        const { createdNodeIds } = await store.edit({
          type: "addTable",
          slideId: activeSlideId,
          rows,
          columns,
          ...options,
        });
        const nodeId = createdNodeIds?.[0];
        if (nodeId) controller.editText(nodeId);
        return nodeId;
      },
      [channel, store],
    ),
    changeTable: React.useCallback(
      async (change: TableChange) => requireController(channel).changeTable(change),
      [channel],
    ),
    setTextStyle: React.useCallback(
      async (style: TextStyle) => requireController(channel).setTextStyle(style),
      [channel],
    ),
    setParagraphStyle: React.useCallback(
      async (style: ParagraphStyle) => requireController(channel).setParagraphStyle(style),
      [channel],
    ),
  };
}
