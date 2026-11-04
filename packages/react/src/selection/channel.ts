import * as React from "react";

import type { AddTextBoxOperation, ParagraphStyle, SlideNode, TextStyle } from "@diceui/pptx-core";

import { useSlide, useStoreContext } from "../context";
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

export interface SelectionSnapshot {
  /** Slide the selection belongs to. */
  slideId: string | null;
  mode: SelectionState["mode"];
  nodeIds: string[];
  textStyle: ActiveTextStyle | null;
}

/** Commands the mounted `Presentation.Selection` serves. */
export interface SelectionController {
  select: (nodeIds: string[]) => void;
  editText: (nodeId: string) => void;
  setTextStyle: (style: TextStyle) => Promise<void>;
  setParagraphStyle: (style: ParagraphStyle) => Promise<void>;
}

const IDLE_SNAPSHOT: SelectionSnapshot = {
  slideId: null,
  mode: "idle",
  nodeIds: [],
  textStyle: null,
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
    getIsSameTextStyle(a.textStyle, b.textStyle)
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

  /**
   * Formatting of the text being edited (in text mode), or of every selected
   * shape's text (otherwise). `null` when the selection has no text.
   */
  textStyle: ActiveTextStyle | null;

  /** Selects the given nodes on the active slide. Ends text editing. */
  select: (nodeIds: string[]) => void;

  /** Clears the selection. Ends text editing. */
  clear: () => void;

  /**
   * Starts editing a shape's text, caret at the end. Waits for a slide
   * re-render, so it works right after an edit that created the shape.
   */
  editText: (nodeId: string) => void;

  /**
   * Adds a text box to the active slide, centered unless `position` is given,
   * and starts editing it. Resolves with the new node's id.
   */
  insertTextBox: (
    options?: Omit<AddTextBoxOperation, "type" | "slideId">,
  ) => Promise<string | undefined>;

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

  const nodeIds = snapshot.slideId === slideId ? snapshot.nodeIds : IDLE_SNAPSHOT.nodeIds;
  const nodes = React.useMemo(
    () =>
      slide
        ? nodeIds
            .map((id) => slide.nodes.find((node) => node.id === id))
            .filter((node): node is SlideNode => node !== undefined)
        : [],
    [slide, nodeIds],
  );

  return {
    mode: snapshot.slideId === slideId ? snapshot.mode : "idle",
    nodeIds,
    nodes,
    textStyle: snapshot.slideId === slideId ? snapshot.textStyle : null,
    select: React.useCallback((ids: string[]) => channel.getController()?.select(ids), [channel]),
    clear: React.useCallback(() => channel.getController()?.select([]), [channel]),
    editText: React.useCallback(
      (nodeId: string) => channel.getController()?.editText(nodeId),
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
