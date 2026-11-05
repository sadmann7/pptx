import * as React from "react";

import type { SlideNode } from "@diceui/pptx-core";

import type { HandleDirection, Rect } from "./geometry";
import type { InternalState, SelectionState } from "./selection";

/**
 * What the overlay hands its parts. Internal: each part exposes only its own
 * state, so the overlay's interaction state stays free to change.
 */
export interface SelectionPartsContextValue {
  state: InternalState;
  zoom: number;
  selectedNodes: SlideNode[];
  /** Where each shape lands mid-resize. */
  resizeRects: Map<string, Rect> | null;
  /**
   * False mid-band: the handles are unreachable then (the pointer is captured)
   * and would sit under the rubber band. Until the band starts, the press is
   * still a click and must not blink the handles of what is already selected.
   */
  showResizeHandles: boolean;
  /** The slide-origin frame the parts are positioned in. */
  frameRef: React.RefObject<HTMLDivElement | null>;
  /** Slide size in slide pixels: snap guides run across all of it. */
  slideWidth: number;
  slideHeight: number;
  onHandlePointerDown: (
    event: React.PointerEvent<HTMLDivElement>,
    handle: HandleDirection,
    nodeId: string,
  ) => void;
  onBorderPointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
  onRotatePointerDown: (event: React.PointerEvent<HTMLDivElement>, nodeId: string) => void;
}

export const SelectionPartsContext = React.createContext<SelectionPartsContextValue | null>(null);

export function useSelectionParts(partName: string): SelectionPartsContextValue {
  const context = React.useContext(SelectionPartsContext);
  if (!context) {
    throw new Error(`\`${partName}\` must be used within \`Presentation.Selection\`.`);
  }
  return context;
}

export interface SelectionBoxContextValue {
  node: SlideNode;
  /** The outline's current angle, including a rotation in progress. */
  rotation: number;
  /** The outline's size on screen, in px. */
  width: number;
  height: number;
  showHandles: boolean;
  showRotateHandle: boolean;
}

export const SelectionBoxContext = React.createContext<SelectionBoxContextValue | null>(null);

export const SELECTION_BOX_NAME = "Presentation.SelectionBox";

/**
 * A press reads as a move only once it drags past the threshold, so a click
 * that selects another shape doesn't blink toolbars that hide while moving.
 */
export function getReportedMode(state: InternalState): SelectionState["mode"] {
  return state.mode === "move" && !state.moved ? "selected" : state.mode;
}
