import * as React from "react";

import type { SlideNode } from "@diceui/pptx-core";

import type { HandleDirection, Rect } from "./geometry";
import type { InternalState, SelectionState } from "./selection";

// Internal, so the overlay's interaction state stays free to change.
export interface SelectionPartsContextValue {
  state: InternalState;
  zoom: number;
  selectedNodes: SlideNode[];
  resizeRects: Map<string, Rect> | null;
  // False mid-band, when the pointer is captured; before the band starts the press is still a click.
  showResizeHandles: boolean;
  frameRef: React.RefObject<HTMLDivElement | null>;
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
  rotation: number;
  // Screen px.
  width: number;
  height: number;
  showHandles: boolean;
  showRotateHandle: boolean;
}

export const SelectionBoxContext = React.createContext<SelectionBoxContextValue | null>(null);

export const SELECTION_BOX_NAME = "Presentation.SelectionBox";

// A press is a move only past the threshold, so a click doesn't blink toolbars that hide while moving.
export function getReportedMode(state: InternalState): SelectionState["mode"] {
  return state.mode === "move" && !state.moved ? "selected" : state.mode;
}
