/**
 * How `Presentation.Selection` hands its state to the overlay parts drawn
 * inside it, and how `Presentation.SelectionBox` hands each shape to the
 * handles inside it.
 */

import * as React from "react";

import type { SlideNode } from "@diceui/pptx-core";

import type { GripDirection, ResizeScale } from "./geometry";
import type { InternalState } from "./selection";

/**
 * What the overlay hands its parts. Internal: each part exposes only its own
 * state, so the overlay's interaction state stays free to change.
 */
export interface SelectionPartsContextValue {
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
  /** Slide size in slide pixels: snap guides run across all of it. */
  slideWidth: number;
  slideHeight: number;
  onGripPointerDown: (
    event: React.PointerEvent<HTMLDivElement>,
    grip: GripDirection,
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
  showHandles: boolean;
  showRotateHandle: boolean;
}

export const SelectionBoxContext = React.createContext<SelectionBoxContextValue | null>(null);

export const SELECTION_BOX_NAME = "Presentation.SelectionBox";
