import * as React from "react";

import type { SlideNode } from "@diceui/pptx-core";

import type { StatefulPrimitiveProps } from "../render";
import { renderElement } from "../render";
import {
  getReportedMode,
  SELECTION_BOX_NAME,
  SelectionBoxContext,
  useSelectionParts,
} from "./context";
import { getNodeRect, normalizeAngle } from "./geometry";
import type { SelectionState } from "./selection";

const FRAME_GRAB_SIZE = 10;

const FRAME_GRAB_INSET = FRAME_GRAB_SIZE / 2;

const FRAME_GRAB_HITBOXES: React.CSSProperties[] = [
  {
    left: -FRAME_GRAB_INSET,
    right: -FRAME_GRAB_INSET,
    top: -FRAME_GRAB_INSET,
    height: FRAME_GRAB_SIZE,
  },
  {
    left: -FRAME_GRAB_INSET,
    right: -FRAME_GRAB_INSET,
    bottom: -FRAME_GRAB_INSET,
    height: FRAME_GRAB_SIZE,
  },
  {
    left: -FRAME_GRAB_INSET,
    top: FRAME_GRAB_INSET,
    bottom: FRAME_GRAB_INSET,
    width: FRAME_GRAB_SIZE,
  },
  {
    right: -FRAME_GRAB_INSET,
    top: FRAME_GRAB_INSET,
    bottom: FRAME_GRAB_INSET,
    width: FRAME_GRAB_SIZE,
  },
];

const TABLE_FRAME_GRAB_HITBOXES: React.CSSProperties[] = [
  {
    left: -FRAME_GRAB_SIZE,
    right: -FRAME_GRAB_SIZE,
    top: -FRAME_GRAB_SIZE,
    height: FRAME_GRAB_SIZE,
  },
  {
    left: -FRAME_GRAB_SIZE,
    right: -FRAME_GRAB_SIZE,
    bottom: -FRAME_GRAB_SIZE,
    height: FRAME_GRAB_SIZE,
  },
  { left: -FRAME_GRAB_SIZE, top: 0, bottom: 0, width: FRAME_GRAB_SIZE },
  { right: -FRAME_GRAB_SIZE, top: 0, bottom: 0, width: FRAME_GRAB_SIZE },
];

/** The current state of one selection outline. */
export interface SelectionBoxState {
  /** The selected shape the outline belongs to. */
  node: SlideNode;

  /** Interaction mode of the overlay. `"text"` while the shape's text is edited. */
  mode: SelectionState["mode"];

  /** Whether the selection is being dragged. */
  dragging: boolean;

  /** Whether the selection is being rotated. */
  rotating: boolean;

  /** The outline's angle in degrees, including a rotation in progress. */
  rotation: number;
}

export interface SelectionBoxProps extends StatefulPrimitiveProps<"div", SelectionBoxState> {}

export const SelectionBox = React.forwardRef<HTMLDivElement, SelectionBoxProps>(
  function SelectionBox(
    { render, className, style, children, ...selectionBoxProps },
    forwardedRef,
  ) {
    const { state, zoom, selectedNodes, resizeRects, showResizeHandles, onBorderPointerDown } =
      useSelectionParts(SELECTION_BOX_NAME);

    const isTextMode = state.mode === "text";
    // Not on pointer-down, or a Shift+click that only toggles the selection blinks the handles.
    const isDragging = state.mode === "move" && state.moved;
    const isRotating = state.mode === "rotate";

    return selectedNodes.map((node) => {
      let rect = getNodeRect(node);
      if (state.mode === "move" && state.moved) {
        rect = { ...rect, x: rect.x + state.dx, y: rect.y + state.dy };
      } else if (state.mode === "resize") {
        rect = resizeRects?.get(node.id) ?? rect;
      }
      const rotation = isRotating ? normalizeAngle(node.rotation + state.delta) : node.rotation;
      const showHandles = showResizeHandles && !isDragging && !isRotating;
      const showRotateHandle =
        showResizeHandles &&
        !isDragging &&
        state.mode !== "resize" &&
        // PowerPoint doesn't rotate tables or charts.
        node.nodeType !== "table" &&
        node.nodeType !== "chart";
      const boxState: SelectionBoxState = {
        node,
        mode: getReportedMode(state),
        dragging: isDragging,
        rotating: isRotating,
        rotation,
      };

      return (
        <SelectionBoxContext.Provider
          key={node.id}
          value={{
            node,
            rotation,
            width: rect.w * zoom,
            height: rect.h * zoom,
            showHandles,
            showRotateHandle,
          }}
        >
          {renderElement(
            "div",
            { render, className, style },
            {
              state: boxState,
              ref: forwardedRef,
              props: [
                {
                  "data-presentation-selection-box": "",
                  "data-mode": boxState.mode,
                  "data-dragging": isDragging ? "" : undefined,
                  "data-rotating": isRotating ? "" : undefined,
                  children: (
                    <>
                      {/* Under the handles, which sit on the frame too. */}
                      {(isTextMode || (node.nodeType === "table" && state.mode === "selected")) && (
                        <FrameGrabHitboxes
                          hitboxes={
                            node.nodeType === "table"
                              ? TABLE_FRAME_GRAB_HITBOXES
                              : FRAME_GRAB_HITBOXES
                          }
                          onPointerDown={onBorderPointerDown}
                        />
                      )}
                      {children}
                    </>
                  ),
                  style: {
                    position: "absolute",
                    left: rect.x * zoom,
                    top: rect.y * zoom,
                    width: rect.w * zoom,
                    height: rect.h * zoom,
                    transform: rotation !== 0 ? `rotate(${rotation}deg)` : undefined,
                    pointerEvents: "none",
                  },
                },
                selectionBoxProps,
              ],
            },
          )}
        </SelectionBoxContext.Provider>
      );
    });
  },
);

interface FrameGrabHitboxesProps {
  hitboxes: React.CSSProperties[];
  onPointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
}

// While editing text, dragging the frame moves the shape, as in PowerPoint.
function FrameGrabHitboxes({ hitboxes, onPointerDown }: FrameGrabHitboxesProps) {
  return (
    <>
      {hitboxes.map((style, index) => (
        <div
          key={index}
          aria-hidden="true"
          data-presentation-border-move=""
          onPointerDown={onPointerDown}
          style={{
            position: "absolute",
            ...style,
            pointerEvents: "auto",
            touchAction: "none",
          }}
        />
      ))}
    </>
  );
}

export namespace SelectionBox {
  export type State = SelectionBoxState;
  export type Props = SelectionBoxProps;
}
