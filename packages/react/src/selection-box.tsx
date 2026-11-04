import * as React from "react";

import type { SlideNode } from "@diceui/pptx-core";

import type { StatefulPrimitiveProps } from "./render";
import { renderElement } from "./render";
import type { SelectionState } from "./selection";
import { SELECTION_BOX_NAME, SelectionBoxContext, useSelectionParts } from "./selection-context";
import { getNodeRect, normalizeAngle, scaleRectFromGrip } from "./selection-geometry";

const FRAME_GRAB_SIZE = 10;

const FRAME_GRAB_INSET = FRAME_GRAB_SIZE / 2;

/** Top, bottom, left, and right frame grab hitboxes straddling the selection border. */
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

// ---------------------------------------------------------------------------
// Overlay parts
// ---------------------------------------------------------------------------

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

/**
 * The outline around each selected shape. Rendered once per selected shape,
 * following it through moves, resizes, and rotation.
 *
 * Place inside `Presentation.Selection`, with `Presentation.ResizeHandle`
 * inside it for handles.
 *
 * Carries `data-selection-box`, `data-mode`, `data-dragging`, and
 * `data-rotating`.
 */
export const SelectionBox = React.forwardRef<HTMLDivElement, SelectionBoxProps>(
  function SelectionBox(
    { render, className, style, children, ...selectionBoxProps },
    forwardedRef,
  ) {
    const { state, zoom, selectedNodes, resizeScale, showResizeHandles, onBorderPointerDown } =
      useSelectionParts(SELECTION_BOX_NAME);

    const isTextMode = state.mode === "text";
    // Handles hide once a drag is actually under way, not merely on pointer-down:
    // a press that only toggles the selection (Shift+click) would otherwise blink
    // them off and back on.
    const isDragging = state.mode === "move" && state.moved;
    const isRotating = state.mode === "rotate";

    return selectedNodes.map((node) => {
      let rect = getNodeRect(node);
      if (state.mode === "move" && state.moved) {
        rect = { ...rect, x: rect.x + state.dx, y: rect.y + state.dy };
      } else if (state.mode === "resize" && resizeScale) {
        rect = scaleRectFromGrip(rect, state.grip, resizeScale);
      }
      const rotation = isRotating ? normalizeAngle(node.rotation + state.delta) : node.rotation;
      // Resizing works in slide axes, so a rotated shape has no resize handles.
      const showHandles =
        showResizeHandles && node.rotation === 0 && !isDragging && !isRotating && !isTextMode;
      const showRotateHandle =
        showResizeHandles &&
        !isDragging &&
        !isTextMode &&
        state.mode !== "resize" &&
        // PowerPoint doesn't rotate tables or charts.
        node.nodeType !== "table" &&
        node.nodeType !== "chart";
      const boxState: SelectionBoxState = {
        node,
        mode: state.mode,
        dragging: isDragging,
        rotating: isRotating,
        rotation,
      };

      return (
        <SelectionBoxContext.Provider
          key={node.id}
          value={{ node, rotation, showHandles, showRotateHandle }}
        >
          {renderElement(
            "div",
            { render, className, style },
            {
              state: boxState,
              ref: forwardedRef,
              props: [
                {
                  "data-selection-box": "",
                  "data-mode": state.mode,
                  "data-dragging": isDragging ? "" : undefined,
                  "data-rotating": isRotating ? "" : undefined,
                  children: (
                    <>
                      {children}
                      {isTextMode && <FrameGrabHitboxes onPointerDown={onBorderPointerDown} />}
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
  onPointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
}

/**
 * Invisible grab hitboxes straddling the selection frame while editing text.
 * Dragging them moves the shape (PowerPoint: grab the frame to move).
 * Each hitbox extends both inward and outward from the border line.
 */
function FrameGrabHitboxes({ onPointerDown }: FrameGrabHitboxesProps) {
  return (
    <>
      {FRAME_GRAB_HITBOXES.map((style, index) => (
        <div
          key={index}
          aria-hidden="true"
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

export namespace SelectionBox {
  export type State = SelectionBoxState;
  export type Props = SelectionBoxProps;
}
