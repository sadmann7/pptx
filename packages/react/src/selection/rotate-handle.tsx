import * as React from "react";

import type { StatefulPrimitiveProps } from "../render";
import { renderElement } from "../render";
import { SELECTION_BOX_NAME, SelectionBoxContext, useSelectionParts } from "./context";

/**
 * The cursor over a rotate handle and while rotating: a crosshair, like Google
 * Slides. Apps can swap it with `--presentation-rotate-cursor`.
 */
export const ROTATE_CURSOR = "var(--presentation-rotate-cursor, crosshair)";

const ROTATE_HANDLE_NAME = "Presentation.RotateHandle";

/** The current state of a rotate handle. */
export interface RotateHandleState {
  /** The shape's angle in degrees, including a rotation in progress. */
  rotation: number;

  /** Whether the shape is being rotated. */
  rotating: boolean;
}

export interface RotateHandleProps extends StatefulPrimitiveProps<"div", RotateHandleState> {}

/**
 * The rotation handle of a selected shape, above the middle of its top edge.
 * Drag it to rotate the selection around the shape's center; hold Shift to
 * step by 15°. Rotates with the outline, and is hidden while moving,
 * resizing, editing text, and on tables and charts.
 *
 * Its distance above the outline is `--presentation-rotate-handle-offset`
 * (20px by default). Unstyled when composed: give it a size.
 *
 * Place inside `Presentation.SelectionBox`. Carries `data-rotate-handle`, and
 * `data-rotating` while dragged.
 */
export const RotateHandle = React.forwardRef<HTMLDivElement, RotateHandleProps>(
  function RotateHandle({ render, className, style, ...rotateHandleProps }, forwardedRef) {
    const { state, onRotatePointerDown } = useSelectionParts(ROTATE_HANDLE_NAME);
    const box = React.useContext(SelectionBoxContext);
    if (!box) {
      throw new Error(`\`${ROTATE_HANDLE_NAME}\` must be used within \`${SELECTION_BOX_NAME}\`.`);
    }
    if (!box.showRotateHandle) return null;
    const isRotating = state.mode === "rotate";

    return renderElement(
      "div",
      { render, className, style },
      {
        state: { rotation: box.rotation, rotating: isRotating },
        ref: forwardedRef,
        props: [
          {
            "aria-hidden": true,
            "data-rotate-handle": "",
            "data-rotating": isRotating ? "" : undefined,
            onPointerDown: (event: React.PointerEvent<HTMLDivElement>) =>
              onRotatePointerDown(event, box.node.id),
            style: {
              position: "absolute",
              left: "50%",
              top: "calc(-1 * var(--presentation-rotate-handle-offset, 20px))",
              transform: "translate(-50%, -50%)",
              cursor: ROTATE_CURSOR,
              pointerEvents: "auto",
              touchAction: "none",
            },
          },
          rotateHandleProps,
        ],
      },
    );
  },
);

export namespace RotateHandle {
  export type State = RotateHandleState;
  export type Props = RotateHandleProps;
}
