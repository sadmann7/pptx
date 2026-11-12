import * as React from "react";

import type { StatefulPrimitiveProps } from "../render";
import { renderElement } from "../render";
import { SELECTION_BOX_NAME, SelectionBoxContext, useSelectionParts } from "./context";

// A crosshair, like Google Slides; apps can swap it with `--presentation-rotate-cursor`.
export const ROTATE_CURSOR = "var(--presentation-rotate-cursor, crosshair)";

const ROTATE_HANDLE_NAME = "Presentation.RotateHandle";

export interface RotateHandleState {
  /** The shape's angle in degrees, including a rotation in progress. */
  rotation: number;

  /** Whether the shape is being rotated. */
  rotating: boolean;
}

export interface RotateHandleProps extends StatefulPrimitiveProps<"div", RotateHandleState> {}

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
            "data-presentation-rotate-handle": "",
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
