import * as React from "react";

import type { StatefulPrimitiveProps } from "../render";
import { renderElement } from "../render";
import { SELECTION_BOX_NAME, SelectionBoxContext, useSelectionParts } from "./context";

/** A circular arrow, outlined in white so it reads on any background. */
const ROTATE_CURSOR_SVG =
  "<svg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24'>" +
  "<g fill='none' stroke-linecap='round'>" +
  "<path d='M5 12a7 7 0 1 1 7 7' stroke='white' stroke-width='4'/>" +
  "<path d='M5 12a7 7 0 1 1 7 7' stroke='black' stroke-width='1.75'/>" +
  "</g>" +
  "<path d='M8.5 19 13 15.5v7z' fill='black' stroke='white' stroke-width='1' stroke-linejoin='round'/>" +
  "</svg>";

/**
 * The cursor over a rotate handle and while rotating. Apps can swap it with
 * `--presentation-rotate-cursor`; browsers without SVG cursors fall back to
 * `grab`.
 */
export const ROTATE_CURSOR = `var(--presentation-rotate-cursor, url("data:image/svg+xml,${encodeURIComponent(ROTATE_CURSOR_SVG)}") 12 12, grab)`;

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
