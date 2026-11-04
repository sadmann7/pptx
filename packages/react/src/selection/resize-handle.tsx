import * as React from "react";

import type { StatefulPrimitiveProps } from "../render";
import { renderElement } from "../render";
import { SELECTION_BOX_NAME, SelectionBoxContext, useSelectionParts } from "./context";
import type { GripDirection } from "./geometry";
import { GRIP_DIRECTIONS, getGripCursor } from "./geometry";

const RESIZE_HANDLE_NAME = "Presentation.ResizeHandle";

const GRIP_POSITIONS: Record<GripDirection, React.CSSProperties> = {
  nw: { left: 0, top: 0 },
  n: { left: "50%", top: 0 },
  ne: { left: "100%", top: 0 },
  e: { left: "100%", top: "50%" },
  se: { left: "100%", top: "100%" },
  s: { left: "50%", top: "100%" },
  sw: { left: 0, top: "100%" },
  w: { left: 0, top: "50%" },
};

/** The current state of one resize handle. */
export interface ResizeHandleState {
  /** Which edge or corner the handle drags. */
  direction: GripDirection;
}

export interface ResizeHandleProps extends StatefulPrimitiveProps<"div", ResizeHandleState> {}

/**
 * The resize handles of a selected shape: one per edge and corner, centered
 * on the outline. Hidden while dragging, editing text, and on rotated shapes.
 *
 * Unstyled when composed: give it a size, or it is a zero-size target.
 *
 * Place inside `Presentation.SelectionBox`. Carries `data-resize-grip`, set
 * to its direction.
 */
export const ResizeHandle = React.forwardRef<HTMLDivElement, ResizeHandleProps>(
  function ResizeHandle({ render, className, style, ...resizeHandleProps }, forwardedRef) {
    const { onGripPointerDown } = useSelectionParts(RESIZE_HANDLE_NAME);
    const box = React.useContext(SelectionBoxContext);
    if (!box) {
      throw new Error(`\`${RESIZE_HANDLE_NAME}\` must be used within \`${SELECTION_BOX_NAME}\`.`);
    }
    if (!box.showHandles) return null;

    return GRIP_DIRECTIONS.map((direction) => (
      <React.Fragment key={direction}>
        {renderElement(
          "div",
          { render, className, style },
          {
            state: { direction },
            ref: forwardedRef,
            props: [
              {
                "aria-hidden": true,
                "data-resize-grip": direction,
                onPointerDown: (event: React.PointerEvent<HTMLDivElement>) =>
                  onGripPointerDown(event, direction, box.node.id),
                style: {
                  position: "absolute",
                  ...GRIP_POSITIONS[direction],
                  transform: "translate(-50%, -50%)",
                  cursor: getGripCursor(direction, box.rotation),
                  pointerEvents: "auto",
                  touchAction: "none",
                },
              },
              resizeHandleProps,
            ],
          },
        )}
      </React.Fragment>
    ));
  },
);

export namespace ResizeHandle {
  export type State = ResizeHandleState;
  export type Props = ResizeHandleProps;
}
