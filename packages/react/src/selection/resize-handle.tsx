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

/**
 * Below this many screen px on both axes, handles sit outside the box so its
 * inside stays draggable. A box short on only one axis can still be grabbed
 * between its handles.
 */
const SMALL_SIZE = 24;

/** How far outside a small box its handles sit, in screen px. */
const SMALL_OFFSET = 12;

/**
 * Handles that only change a flat axis (a line's height) are left out, so a
 * line shows its two ends, as in Google Slides.
 */
function getIsGripShown(direction: GripDirection, width: number, height: number): boolean {
  const changesHeight = direction.includes("n") || direction.includes("s");
  const changesWidth = direction.includes("e") || direction.includes("w");
  if (height < 1 && width >= 1) return !changesHeight;
  if (width < 1 && height >= 1) return !changesWidth;
  return true;
}

function getGripOffset(direction: GripDirection, width: number, height: number) {
  if (width >= SMALL_SIZE || height >= SMALL_SIZE) return { x: 0, y: 0 };
  const x = direction.includes("e") ? SMALL_OFFSET : direction.includes("w") ? -SMALL_OFFSET : 0;
  const y = direction.includes("s") ? SMALL_OFFSET : direction.includes("n") ? -SMALL_OFFSET : 0;
  return { x, y };
}

/** The current state of one resize handle. */
export interface ResizeHandleState {
  /** Which edge or corner the handle drags. */
  direction: GripDirection;
}

export interface ResizeHandleProps extends StatefulPrimitiveProps<"div", ResizeHandleState> {}

/**
 * The resize handles of a selected shape: one per edge and corner, centered
 * on the outline. On a shape too small to grab between them they sit just
 * outside it, and a flat line shows only its two ends. Hidden while dragging.
 *
 * Unstyled when composed: give it a size, or it is a zero-size target.
 *
 * Place inside `Presentation.SelectionBox`. Carries `data-resize-handle`, set
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

    return GRIP_DIRECTIONS.filter((direction) =>
      getIsGripShown(direction, box.width, box.height),
    ).map((direction) => {
      const offset = getGripOffset(direction, box.width, box.height);
      return (
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
                  "data-resize-handle": direction,
                  onPointerDown: (event: React.PointerEvent<HTMLDivElement>) =>
                    onGripPointerDown(event, direction, box.node.id),
                  style: {
                    position: "absolute",
                    ...GRIP_POSITIONS[direction],
                    transform: `translate(calc(-50% + ${offset.x}px), calc(-50% + ${offset.y}px))`,
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
      );
    });
  },
);

export namespace ResizeHandle {
  export type State = ResizeHandleState;
  export type Props = ResizeHandleProps;
}
