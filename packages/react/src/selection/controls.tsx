import * as React from "react";

import type { PrimitiveProps } from "../render";
import { renderElement } from "../render";
import { CONTROLS_ATTRIBUTE, markControlsEvent } from "./channel";

export interface ControlsProps extends PrimitiveProps<"div"> {}

function onControlsEvent(event: React.SyntheticEvent): void {
  markControlsEvent(event.nativeEvent);
}

/**
 * Editor UI outside the slide, such as a toolbar. Pressing it keeps the selection
 * and text editing, and so does pressing the popovers and menus it renders, even
 * when they portal out.
 */
export const Controls = React.forwardRef<HTMLDivElement, ControlsProps>(function Controls(
  { render, ...controlsProps },
  forwardedRef,
) {
  return renderElement(
    "div",
    { render },
    {
      state: {},
      ref: forwardedRef,
      props: [
        {
          [CONTROLS_ATTRIBUTE]: "",
          onPointerDownCapture: onControlsEvent,
          onKeyDownCapture: onControlsEvent,
          onCopyCapture: onControlsEvent,
          onCutCapture: onControlsEvent,
          onPasteCapture: onControlsEvent,
        },
        controlsProps,
      ],
    },
  );
});

export namespace Controls {
  export type Props = ControlsProps;
}
