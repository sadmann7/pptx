import * as React from "react";

import type { StatefulPrimitiveProps } from "../render";
import { renderElement } from "../render";
import { useSelectionParts } from "./context";
import { isBandDrag } from "./geometry";

const MARQUEE_NAME = "Presentation.Marquee";

export interface MarqueeProps extends StatefulPrimitiveProps<"div", Record<string, never>> {}

/**
 * The rubber band drawn while dragging across empty canvas to select.
 *
 * Place inside `Presentation.Selection`. Carries `data-marquee`.
 */
export const Marquee = React.forwardRef<HTMLDivElement, MarqueeProps>(function Marquee(
  { render, className, style, ...marqueeProps },
  forwardedRef,
) {
  const { state, frameRef } = useSelectionParts(MARQUEE_NAME);
  if (state.mode !== "marquee") return null;
  if (!isBandDrag(state.startX, state.startY, state.curX, state.curY)) return null;
  // State coordinates are viewport-relative; the frame is the positioning context.
  const origin = frameRef.current?.getBoundingClientRect();
  if (!origin) return null;

  return renderElement(
    "div",
    { render, className, style },
    {
      state: {},
      ref: forwardedRef,
      props: [
        {
          "data-marquee": "",
          style: {
            position: "absolute",
            left: Math.min(state.startX, state.curX) - origin.left,
            top: Math.min(state.startY, state.curY) - origin.top,
            width: Math.abs(state.curX - state.startX),
            height: Math.abs(state.curY - state.startY),
            pointerEvents: "none",
          },
        },
        marqueeProps,
      ],
    },
  );
});

export namespace Marquee {
  export type Props = MarqueeProps;
}
