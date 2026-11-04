import * as React from "react";

import type { StatefulPrimitiveProps } from "./render";
import { renderElement } from "./render";
import { useSelectionParts } from "./selection-context";
import type { SnapGuide as SnapGuideData } from "./snap";

const SNAP_GUIDE_NAME = "Presentation.SnapGuide";

/** The current state of one snap guide. */
export interface SnapGuideState {
  /** `"x"` for a vertical line, `"y"` for a horizontal one. */
  axis: SnapGuideData["axis"];

  /** Where the line sits on its axis, in slide pixels. */
  position: number;
}

export interface SnapGuideProps extends StatefulPrimitiveProps<"div", SnapGuideState> {}

/**
 * The guides shown while a move or resize snaps: one line per edge or center
 * that lined up, running across the whole slide.
 *
 * The element is the line itself, `--presentation-snap-guide-width` thick
 * (1px by default) and centered on the snapped position, so a background
 * color draws it.
 *
 * Place inside `Presentation.Selection`. Carries `data-snap-guide`, set to
 * its axis.
 */
export const SnapGuide = React.forwardRef<HTMLDivElement, SnapGuideProps>(function SnapGuide(
  { render, className, style, ...snapGuideProps },
  forwardedRef,
) {
  const { state, zoom, slideWidth, slideHeight } = useSelectionParts(SNAP_GUIDE_NAME);
  if (state.mode !== "move" && state.mode !== "resize") return null;

  return (state.guides ?? []).map((guide) => {
    const isVertical = guide.axis === "x";
    // Run the line across the whole slide, as editors do, rather than only
    // between what lined up: a resize snaps the shape's own edge, and a short
    // guide there would just look like its outline changing color. Reach past
    // the slide for shapes that overhang it.
    const start = Math.min(0, guide.start);
    const end = Math.max(isVertical ? slideHeight : slideWidth, guide.end);
    const length = (end - start) * zoom;
    const thickness = "var(--presentation-snap-guide-width, 1px)";
    return (
      <React.Fragment key={`${guide.axis}:${guide.position}`}>
        {renderElement(
          "div",
          { render, className, style },
          {
            state: { axis: guide.axis, position: guide.position },
            ref: forwardedRef,
            props: [
              {
                "aria-hidden": true,
                "data-snap-guide": guide.axis,
                style: {
                  position: "absolute",
                  left: isVertical ? guide.position * zoom : start * zoom,
                  top: isVertical ? start * zoom : guide.position * zoom,
                  width: isVertical ? thickness : length,
                  height: isVertical ? length : thickness,
                  transform: isVertical ? "translateX(-50%)" : "translateY(-50%)",
                  pointerEvents: "none",
                },
              },
              snapGuideProps,
            ],
          },
        )}
      </React.Fragment>
    );
  });
});

export namespace SnapGuide {
  export type State = SnapGuideState;
  export type Props = SnapGuideProps;
}
