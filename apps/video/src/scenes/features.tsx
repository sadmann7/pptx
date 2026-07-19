import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { theme } from "../theme";

const EASE_OUT = Easing.bezier(0.16, 1, 0.3, 1);

const FEATURES = [
  "Client-side rendering, no server round-trips",
  "Editing with drag, resize, and text, plus undo/redo",
  "Shapes, tables, charts, groups, and themes",
  "Headless primitives that compose like Radix",
];

/** Frames between consecutive feature reveals. */
const ITEM_STAGGER = 14;

export function FeaturesScene() {
  const frame = useCurrentFrame();

  return (
    <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
      <div
        style={{
          fontSize: 44,
          fontWeight: 600,
          color: theme.text,
          marginBottom: 56,
          opacity: interpolate(frame, [0, 20], [0, 1], {
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          }),
        }}
      >
        Built for the web
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 28 }}>
        {FEATURES.map((feature, index) => (
          <div
            key={feature}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 24,
              fontSize: 34,
              color: theme.text,
              opacity: interpolate(
                frame,
                [15 + index * ITEM_STAGGER, 35 + index * ITEM_STAGGER],
                [0, 1],
                { extrapolateLeft: "clamp", extrapolateRight: "clamp", easing: EASE_OUT },
              ),
              translate: `${interpolate(
                frame,
                [15 + index * ITEM_STAGGER, 35 + index * ITEM_STAGGER],
                [40, 0],
                { extrapolateLeft: "clamp", extrapolateRight: "clamp", easing: EASE_OUT },
              )}px 0px`,
            }}
          >
            <span
              style={{
                width: 14,
                height: 14,
                borderRadius: 7,
                background: theme.accent,
                flexShrink: 0,
              }}
            />
            {feature}
          </div>
        ))}
      </div>
    </AbsoluteFill>
  );
}
