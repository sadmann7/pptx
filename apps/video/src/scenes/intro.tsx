import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { theme } from "../theme";

const EASE_OUT = Easing.bezier(0.16, 1, 0.3, 1);

export function IntroScene() {
  const frame = useCurrentFrame();

  return (
    <AbsoluteFill
      style={{
        alignItems: "center",
        justifyContent: "center",
        opacity: interpolate(frame, [70, 90], [1, 0], {
          extrapolateLeft: "clamp",
          extrapolateRight: "clamp",
        }),
      }}
    >
      <div
        style={{
          fontSize: 110,
          fontWeight: 700,
          color: theme.text,
          letterSpacing: "-0.03em",
          opacity: interpolate(frame, [0, 25], [0, 1], {
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          }),
          translate: `0px ${interpolate(frame, [0, 25], [40, 0], {
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          })}px`,
        }}
      >
        @diceui/<span style={{ color: theme.accent }}>pptx</span>
      </div>
      <div
        style={{
          fontSize: 38,
          color: theme.muted,
          marginTop: 28,
          opacity: interpolate(frame, [15, 40], [0, 1], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          }),
          translate: `0px ${interpolate(frame, [15, 40], [30, 0], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          })}px`,
        }}
      >
        Render and edit PowerPoint decks in the browser
      </div>
    </AbsoluteFill>
  );
}
