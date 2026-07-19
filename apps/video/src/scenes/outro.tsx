import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { theme } from "../theme";

const EASE_OUT = Easing.bezier(0.16, 1, 0.3, 1);

export function OutroScene() {
  const frame = useCurrentFrame();

  return (
    <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
      <div
        style={{
          fontSize: 72,
          fontWeight: 700,
          color: theme.text,
          letterSpacing: "-0.02em",
          opacity: interpolate(frame, [0, 20], [0, 1], {
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          }),
          scale: `${interpolate(frame, [0, 20], [0.94, 1], {
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          })}`,
        }}
      >
        @diceui/<span style={{ color: theme.accent }}>pptx</span>
      </div>
      <div
        style={{
          marginTop: 44,
          padding: "20px 44px",
          borderRadius: 12,
          background: theme.surface,
          border: `1px solid ${theme.border}`,
          fontFamily: theme.fontMono,
          fontSize: 32,
          color: theme.text,
          opacity: interpolate(frame, [15, 35], [0, 1], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          }),
        }}
      >
        <span style={{ color: theme.muted }}>$</span> pnpm add @diceui/pptx
      </div>
    </AbsoluteFill>
  );
}
