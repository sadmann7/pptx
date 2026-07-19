import type { CSSProperties } from "react";

import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { SoftBlurIn } from "../components/remocn/soft-blur-in";
import { SceneBg } from "../components/scene-bg";
import { geistMono, geistSans } from "../fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
  "--font-geist-mono": geistMono,
} as CSSProperties;

export function OutroScene() {
  const frame = useCurrentFrame();

  const installOpacity = interpolate(frame, [18, 36], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });
  const installY = interpolate(frame, [18, 36], [14, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });

  const urlOpacity = interpolate(frame, [30, 48], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });

  return (
    <AbsoluteFill style={FONT_VARS}>
      <SceneBg />
      <AbsoluteFill
        style={{
          alignItems: "center",
          justifyContent: "center",
          flexDirection: "column",
          gap: 36,
        }}
      >
        <SoftBlurIn text="@diceui/pptx" fontSize={88} fontWeight={700} color="#fafafa" />

        <div
          style={{
            padding: "16px 40px",
            borderRadius: 10,
            background: "rgba(24, 24, 27, 0.8)",
            border: "1px solid rgba(255,255,255,0.08)",
            fontFamily: geistMono,
            fontSize: 28,
            color: "#fafafa",
            opacity: installOpacity,
            transform: `translateY(${installY}px)`,
          }}
        >
          <span style={{ color: "#71717a" }}>$</span> pnpm add @diceui/pptx
        </div>

        <div
          style={{
            fontSize: 20,
            fontFamily: geistSans,
            fontWeight: 500,
            color: "#71717a",
            letterSpacing: "0.06em",
            textTransform: "uppercase",
            opacity: urlOpacity,
          }}
        >
          Open source &middot; diceui.com
        </div>
      </AbsoluteFill>
    </AbsoluteFill>
  );
}
