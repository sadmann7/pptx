import type { CSSProperties } from "react";

import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { SoftBlurIn } from "@/components/remocn/soft-blur-in";
import { SceneBg } from "@/components/scene-bg";
import { geistSans } from "@/fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as CSSProperties;

export function IntroScene() {
  const frame = useCurrentFrame();

  const taglineOpacity = interpolate(frame, [25, 45], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });
  const taglineY = interpolate(frame, [25, 45], [16, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });

  return (
    <AbsoluteFill style={FONT_VARS}>
      <SceneBg />
      <AbsoluteFill
        style={{
          alignItems: "center",
          justifyContent: "center",
          flexDirection: "column",
          gap: 28,
        }}
      >
        <SoftBlurIn
          text="PowerPoint in the Browser"
          fontSize={96}
          fontWeight={700}
          color="#fafafa"
        />

        <div
          style={{
            fontSize: 26,
            fontFamily: geistSans,
            fontWeight: 500,
            color: "#a1a1aa",
            letterSpacing: "0.04em",
            opacity: taglineOpacity,
            transform: `translateY(${taglineY}px)`,
          }}
        >
          Render, edit, and compose .pptx files with React
        </div>
      </AbsoluteFill>
    </AbsoluteFill>
  );
}
