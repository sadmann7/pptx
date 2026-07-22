import type { CSSProperties } from "react";

import { AbsoluteFill, Sequence } from "remotion";

import { ShortSlideRight } from "@/components/remocn/short-slide-right";
import { SceneBg } from "@/components/scene-bg";
import { geistSans } from "@/fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as CSSProperties;

const SNAPS = [
  "Shapes, text, tables, charts, themes.",
  "Drag, resize, text editing, undo/redo.",
  "Headless primitives. Your design system.",
  "Client-side only. No server needed.",
];

const SNAP_FRAMES = 25;

export function FeaturesScene() {
  return (
    <AbsoluteFill style={FONT_VARS}>
      <SceneBg />
      {/* First claim lands hard */}
      <Sequence from={0} durationInFrames={SNAP_FRAMES} layout="none">
        <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
          <ShortSlideRight text={SNAPS[0]!} fontSize={56} fontWeight={600} color="#fafafa" />
        </AbsoluteFill>
      </Sequence>

      {/* Subsequent claims swap in */}
      {SNAPS.slice(1).map((snap, i) => (
        <Sequence
          key={snap}
          from={(i + 1) * SNAP_FRAMES}
          durationInFrames={SNAP_FRAMES}
          layout="none"
        >
          <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
            <ShortSlideRight text={snap} fontSize={56} fontWeight={600} color="#fafafa" />
          </AbsoluteFill>
        </Sequence>
      ))}
    </AbsoluteFill>
  );
}
