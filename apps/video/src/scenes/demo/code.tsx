import type { CSSProperties } from "react";

import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { GlassCodeBlock } from "@/components/remocn/glass-code-block";
import { SceneBg } from "@/components/scene-bg";
import { geistMono, geistSans } from "@/lib/fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
  "--font-geist-mono": geistMono,
} as CSSProperties;

const CODE = `import { Presentation } from "@diceui/pptx";

export function Viewer({ file }) {
  return (
    <Presentation.Root file={file}>
      <Presentation.Viewport autoFit>
        <Presentation.Slide />
      </Presentation.Viewport>
      <Presentation.ThumbnailList />
    </Presentation.Root>
  );
}`;

export function CodeScene() {
  const frame = useCurrentFrame();

  const titleOpacity = interpolate(frame, [0, 18], [0, 1], {
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });
  const titleY = interpolate(frame, [0, 18], [20, 0], {
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
          gap: 48,
        }}
      >
        <div
          style={{
            fontSize: 44,
            fontFamily: geistSans,
            fontWeight: 600,
            color: "#fafafa",
            letterSpacing: "-0.01em",
            opacity: titleOpacity,
            transform: `translateY(${titleY}px)`,
          }}
        >
          One import. Composable primitives.
        </div>

        <GlassCodeBlock
          title="viewer.tsx"
          code={CODE}
          width={820}
          height={500}
          fontSize={18}
          staggerFrames={3}
          aura
        />
      </AbsoluteFill>
    </AbsoluteFill>
  );
}
