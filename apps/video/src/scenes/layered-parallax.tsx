/**
 * Scene 2: Layered Parallax
 *
 * Multiple real PPTX slides float at different depths. The camera pans
 * laterally, and each layer moves at a different speed based on its
 * depth — classic parallax.
 */
import type { CSSProperties } from "react";

import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { PptxCard } from "../components/pptx-card";
import { SceneBg } from "../components/scene-bg";
import { geistSans } from "../fonts";

const clamp = { extrapolateLeft: "clamp", extrapolateRight: "clamp" } as const;
const ease = Easing.bezier(0.22, 1, 0.36, 1);

const progress = (frame: number, from: number, to: number) =>
  interpolate(frame, [from, to], [0, 1], { ...clamp, easing: ease });

const fadeWindow = (frame: number, duration: number, edge = 16) =>
  interpolate(frame, [0, edge, duration - edge, duration], [0, 1, 1, 0], clamp);

const font: CSSProperties = { fontFamily: geistSans };

interface ParallaxLayer {
  file: string;
  x: number;
  y: number;
  depth: number;
  rotate: number;
  label: string;
}

const LAYERS: ParallaxLayer[] = [
  { file: "demo.pptx", x: 80, y: 160, depth: 0.5, rotate: -7, label: "Parsing" },
  { file: "sakura-chroma.pptx", x: 620, y: 80, depth: 0.85, rotate: 4, label: "Rendering" },
  { file: "retro-windows.pptx", x: 1100, y: 320, depth: 1.2, rotate: -3, label: "Editing" },
];

function FeatureTag({ children, style }: { children: React.ReactNode; style?: CSSProperties }) {
  return (
    <div
      style={{
        ...font,
        padding: "10px 16px",
        borderRadius: 999,
        border: "1px solid #3b82f6",
        background: "rgba(9,9,11,0.86)",
        color: "#3b82f6",
        fontWeight: 700,
        fontSize: 16,
        letterSpacing: 0.2,
        boxShadow: "0 14px 40px rgba(0,0,0,.35)",
        ...style,
      }}
    >
      {children}
    </div>
  );
}

export const LAYERED_PARALLAX_FRAMES = 140;

export function LayeredParallaxScene() {
  const frame = useCurrentFrame();
  const duration = LAYERED_PARALLAX_FRAMES;
  const t = progress(frame, 0, duration - 1);
  const opacity = fadeWindow(frame, duration);

  return (
    <AbsoluteFill>
      <SceneBg />
      <div style={{ opacity }}>
        {LAYERS.map((layer, index) => (
          <div
            key={layer.file}
            style={{
              position: "absolute",
              left: layer.x,
              top: layer.y,
              width: 760,
              height: 428,
              transform: [
                `translate3d(${interpolate(t, [0, 1], [120 * layer.depth, -150 * layer.depth])}px, ${interpolate(t, [0, 1], [70 * layer.depth, -55 * layer.depth])}px, 0)`,
                `rotate(${layer.rotate + interpolate(t, [0, 1], [-1.3, 1.3])}deg)`,
                `scale(${0.88 + index * 0.03})`,
              ].join(" "),
              borderRadius: 10,
              overflow: "hidden",
              boxShadow: "0 60px 100px rgba(0,0,0,.48)",
            }}
          >
            <PptxCard file={layer.file} width={760} height={428} />
          </div>
        ))}

        {/* Floating feature tags that move at their own parallax rate */}
        <FeatureTag
          style={{
            position: "absolute",
            left: 370 + interpolate(t, [0, 1], [90, -190]),
            top: 700,
          }}
        >
          {LAYERS[0]!.label}
        </FeatureTag>
        <FeatureTag
          style={{
            position: "absolute",
            left: 920 + interpolate(t, [0, 1], [170, -240]),
            top: 180,
          }}
        >
          {LAYERS[1]!.label}
        </FeatureTag>
        <FeatureTag
          style={{
            position: "absolute",
            left: 1310 + interpolate(t, [0, 1], [240, -320]),
            top: 790,
          }}
        >
          {LAYERS[2]!.label}
        </FeatureTag>
      </div>
    </AbsoluteFill>
  );
}
