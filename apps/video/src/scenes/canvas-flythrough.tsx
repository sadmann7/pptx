/**
 * Scene: Canvas Flythrough
 *
 * All slides are scattered on a large canvas. A single camera
 * flies through the space, visiting each slide with smooth
 * eased stops, zoom changes, and tilts. Dashed connector lines
 * link the slides.
 */
import type { CSSProperties } from "react";

import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { PptxCard } from "@/components/pptx-card";
import { SceneBg } from "@/components/scene-bg";
import { geistSans } from "@/fonts";

const clamp = { extrapolateLeft: "clamp", extrapolateRight: "clamp" } as const;
const ease = Easing.bezier(0.22, 1, 0.36, 1);

const font: CSSProperties = { fontFamily: geistSans };

interface CanvasSlide {
  file: string;
  x: number;
  y: number;
}

// Larger slides on the canvas
const SLIDE_W = 1120;
const SLIDE_H = 630;

const SLIDES: CanvasSlide[] = [
  { file: "demo.pptx", x: 0, y: 0 },
  { file: "sakura-chroma.pptx", x: 1400, y: 200 },
  { file: "cobalt-grid.pptx", x: 2500, y: 1100 },
  { file: "bold-poster.pptx", x: 900, y: 1300 },
];

export const CANVAS_FLYTHROUGH_FRAMES = 180;

export function CanvasFlythroughScene() {
  const frame = useCurrentFrame();
  const duration = CANVAS_FLYTHROUGH_FRAMES;

  // More dynamic camera with bigger sweeps and zoom changes
  const stops = [0, duration * 0.28, duration * 0.58, duration * 0.85, duration - 1];
  const cameraX = interpolate(frame, stops, [80, -1000, -1600, -500, -500], {
    ...clamp,
    easing: ease,
  });
  const cameraY = interpolate(frame, stops, [20, -120, -750, -1100, -1100], {
    ...clamp,
    easing: ease,
  });
  const zoom = interpolate(frame, stops, [0.92, 0.78, 0.65, 0.82, 0.82], {
    ...clamp,
    easing: ease,
  });
  const tilt = interpolate(frame, stops, [0, -1.6, 1.4, -0.8, -0.8], {
    ...clamp,
    easing: ease,
  });

  // Entrance fade
  const enterOpacity = interpolate(frame, [0, 18], [0, 1], clamp);

  return (
    <AbsoluteFill>
      <SceneBg />
      <div
        style={{
          position: "absolute",
          left: 400,
          top: 240,
          width: 3800,
          height: 2200,
          transformOrigin: "0 0",
          opacity: enterOpacity,
          transform: [
            `perspective(1800px)`,
            `translate3d(${cameraX}px, ${cameraY}px, 0)`,
            `rotateX(${tilt * 0.5}deg)`,
            `rotateZ(${tilt}deg)`,
            `scale(${zoom})`,
          ].join(" "),
        }}
      >
        {SLIDES.map((slide) => (
          <div
            key={slide.file}
            style={{
              position: "absolute",
              left: slide.x,
              top: slide.y,
              width: SLIDE_W,
              height: SLIDE_H,
              borderRadius: 12,
              overflow: "hidden",
              boxShadow: "0 40px 90px rgba(0,0,0,.5)",
            }}
          >
            <PptxCard file={slide.file} width={SLIDE_W} height={SLIDE_H} />
          </div>
        ))}

        {/* Connector lines */}
        <svg width="3600" height="2000" style={{ position: "absolute", inset: 0, zIndex: -1 }}>
          {/* Slide 0 → Slide 1 */}
          <path
            d={`M${SLIDE_W} ${SLIDE_H * 0.45} C${SLIDE_W + 180} ${SLIDE_H * 0.45} ${1400 - 180} ${200 + SLIDE_H * 0.5} ${1400} ${200 + SLIDE_H * 0.5}`}
            fill="none"
            stroke="#3b82f6"
            strokeWidth="3"
            strokeDasharray="10 12"
            opacity=".4"
          />
          {/* Slide 1 → Slide 2 */}
          <path
            d={`M${1400 + SLIDE_W * 0.6} ${200 + SLIDE_H} C${1400 + SLIDE_W * 0.6} ${200 + SLIDE_H + 200} ${2500 + SLIDE_W * 0.3} ${1100 - 100} ${2500 + SLIDE_W * 0.3} ${1100}`}
            fill="none"
            stroke="#3b82f6"
            strokeWidth="3"
            strokeDasharray="10 12"
            opacity=".4"
          />
          {/* Slide 2 → Slide 3 */}
          <path
            d={`M${2500} ${1100 + SLIDE_H * 0.5} C${2100} ${1100 + SLIDE_H * 0.5 + 100} ${900 + SLIDE_W + 200} ${1300 + SLIDE_H * 0.4} ${900 + SLIDE_W} ${1300 + SLIDE_H * 0.4}`}
            fill="none"
            stroke="#3b82f6"
            strokeWidth="3"
            strokeDasharray="10 12"
            opacity=".4"
          />
        </svg>
      </div>

      <div
        style={{
          ...font,
          position: "absolute",
          right: 70,
          bottom: 54,
          fontSize: 16,
          color: "#71717a",
          opacity: interpolate(frame, [20, 40], [0, 1], clamp),
        }}
      >
        One camera · multiple decks · zero plugins
      </div>
    </AbsoluteFill>
  );
}
