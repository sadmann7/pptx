/**
 * Scene: Perspective Drift
 *
 * A real PPTX deck inside an editor window floats in 3D space.
 * The virtual camera drifts across it — panning, tilting, and
 * dollying — creating cinematic 2.5D motion.
 */
import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { EditorWindow } from "@/components/editor-window";
import { PptxCard } from "@/components/pptx-card";
import { SceneBg } from "@/components/scene-bg";

const clamp = { extrapolateLeft: "clamp", extrapolateRight: "clamp" } as const;
const ease = Easing.bezier(0.22, 1, 0.36, 1);

const progress = (frame: number, from: number, to: number) =>
  interpolate(frame, [from, to], [0, 1], { ...clamp, easing: ease });

const fadeWindow = (frame: number, duration: number, edge = 16) =>
  interpolate(frame, [0, edge, duration - edge, duration], [0, 1, 1, 0], clamp);

export const PERSPECTIVE_DRIFT_FRAMES = 140;

// Editor window fills most of the composition
const WIN_W = 1560;
const WIN_H = 920;
const SLIDE_W = 1120;
const SLIDE_H = 630;

export function PerspectiveDriftScene() {
  const frame = useCurrentFrame();
  const duration = PERSPECTIVE_DRIFT_FRAMES;
  const t = progress(frame, 0, duration - 1);
  const intro = progress(frame, 0, 28);
  const cursorT = progress(frame, 42, 108);

  return (
    <AbsoluteFill>
      <SceneBg />
      <div
        style={{
          position: "absolute",
          left: 180,
          top: 80,
          opacity: intro * fadeWindow(frame, duration),
          transformStyle: "preserve-3d",
          transform: [
            `perspective(1700px)`,
            `translate3d(${interpolate(t, [0, 1], [130, -130])}px, ${interpolate(t, [0, 1], [70, -50])}px, 0)`,
            `rotateX(${interpolate(t, [0, 1], [5, -2])}deg)`,
            `rotateY(${interpolate(t, [0, 1], [-14, 6])}deg)`,
            `scale(${interpolate(intro, [0, 1], [0.85, 1.05])})`,
          ].join(" "),
        }}
      >
        <EditorWindow
          width={WIN_W}
          height={WIN_H}
          filename="editorial-forest.pptx"
          cursor={{
            x: interpolate(cursorT, [0, 1], [900, 560]),
            y: interpolate(cursorT, [0, 1], [620, 360]),
          }}
        >
          <PptxCard
            file="demo.pptx"
            width={SLIDE_W}
            height={SLIDE_H}
            style={{ boxShadow: "0 28px 70px rgba(0,0,0,.48)" }}
          />
        </EditorWindow>
      </div>
    </AbsoluteFill>
  );
}
