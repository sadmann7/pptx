import type * as React from "react";

import { AbsoluteFill, interpolate, useCurrentFrame } from "remotion";

import { color } from "@/lib/theme";
import { clamp, CONTENT_FADE, FADE_DURATION } from "@/lib/timing";

function Noise() {
  return (
    <AbsoluteFill
      style={{
        opacity: 0.045,
        backgroundImage:
          'url("data:image/svg+xml,%3Csvg viewBox=%270 0 180 180%27 xmlns=%27http://www.w3.org/2000/svg%27%3E%3Cfilter id=%27n%27%3E%3CfeTurbulence type=%27fractalNoise%27 baseFrequency=%271.1%27 numOctaves=%273%27 stitchTiles=%27stitch%27/%3E%3C/filter%3E%3Crect width=%27100%25%27 height=%27100%25%27 filter=%27url(%23n)%27 opacity=%27.8%27/%3E%3C/svg%3E")',
        pointerEvents: "none",
      }}
    />
  );
}

export function Backdrop() {
  return (
    <AbsoluteFill
      style={{
        color: color.white,
        overflow: "hidden",
        background: [
          "radial-gradient(ellipse 80% 60% at 20% 80%, rgba(124,58,237,.28), transparent)",
          "radial-gradient(ellipse 70% 50% at 80% 20%, rgba(59,130,246,.25), transparent)",
          "radial-gradient(ellipse 60% 40% at 50% 50%, rgba(236,72,153,.15), transparent)",
          color.ink,
        ].join(", "),
      }}
    >
      <AbsoluteFill
        style={{
          opacity: 0.12,
          backgroundImage:
            "linear-gradient(rgba(255,255,255,.06) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,.06) 1px, transparent 1px)",
          backgroundSize: "72px 72px",
          maskImage: "linear-gradient(to bottom, transparent, black 20%, black 80%, transparent)",
        }}
      />
      <Noise />
    </AbsoluteFill>
  );
}

interface SceneContentProps extends React.ComponentProps<"div"> {
  durationInFrames: number;
  fadeOut?: boolean;
}

export function SceneContent({
  durationInFrames,
  fadeOut = true,
  style,
  ...props
}: SceneContentProps) {
  const frame = useCurrentFrame();
  const exitStart = durationInFrames - FADE_DURATION - CONTENT_FADE;
  const opacity = fadeOut
    ? interpolate(frame, [exitStart, exitStart + CONTENT_FADE], [1, 0], clamp)
    : 1;

  return <AbsoluteFill style={{ opacity, ...style }} {...props} />;
}
