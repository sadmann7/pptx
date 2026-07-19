import type { ReactNode } from "react";

import { AbsoluteFill, interpolate, useCurrentFrame, useVideoConfig } from "remotion";

interface DriftProps {
  grow?: number;
  children: ReactNode;
}

/**
 * Wraps children in a slow camera push-in so no frame is ever static.
 * `grow` of 0.03-0.05 is the working range.
 */
export function Drift({ grow = 0.04, children }: DriftProps) {
  const frame = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();

  const scale = interpolate(frame, [0, durationInFrames], [1, 1 + grow], {
    extrapolateRight: "clamp",
  });

  return <AbsoluteFill style={{ scale: `${scale}` }}>{children}</AbsoluteFill>;
}
