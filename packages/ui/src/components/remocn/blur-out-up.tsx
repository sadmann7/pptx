"use client";

import { Easing, interpolate, useCurrentFrame, useVideoConfig } from "remotion";

export interface BlurOutUpProps {
  text: string;
  blur?: number;
  fontSize?: number;
  color?: string;
  fontWeight?: number;
  speed?: number;
}

export function BlurOutUp({
  text,
  blur = 12,
  fontSize = 72,
  color = "#fafafa",
  fontWeight = 600,
  speed = 1,
}: BlurOutUpProps) {
  const frame = useCurrentFrame() * speed;
  const { durationInFrames } = useVideoConfig();

  const enterDur = 18;
  const exitDur = 14;
  const exitStart = Math.max(enterDur + 8, durationInFrames - exitDur);

  const enterEase = Easing.bezier(0.22, 1, 0.36, 1);
  const exitEase = Easing.bezier(0.64, 0, 0.78, 0);

  const enterP = interpolate(frame, [0, enterDur], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: enterEase,
  });

  const exitP = interpolate(frame, [exitStart, durationInFrames], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: exitEase,
  });

  const opacity = enterP * (1 - exitP);

  const yEnter = interpolate(frame, [0, enterDur], [10, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: enterEase,
  });

  const yExit = interpolate(frame, [exitStart, durationInFrames], [0, -18], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: exitEase,
  });

  const blurEnter = interpolate(frame, [0, enterDur], [blur * 0.5, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: enterEase,
  });

  const blurExit = interpolate(frame, [exitStart, durationInFrames], [0, blur], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: exitEase,
  });

  const blurVal = blurEnter + blurExit;

  return (
    <div
      style={{
        position: "absolute",
        inset: 0,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      }}
    >
      <span
        style={{
          fontSize,
          fontWeight,
          color,
          letterSpacing: "-0.03em",
          fontFamily: "var(--font-geist-sans), -apple-system, BlinkMacSystemFont, sans-serif",
          display: "inline-block",
          opacity,
          transform: `translateY(${yEnter + yExit}px)`,
          filter: `blur(${blurVal}px)`,
        }}
      >
        {text}
      </span>
    </div>
  );
}
