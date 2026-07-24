"use client";

import { Easing, interpolate, useCurrentFrame } from "remotion";

export interface SoftBlurInProps {
  text: string;
  blur?: number;
  fontSize?: number;
  color?: string;
  fontWeight?: number;
  speed?: number;
  className?: string;
}

export function SoftBlurIn({
  text,
  blur = 12,
  fontSize = 72,
  color = "#fafafa",
  fontWeight = 600,
  speed = 1,
  className,
}: SoftBlurInProps) {
  const frame = useCurrentFrame() * speed;

  const chars = Array.from(text);
  const charDurationFrames = 27;
  const staggerFrames = 1;

  return (
    <div
      className={className}
      style={{
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
      }}
    >
      <span
        style={{
          display: "inline-flex",
          flexWrap: "wrap",
          justifyContent: "center",
          fontSize,
          fontWeight,
          fontFamily: "var(--font-geist-sans), -apple-system, BlinkMacSystemFont, sans-serif",
          lineHeight: 1.1,
        }}
      >
        {chars.map((char, i) => {
          const local = frame - i * staggerFrames;
          const easing = Easing.bezier(0.22, 1, 0.36, 1);
          const opacity = interpolate(local, [0, charDurationFrames], [0, 1], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing,
          });
          const y = interpolate(local, [0, charDurationFrames], [16, 0], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing,
          });
          const blurAmount = interpolate(local, [0, charDurationFrames], [blur, 0], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing,
          });
          return (
            <span
              key={i}
              style={{
                display: "inline-block",
                color,
                opacity,
                transform: `translateY(${y}px)`,
                filter: `blur(${blurAmount}px)`,
                whiteSpace: char === " " ? "pre" : undefined,
              }}
            >
              {char}
            </span>
          );
        })}
      </span>
    </div>
  );
}
