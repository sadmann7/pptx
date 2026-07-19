import { Easing, interpolate, useCurrentFrame } from "remotion";

export interface CursorKeyframe {
  frame: number;
  x: number;
  y: number;
  pressed?: boolean;
}

interface CursorProps {
  keyframes: CursorKeyframe[];
  appearFrame?: number;
}

export function Cursor({ keyframes, appearFrame = 0 }: CursorProps) {
  const frame = useCurrentFrame();

  if (frame < appearFrame || keyframes.length === 0) return null;

  const frames = keyframes.map((k) => k.frame);
  const xs = keyframes.map((k) => k.x);
  const ys = keyframes.map((k) => k.y);

  const x = interpolate(frame, frames, xs, {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.33, 1, 0.68, 1),
  });
  const y = interpolate(frame, frames, ys, {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.33, 1, 0.68, 1),
  });

  const isPressed = keyframes.some(
    (k, i) =>
      k.pressed &&
      frame >= k.frame &&
      (i === keyframes.length - 1 || frame < keyframes[i + 1]!.frame),
  );

  const opacity = interpolate(frame, [appearFrame, appearFrame + 8], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });

  return (
    <div
      style={{
        position: "absolute",
        left: x,
        top: y,
        opacity,
        scale: isPressed ? "0.9" : "1",
        transition: "scale 0.08s ease",
        pointerEvents: "none",
        zIndex: 9999,
        filter: "drop-shadow(0 2px 4px rgba(0,0,0,0.4))",
      }}
    >
      <svg width="24" height="36" viewBox="0 0 24 36" fill="none">
        <path
          d="M5.5 0L5.5 27.5L11.5 21.5L18.5 33L22 31L15 19.5L23 19.5L5.5 0Z"
          fill="white"
          stroke="black"
          strokeWidth="1.5"
          strokeLinejoin="round"
        />
      </svg>
    </div>
  );
}
