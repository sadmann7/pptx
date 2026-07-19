import { Easing, interpolate, useCurrentFrame } from "remotion";

export interface ShortSlideRightProps {
  text: string;
  distance?: number;
  staggerDelay?: number;
  fontSize?: number;
  color?: string;
  fontWeight?: number;
  speed?: number;
  className?: string;
}

export function ShortSlideRight({
  text,
  distance = 24,
  staggerDelay = 3,
  fontSize = 72,
  color = "#fafafa",
  fontWeight = 600,
  speed = 1,
  className,
}: ShortSlideRightProps) {
  const frame = useCurrentFrame() * speed;

  const words = text.split(" ");

  const enterDur = 16;
  const wordOpacityDur = 6;
  const easing = Easing.bezier(0.2, 0.8, 0.2, 1);

  const x = interpolate(frame, [0, enterDur], [-distance, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing,
  });

  const blurVal = interpolate(frame, [0, enterDur], [1.2, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
    easing,
  });

  return (
    <div
      className={className}
      style={{
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        width: "100%",
        height: "100%",
      }}
    >
      <span
        style={{
          display: "inline-flex",
          gap: "0.3em",
          fontSize,
          fontWeight,
          fontFamily: "var(--font-geist-sans), -apple-system, BlinkMacSystemFont, sans-serif",
          transform: `translateX(${x}px)`,
          filter: `blur(${blurVal}px)`,
        }}
      >
        {words.map((word, i) => {
          const opacity = interpolate(frame - i * staggerDelay, [0, wordOpacityDur], [0, 1], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing,
          });
          return (
            <span key={i} style={{ color, opacity }}>
              {word}
            </span>
          );
        })}
      </span>
    </div>
  );
}
