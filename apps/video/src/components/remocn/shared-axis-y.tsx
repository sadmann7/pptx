import { Easing, interpolate, useCurrentFrame } from "remotion";

export interface SharedAxisYProps {
  fromText: string;
  toText: string;
  fontSize?: number;
  color?: string;
  fontWeight?: number;
  speed?: number;
  className?: string;
}

export function SharedAxisY({
  fromText,
  toText,
  fontSize = 72,
  color = "#fafafa",
  fontWeight = 600,
  speed = 1,
  className,
}: SharedAxisYProps) {
  const frame = useCurrentFrame() * speed;

  const fromWords = fromText.split(" ");
  const toWords = toText.split(" ");

  const enterDur = 5;
  const exitDur = 4;
  const enterStagger = 2;
  const exitStagger = 2;
  const overlapF = 0;
  const microDelayF = 1;

  const exitTotal = exitDur + (fromWords.length - 1) * exitStagger;
  const newStart = Math.max(0, exitTotal - overlapF + microDelayF);

  const fontStack = "var(--font-geist-sans), -apple-system, BlinkMacSystemFont, sans-serif";

  return (
    <div
      className={className}
      style={{
        position: "relative",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        width: "100%",
        height: "100%",
      }}
    >
      {/* Exiting text */}
      <div
        style={{
          position: "absolute",
          display: "flex",
          gap: "0.3em",
          fontSize,
          fontWeight,
          fontFamily: fontStack,
          color,
        }}
      >
        {fromWords.map((word, i) => {
          const local = frame - i * exitStagger;
          const opacity = interpolate(local, [0, exitDur], [1, 0], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing: Easing.step1,
          });
          return (
            <span key={`from-${i}`} style={{ opacity }}>
              {word}
            </span>
          );
        })}
      </div>

      {/* Entering text */}
      <div
        style={{
          position: "absolute",
          display: "flex",
          gap: "0.3em",
          fontSize,
          fontWeight,
          fontFamily: fontStack,
          color,
        }}
      >
        {toWords.map((word, j) => {
          const local = frame - newStart - j * enterStagger;
          const opacity = interpolate(local, [0, enterDur], [0, 1], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing: Easing.step1,
          });
          return (
            <span key={`to-${j}`} style={{ opacity }}>
              {word}
            </span>
          );
        })}
      </div>
    </div>
  );
}
