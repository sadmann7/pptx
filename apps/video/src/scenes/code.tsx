import { AbsoluteFill, Easing, interpolate, useCurrentFrame } from "remotion";

import { theme } from "../theme";

const EASE_OUT = Easing.bezier(0.16, 1, 0.3, 1);

const CODE_LINES: { text: string; indent: number }[] = [
  { text: 'import { Presentation } from "@diceui/pptx";', indent: 0 },
  { text: "", indent: 0 },
  { text: "export function Viewer({ file }: { file: File }) {", indent: 0 },
  { text: "return (", indent: 1 },
  { text: "<Presentation.Root file={file}>", indent: 2 },
  { text: "<Presentation.Viewport autoFit>", indent: 3 },
  { text: "<Presentation.Slide />", indent: 4 },
  { text: "</Presentation.Viewport>", indent: 3 },
  { text: "<Presentation.ThumbnailList />", indent: 3 },
  { text: "</Presentation.Root>", indent: 2 },
  { text: ");", indent: 1 },
  { text: "}", indent: 0 },
];

/** Frames between consecutive line reveals. */
const LINE_STAGGER = 5;

export function CodeScene() {
  const frame = useCurrentFrame();

  return (
    <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
      <div
        style={{
          fontSize: 44,
          fontWeight: 600,
          color: theme.text,
          marginBottom: 40,
          opacity: interpolate(frame, [0, 20], [0, 1], {
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          }),
        }}
      >
        Composable primitives, one import
      </div>
      <div
        style={{
          background: theme.surface,
          border: `1px solid ${theme.border}`,
          borderRadius: 16,
          padding: "40px 56px",
          fontFamily: theme.fontMono,
          fontSize: 28,
          lineHeight: 1.7,
          color: theme.text,
          opacity: interpolate(frame, [5, 25], [0, 1], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          }),
          translate: `0px ${interpolate(frame, [5, 25], [30, 0], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          })}px`,
        }}
      >
        {CODE_LINES.map((line, index) => (
          <div
            key={index}
            style={{
              paddingLeft: line.indent * 32,
              minHeight: "1.7em",
              opacity: interpolate(
                frame,
                [20 + index * LINE_STAGGER, 30 + index * LINE_STAGGER],
                [0, 1],
                { extrapolateLeft: "clamp", extrapolateRight: "clamp" },
              ),
            }}
          >
            {line.text}
          </div>
        ))}
      </div>
    </AbsoluteFill>
  );
}
