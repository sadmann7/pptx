import * as React from "react";
import type { CSSProperties } from "react";

import type { PresentationStore } from "@diceui/pptx";
import { Presentation } from "@diceui/pptx";
import {
  AbsoluteFill,
  continueRender,
  delayRender,
  Easing,
  interpolate,
  staticFile,
  useCurrentFrame,
} from "remotion";

import type { CursorKeyframe } from "../components/cursor";
import { Cursor } from "../components/cursor";
import { SceneBg } from "../components/scene-bg";
import { geistSans } from "../fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as CSSProperties;

const VIEWPORT_W = 1440;
const VIEWPORT_H = 810;
const VIEWPORT_LEFT = (1920 - VIEWPORT_W) / 2;
const VIEWPORT_TOP = 100;

const CURSOR_KEYFRAMES: CursorKeyframe[] = [
  { frame: 10, x: 960, y: 600 },
  { frame: 30, x: 700, y: 350 },
  { frame: 40, x: 700, y: 350, pressed: true },
  { frame: 80, x: 1100, y: 300, pressed: true },
  { frame: 85, x: 1100, y: 300 },
  { frame: 100, x: 500, y: 450 },
  { frame: 115, x: 550, y: 480 },
  { frame: 120, x: 550, y: 480, pressed: true },
  { frame: 160, x: 700, y: 550, pressed: true },
  { frame: 165, x: 700, y: 550 },
  { frame: 180, x: 800, y: 400 },
  { frame: 195, x: 850, y: 420, pressed: true },
  { frame: 200, x: 850, y: 420 },
  { frame: 230, x: 850, y: 420 },
];

const ACTION_LABELS: { start: number; end: number; text: string }[] = [
  { start: 30, end: 90, text: "Drag shapes" },
  { start: 100, end: 170, text: "Resize elements" },
  { start: 180, end: 240, text: "Edit text inline" },
];

export const INTERACTION_FRAMES = 240;

export function InteractionScene() {
  const frame = useCurrentFrame();
  const [data, setData] = React.useState<ArrayBuffer | null>(null);
  const storeRef = React.useRef<PresentationStore | null>(null);
  const [handle] = React.useState(() => delayRender("load interaction deck"));

  React.useEffect(() => {
    let cancelled = false;
    fetch(staticFile("demo.pptx"))
      .then((r) => r.arrayBuffer())
      .then((buf) => {
        if (!cancelled) setData(buf);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const containerOpacity = interpolate(frame, [0, 15], [0, 1], {
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });

  const activeLabel = ACTION_LABELS.find((l) => frame >= l.start && frame < l.end);
  const labelOpacity = activeLabel
    ? interpolate(frame, [activeLabel.start, activeLabel.start + 12], [0, 1], {
        extrapolateLeft: "clamp",
        extrapolateRight: "clamp",
      })
    : 0;

  return (
    <AbsoluteFill style={FONT_VARS}>
      <SceneBg />
      <div
        style={{
          position: "absolute",
          left: VIEWPORT_LEFT,
          top: VIEWPORT_TOP,
          opacity: containerOpacity,
          borderRadius: 14,
          overflow: "hidden",
          boxShadow: "0 40px 100px rgba(0,0,0,0.6)",
        }}
      >
        {data && (
          <Presentation.Root
            file={data}
            onLoad={(store) => {
              storeRef.current = store;
              continueRender(handle);
            }}
          >
            <Presentation.Viewport
              autoFit
              style={{ width: VIEWPORT_W, height: VIEWPORT_H, overflow: "hidden" }}
            >
              <Presentation.Slide />
            </Presentation.Viewport>
          </Presentation.Root>
        )}
      </div>

      <Cursor keyframes={CURSOR_KEYFRAMES} appearFrame={10} />

      {activeLabel && (
        <div
          style={{
            position: "absolute",
            bottom: 80,
            left: 0,
            right: 0,
            textAlign: "center",
            fontSize: 24,
            fontFamily: geistSans,
            fontWeight: 600,
            color: "#3b82f6",
            letterSpacing: "0.06em",
            textTransform: "uppercase",
            opacity: labelOpacity,
          }}
        >
          {activeLabel.text}
        </div>
      )}
    </AbsoluteFill>
  );
}
