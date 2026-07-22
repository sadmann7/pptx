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
  Sequence,
  staticFile,
  useCurrentFrame,
} from "remotion";

import { SceneBg } from "../components/scene-bg";
import { geistSans } from "../fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as CSSProperties;

interface DeckEntry {
  file: string;
  label: string;
  slideIndex: number;
}

const DECKS: DeckEntry[] = [
  { file: "demo.pptx", label: "Editorial Forest", slideIndex: 0 },
  { file: "sakura-chroma.pptx", label: "Sakura Chroma", slideIndex: 0 },
  { file: "retro-windows.pptx", label: "Retro Windows", slideIndex: 0 },
  { file: "cobalt-grid.pptx", label: "Cobalt Grid", slideIndex: 0 },
  { file: "bold-poster.pptx", label: "Bold Poster", slideIndex: 0 },
  { file: "playful.pptx", label: "Playful", slideIndex: 0 },
];

const CARD_W = 1100;
const CARD_H = 619;

const FRAMES_PER_DECK = 55;
export const SHOWCASE_FRAMES = DECKS.length * FRAMES_PER_DECK;

// Each deck gets a different camera move for variety
interface CameraPreset {
  // Card starting tilt
  rotateY: number;
  rotateX: number;
  // Camera motion over the hold duration
  panX: [number, number]; // start → end translateX
  panY: [number, number]; // start → end translateY
  dolly: [number, number]; // start → end scale (simulates dolly in/out)
  driftRotateY: [number, number]; // subtle rotation drift
  originX: number; // perspective-origin X offset from 50%
}

const PRESETS: CameraPreset[] = [
  // Dolly in + slight left pan
  {
    rotateY: -16,
    rotateX: 5,
    panX: [140, 90],
    panY: [0, -15],
    dolly: [0.95, 1.05],
    driftRotateY: [-16, -12],
    originX: 55,
  },
  // Right tilt, slow truck right
  {
    rotateY: 18,
    rotateX: -4,
    panX: [-130, -60],
    panY: [10, -5],
    dolly: [1.0, 1.08],
    driftRotateY: [18, 14],
    originX: 42,
  },
  // Subtle tilt, dolly out + drift up
  {
    rotateY: -12,
    rotateX: 6,
    panX: [80, 120],
    panY: [20, -20],
    dolly: [1.06, 0.98],
    driftRotateY: [-12, -15],
    originX: 54,
  },
  // Right tilt, truck left
  {
    rotateY: 20,
    rotateX: -3,
    panX: [-80, -160],
    panY: [-10, 10],
    dolly: [0.98, 1.06],
    driftRotateY: [20, 16],
    originX: 40,
  },
  // Gentle left, dolly in close
  {
    rotateY: -14,
    rotateX: 4,
    panX: [100, 60],
    panY: [15, -10],
    dolly: [0.96, 1.1],
    driftRotateY: [-14, -10],
    originX: 56,
  },
  // Right, slow drift
  {
    rotateY: 15,
    rotateX: -5,
    panX: [-100, -50],
    panY: [-5, -20],
    dolly: [1.02, 1.08],
    driftRotateY: [15, 12],
    originX: 44,
  },
];

function SingleDeckShowcase({ entry, presetIndex }: { entry: DeckEntry; presetIndex: number }) {
  const frame = useCurrentFrame();
  const [data, setData] = React.useState<ArrayBuffer | null>(null);
  const storeRef = React.useRef<PresentationStore | null>(null);
  const [handle] = React.useState(() => delayRender(`load ${entry.file}`));

  React.useEffect(() => {
    let cancelled = false;
    fetch(staticFile(entry.file))
      .then((r) => r.arrayBuffer())
      .then((buf) => {
        if (!cancelled) setData(buf);
      });
    return () => {
      cancelled = true;
    };
  }, [entry.file]);

  const preset = PRESETS[presetIndex % PRESETS.length]!;
  const ease = Easing.bezier(0.25, 0.46, 0.45, 0.94);

  // Entrance (first 18 frames): fade + scale from below
  const enterDur = 18;
  const enterOpacity = interpolate(frame, [0, enterDur], [0, 1], {
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });
  const enterFromBelow = interpolate(frame, [0, enterDur], [50, 0], {
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });

  // Continuous camera motion over entire duration
  const progress = interpolate(frame, [0, FRAMES_PER_DECK], [0, 1], {
    extrapolateRight: "clamp",
    easing: ease,
  });

  const panX = interpolate(progress, [0, 1], preset.panX);
  const panY = interpolate(progress, [0, 1], preset.panY) + enterFromBelow;
  const dolly = interpolate(progress, [0, 1], preset.dolly);
  const rotateY = interpolate(progress, [0, 1], preset.driftRotateY);
  const rotateX = preset.rotateX;

  // Label
  const labelOpacity = interpolate(frame, [12, 24], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });

  return (
    <AbsoluteFill style={FONT_VARS}>
      <SceneBg />

      <div
        style={{
          position: "absolute",
          width: "100%",
          height: "100%",
          perspective: 1200,
          perspectiveOrigin: `${preset.originX}% 48%`,
        }}
      >
        <div
          style={{
            position: "absolute",
            width: CARD_W,
            height: CARD_H,
            left: "50%",
            top: "50%",
            marginLeft: -CARD_W / 2,
            marginTop: -CARD_H / 2 - 20,
            opacity: enterOpacity,
            transform: [
              `translateX(${panX}px)`,
              `translateY(${panY}px)`,
              `scale(${dolly})`,
              `rotateY(${rotateY}deg)`,
              `rotateX(${rotateX}deg)`,
            ].join(" "),
            transformStyle: "preserve-3d",
            borderRadius: 12,
            overflow: "hidden",
            boxShadow: [
              "0 4px 6px rgba(0,0,0,0.1)",
              "0 20px 40px rgba(0,0,0,0.3)",
              "0 50px 100px rgba(0,0,0,0.4)",
            ].join(", "),
          }}
        >
          {data && (
            <Presentation.Root
              file={data}
              onLoad={(store) => {
                storeRef.current = store;
                if (entry.slideIndex > 0) store.goToIndex(entry.slideIndex);
                continueRender(handle);
              }}
            >
              <Presentation.Viewport
                autoFit
                style={{ width: CARD_W, height: CARD_H, overflow: "hidden" }}
              >
                <Presentation.Slide />
              </Presentation.Viewport>
            </Presentation.Root>
          )}

          {/* Specular highlight on the light-facing edge */}
          <div
            style={{
              position: "absolute",
              inset: 0,
              background:
                preset.rotateY < 0
                  ? "linear-gradient(105deg, rgba(255,255,255,0.06) 0%, transparent 40%)"
                  : "linear-gradient(255deg, rgba(255,255,255,0.06) 0%, transparent 40%)",
              pointerEvents: "none",
            }}
          />
        </div>
      </div>

      <div
        style={{
          position: "absolute",
          bottom: 80,
          left: 0,
          right: 0,
          textAlign: "center",
          fontSize: 18,
          fontFamily: geistSans,
          fontWeight: 500,
          color: "#71717a",
          letterSpacing: "0.1em",
          textTransform: "uppercase",
          opacity: labelOpacity,
        }}
      >
        {entry.label}
      </div>
    </AbsoluteFill>
  );
}

export function ShowcaseScene() {
  return (
    <AbsoluteFill>
      {DECKS.map((entry, i) => (
        <Sequence key={entry.file} from={i * FRAMES_PER_DECK} durationInFrames={FRAMES_PER_DECK}>
          <SingleDeckShowcase entry={entry} presetIndex={i} />
        </Sequence>
      ))}
    </AbsoluteFill>
  );
}
