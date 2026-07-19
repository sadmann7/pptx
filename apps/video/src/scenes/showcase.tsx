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

import { Drift } from "../components/drift";
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

const FRAMES_PER_DECK = 55;
export const SHOWCASE_FRAMES = DECKS.length * FRAMES_PER_DECK;

function SingleDeckShowcase({ entry }: { entry: DeckEntry }) {
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

  const enterDuration = 18;
  const opacity = interpolate(frame, [0, enterDuration], [0, 1], {
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });
  const scale = interpolate(frame, [0, enterDuration], [0.88, 1], {
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });

  return (
    <AbsoluteFill style={FONT_VARS}>
      <SceneBg />
      <Drift grow={0.03}>
        <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
          <div
            style={{
              opacity,
              transform: `scale(${scale})`,
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
                  if (entry.slideIndex > 0) store.goToIndex(entry.slideIndex);
                  continueRender(handle);
                }}
              >
                <Presentation.Viewport
                  autoFit
                  style={{ width: 1440, height: 810, overflow: "hidden" }}
                >
                  <Presentation.Slide />
                </Presentation.Viewport>
              </Presentation.Root>
            )}
          </div>

          <div
            style={{
              marginTop: 28,
              fontSize: 18,
              fontFamily: geistSans,
              fontWeight: 500,
              color: "#71717a",
              letterSpacing: "0.08em",
              textTransform: "uppercase",
              opacity: interpolate(frame, [10, 25], [0, 1], {
                extrapolateLeft: "clamp",
                extrapolateRight: "clamp",
              }),
            }}
          >
            {entry.label}
          </div>
        </AbsoluteFill>
      </Drift>
    </AbsoluteFill>
  );
}

export function ShowcaseScene() {
  return (
    <AbsoluteFill>
      {DECKS.map((entry, i) => (
        <Sequence key={entry.file} from={i * FRAMES_PER_DECK} durationInFrames={FRAMES_PER_DECK}>
          <SingleDeckShowcase entry={entry} />
        </Sequence>
      ))}
    </AbsoluteFill>
  );
}
