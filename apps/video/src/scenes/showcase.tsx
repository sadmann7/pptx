import * as React from "react";

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

import { theme } from "../theme";

const EASE_OUT = Easing.bezier(0.16, 1, 0.3, 1);

/** How long each slide stays on screen before advancing. */
const FRAMES_PER_SLIDE = 80;

/**
 * The real library rendering a real deck: this scene mounts
 * `@diceui/pptx` and drives slide navigation from the video timeline,
 * so what the video shows is the actual renderer output, not a mockup.
 */
export function ShowcaseScene() {
  const frame = useCurrentFrame();

  const [deck, setDeck] = React.useState<ArrayBuffer | null>(null);
  const storeRef = React.useRef<PresentationStore | null>(null);
  const [loadHandle] = React.useState(() => delayRender("load demo.pptx"));

  React.useEffect(() => {
    let cancelled = false;
    fetch(staticFile("demo.pptx"))
      .then((response) => response.arrayBuffer())
      .then((buffer) => {
        if (!cancelled) setDeck(buffer);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  // Drive navigation from the timeline: the store is an external system,
  // so syncing it to the current frame belongs in an effect.
  const slideIndex = Math.floor(frame / FRAMES_PER_SLIDE);
  React.useEffect(() => {
    storeRef.current?.goToIndex(slideIndex);
  }, [slideIndex]);

  return (
    <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
      <div
        style={{
          opacity: interpolate(frame, [0, 20], [0, 1], {
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          }),
          scale: `${interpolate(frame, [0, 20], [0.94, 1], {
            extrapolateRight: "clamp",
            easing: EASE_OUT,
          })}`,
          borderRadius: 16,
          overflow: "hidden",
          border: `1px solid ${theme.border}`,
          boxShadow: "0 40px 120px rgba(0, 0, 0, 0.6)",
        }}
      >
        {deck && (
          <Presentation.Root
            file={deck}
            onLoad={(store) => {
              storeRef.current = store;
              store.goToIndex(slideIndex);
              continueRender(loadHandle);
            }}
          >
            <Presentation.Viewport autoFit style={{ width: 1440, height: 810, overflow: "hidden" }}>
              <Presentation.Slide />
            </Presentation.Viewport>
          </Presentation.Root>
        )}
      </div>
      <div
        style={{
          marginTop: 36,
          fontSize: 30,
          color: theme.muted,
          opacity: interpolate(frame, [10, 30], [0, 1], {
            extrapolateLeft: "clamp",
            extrapolateRight: "clamp",
          }),
        }}
      >
        A real .pptx, rendered live by the library
      </div>
    </AbsoluteFill>
  );
}
