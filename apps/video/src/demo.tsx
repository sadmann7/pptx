import type { CSSProperties } from "react";

import { linearTiming, TransitionSeries } from "@remotion/transitions";
import { fade } from "@remotion/transitions/fade";
import { AbsoluteFill } from "remotion";

import { whipPan } from "./components/whip-pan";
import { geistMono, geistSans } from "./fonts";
import { CodeScene } from "./scenes/code";
import { FeaturesScene } from "./scenes/features";
import { INTERACTION_FRAMES, InteractionScene } from "./scenes/interactions";
import { IntroScene } from "./scenes/intro";
import { OutroScene } from "./scenes/outro";
import { SHOWCASE_FRAMES, ShowcaseScene } from "./scenes/showcase";

export const FPS = 30;

const INTRO = 70;
const SHOWCASE = SHOWCASE_FRAMES;
const INTERACTIONS = INTERACTION_FRAMES;
const CODE = 140;
const FEATURES = 100;
const OUTRO = 80;

const T_WHIP = 20;
const T_FADE = 14;

export const DEMO_DURATION_IN_FRAMES =
  INTRO + SHOWCASE + INTERACTIONS + CODE + FEATURES + OUTRO - T_WHIP - T_FADE * 4;

const FONT_VARS = {
  "--font-geist-sans": geistSans,
  "--font-geist-mono": geistMono,
} as CSSProperties;

export function Demo() {
  return (
    <AbsoluteFill style={FONT_VARS}>
      <TransitionSeries>
        <TransitionSeries.Sequence name="Intro" durationInFrames={INTRO}>
          <IntroScene />
        </TransitionSeries.Sequence>

        <TransitionSeries.Transition
          presentation={whipPan({ direction: "left" })}
          timing={linearTiming({ durationInFrames: T_WHIP })}
        />

        <TransitionSeries.Sequence name="Showcase" durationInFrames={SHOWCASE}>
          <ShowcaseScene />
        </TransitionSeries.Sequence>

        <TransitionSeries.Transition
          presentation={fade()}
          timing={linearTiming({ durationInFrames: T_FADE })}
        />

        <TransitionSeries.Sequence name="Interactions" durationInFrames={INTERACTIONS}>
          <InteractionScene />
        </TransitionSeries.Sequence>

        <TransitionSeries.Transition
          presentation={fade()}
          timing={linearTiming({ durationInFrames: T_FADE })}
        />

        <TransitionSeries.Sequence name="Code" durationInFrames={CODE}>
          <CodeScene />
        </TransitionSeries.Sequence>

        <TransitionSeries.Transition
          presentation={fade()}
          timing={linearTiming({ durationInFrames: T_FADE })}
        />

        <TransitionSeries.Sequence name="Features" durationInFrames={FEATURES}>
          <FeaturesScene />
        </TransitionSeries.Sequence>

        <TransitionSeries.Transition
          presentation={fade()}
          timing={linearTiming({ durationInFrames: T_FADE })}
        />

        <TransitionSeries.Sequence name="Outro" durationInFrames={OUTRO}>
          <OutroScene />
        </TransitionSeries.Sequence>
      </TransitionSeries>
    </AbsoluteFill>
  );
}
