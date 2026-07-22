import type { CSSProperties } from "react";

import { linearTiming, TransitionSeries } from "@remotion/transitions";
import { fade } from "@remotion/transitions/fade";
import { AbsoluteFill } from "remotion";

import { geistMono, geistSans } from "./fonts";
import { CANVAS_FLYTHROUGH_FRAMES, CanvasFlythroughScene } from "./scenes/canvas-flythrough";
import { CodeScene } from "./scenes/code";
import { FeaturesScene } from "./scenes/features";
import { INTERACTION_FRAMES, InteractionScene } from "./scenes/interactions";
import { IntroScene } from "./scenes/intro";
import { OutroScene } from "./scenes/outro";
import { PERSPECTIVE_DRIFT_FRAMES, PerspectiveDriftScene } from "./scenes/perspective-drift";

export const FPS = 30;

const INTRO = 70;
const PERSPECTIVE = PERSPECTIVE_DRIFT_FRAMES;
const FLYTHROUGH = CANVAS_FLYTHROUGH_FRAMES;
const INTERACTIONS = INTERACTION_FRAMES;
const CODE = 140;
const FEATURES = 100;
const OUTRO = 80;

const T_FADE = 14;
const SCENE_COUNT = 7;

export const DEMO_DURATION_IN_FRAMES =
  INTRO +
  PERSPECTIVE +
  FLYTHROUGH +
  INTERACTIONS +
  CODE +
  FEATURES +
  OUTRO -
  T_FADE * (SCENE_COUNT - 1);

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
          presentation={fade()}
          timing={linearTiming({ durationInFrames: T_FADE })}
        />

        <TransitionSeries.Sequence name="PerspectiveDrift" durationInFrames={PERSPECTIVE}>
          <PerspectiveDriftScene />
        </TransitionSeries.Sequence>

        <TransitionSeries.Transition
          presentation={fade()}
          timing={linearTiming({ durationInFrames: T_FADE })}
        />

        <TransitionSeries.Sequence name="CanvasFlythrough" durationInFrames={FLYTHROUGH}>
          <CanvasFlythroughScene />
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
