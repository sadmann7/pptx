import { AbsoluteFill, Sequence } from "remotion";

import { CodeScene } from "./scenes/code";
import { FeaturesScene } from "./scenes/features";
import { IntroScene } from "./scenes/intro";
import { OutroScene } from "./scenes/outro";
import { ShowcaseScene } from "./scenes/showcase";
import { theme } from "./theme";

export const FPS = 30;

const INTRO = { from: 0, duration: 90 };
const SHOWCASE = { from: 80, duration: 250 };
const CODE = { from: 320, duration: 160 };
const FEATURES = { from: 470, duration: 160 };
const OUTRO = { from: 620, duration: 100 };

export const DEMO_DURATION_IN_FRAMES = OUTRO.from + OUTRO.duration;

export function Demo() {
  return (
    <AbsoluteFill style={{ backgroundColor: theme.background, fontFamily: theme.fontSans }}>
      <Sequence name="Intro" from={INTRO.from} durationInFrames={INTRO.duration}>
        <IntroScene />
      </Sequence>
      <Sequence name="Showcase" from={SHOWCASE.from} durationInFrames={SHOWCASE.duration}>
        <ShowcaseScene />
      </Sequence>
      <Sequence name="Code" from={CODE.from} durationInFrames={CODE.duration}>
        <CodeScene />
      </Sequence>
      <Sequence name="Features" from={FEATURES.from} durationInFrames={FEATURES.duration}>
        <FeaturesScene />
      </Sequence>
      <Sequence name="Outro" from={OUTRO.from} durationInFrames={OUTRO.duration}>
        <OutroScene />
      </Sequence>
    </AbsoluteFill>
  );
}
