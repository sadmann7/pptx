import type {
  TransitionPresentation,
  TransitionPresentationComponentProps,
} from "@remotion/transitions";
import { linearTiming, TransitionSeries } from "@remotion/transitions";
import { fade } from "@remotion/transitions/fade";
import { AbsoluteFill, CalculateMetadataFunction, staticFile } from "remotion";

import { CompositionScene } from "@/components/scenes/composition-scene";
import { CtaScene } from "@/components/scenes/cta-scene";
import { DemoScene } from "@/components/scenes/demo-scene";
import { FeaturesScene } from "@/components/scenes/features-scene";
import { ShowcaseScene } from "@/components/scenes/showcase-scene";
import { TitleScene } from "@/components/scenes/title-scene";
import {
  COMPOSITION_DURATION,
  CTA_DURATION,
  DEMO_DURATION,
  EDITOR_DEMO_FILE,
  FEATURES_DURATION,
  HANDOFF_FRAMES,
  launchDuration,
  SPOTLIGHTS_TOTAL,
  TITLE_DURATION,
} from "@/lib/constants";
import { fontVars, geistSans } from "@/lib/fonts";
import { FADE_DURATION } from "@/lib/timing";

const fadeTiming = linearTiming({ durationInFrames: FADE_DURATION });
const handoffTiming = linearTiming({ durationInFrames: HANDOFF_FRAMES });

export const calculateLaunchMetadata: CalculateMetadataFunction<LaunchProps> = async () => {
  let hasEditorDemo = false;
  try {
    const response = await fetch(staticFile(EDITOR_DEMO_FILE), { method: "HEAD" });
    hasEditorDemo = response.ok;
  } catch {
    hasEditorDemo = false;
  }

  return {
    props: { hasEditorDemo },
    durationInFrames: launchDuration(hasEditorDemo),
  };
};

function persistPresentation(): TransitionPresentation<Record<string, never>> {
  return { component: PersistPresentation, props: {} };
}

function PersistPresentation({
  children,
  presentationDirection,
}: TransitionPresentationComponentProps<Record<string, never>>) {
  return (
    <AbsoluteFill style={{ zIndex: presentationDirection === "exiting" ? 1 : 0 }}>
      {children}
    </AbsoluteFill>
  );
}

interface LaunchProps extends Record<string, unknown> {
  hasEditorDemo: boolean;
}

export const DEFAULT_LAUNCH_PROPS: LaunchProps = {
  hasEditorDemo: false,
};

export function Launch({ hasEditorDemo = false }: LaunchProps) {
  const scenes = [
    <TransitionSeries.Sequence key="title" durationInFrames={TITLE_DURATION}>
      <TitleScene />
    </TransitionSeries.Sequence>,
    <TransitionSeries.Sequence key="showcase" durationInFrames={SPOTLIGHTS_TOTAL}>
      <ShowcaseScene />
    </TransitionSeries.Sequence>,
    ...(hasEditorDemo
      ? [
          <TransitionSeries.Sequence key="demo" durationInFrames={DEMO_DURATION}>
            <DemoScene />
          </TransitionSeries.Sequence>,
        ]
      : []),
    <TransitionSeries.Sequence key="composition" durationInFrames={COMPOSITION_DURATION}>
      <CompositionScene />
    </TransitionSeries.Sequence>,
    <TransitionSeries.Sequence key="features" durationInFrames={FEATURES_DURATION}>
      <FeaturesScene />
    </TransitionSeries.Sequence>,
    <TransitionSeries.Sequence key="cta" durationInFrames={CTA_DURATION}>
      <CtaScene />
    </TransitionSeries.Sequence>,
  ];

  return (
    <AbsoluteFill style={{ ...fontVars, fontFamily: geistSans }}>
      <TransitionSeries>
        {scenes.flatMap((scene, index) => {
          if (index === 0) return [scene];
          const from = scenes[index - 1]?.key;
          const handoff = from === "showcase" && scene.key === "composition";
          return [
            handoff ? (
              <TransitionSeries.Transition
                key={`cut-${index}`}
                presentation={persistPresentation()}
                timing={handoffTiming}
              />
            ) : (
              <TransitionSeries.Transition
                key={`cut-${index}`}
                presentation={fade()}
                timing={fadeTiming}
              />
            ),
            scene,
          ];
        })}
      </TransitionSeries>
    </AbsoluteFill>
  );
}
