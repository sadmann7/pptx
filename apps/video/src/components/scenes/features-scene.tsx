import { FocusBlurResolve } from "@pptx/ui/components/remocn/focus-blur-resolve";
import { AbsoluteFill, Sequence } from "remotion";

import { Backdrop, SceneContent } from "@/components/backdrop";
import { CONTENT_LEAD, FEATURES_DURATION, SNAP_BEAT, SNAPS } from "@/lib/constants";
import { color } from "@/lib/theme";

export function FeaturesScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={FEATURES_DURATION}>
        {SNAPS.map((snap, i) => (
          <Sequence
            key={snap}
            from={CONTENT_LEAD + i * SNAP_BEAT}
            durationInFrames={SNAP_BEAT}
            layout="none"
          >
            <FocusBlurResolve text={snap} fontSize={84} fontWeight={600} color={color.white} />
          </Sequence>
        ))}
      </SceneContent>
    </AbsoluteFill>
  );
}
