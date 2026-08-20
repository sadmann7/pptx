import { Typewriter } from "@pptx/ui/components/remocn/typewriter";
import { AbsoluteFill, Sequence } from "remotion";

import { Backdrop, SceneContent } from "@/components/backdrop";
import { CONTENT_LEAD, CTA_DURATION } from "@/lib/constants";
import { color } from "@/lib/theme";

export function CtaScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={CTA_DURATION} fadeOut={false}>
        <Sequence from={CONTENT_LEAD} layout="none">
          <Typewriter
            text="npm i @diceui/pptx"
            fontSize={80}
            fontWeight={700}
            color={color.white}
            charsPerSecond={16}
          />
        </Sequence>
      </SceneContent>
    </AbsoluteFill>
  );
}
