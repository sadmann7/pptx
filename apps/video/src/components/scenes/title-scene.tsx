import { SoftBlurIn } from "@pptx/ui/components/remocn/soft-blur-in";
import { AbsoluteFill } from "remotion";

import { Backdrop, SceneContent } from "@/components/backdrop";
import { TITLE_DURATION } from "@/lib/constants";
import { color } from "@/lib/theme";

export function TitleScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={TITLE_DURATION}>
        <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
          <SoftBlurIn
            text="PowerPoint in the browser."
            fontSize={104}
            fontWeight={800}
            color={color.white}
          />
        </AbsoluteFill>
      </SceneContent>
    </AbsoluteFill>
  );
}
