import { AbsoluteFill, OffthreadVideo, Sequence, staticFile } from "remotion";

import { Backdrop, SceneContent } from "@/components/backdrop";
import {
  CONTENT_LEAD,
  DEMO_CONTENT_FRAMES,
  DEMO_DURATION,
  EDITOR_DEMO_FILE,
} from "@/lib/constants";

export function DemoScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={DEMO_DURATION}>
        <Sequence from={CONTENT_LEAD} durationInFrames={DEMO_CONTENT_FRAMES} layout="none">
          <OffthreadVideo
            src={staticFile(EDITOR_DEMO_FILE)}
            style={{ width: "100%", height: "100%", objectFit: "cover" }}
          />
        </Sequence>
      </SceneContent>
    </AbsoluteFill>
  );
}
