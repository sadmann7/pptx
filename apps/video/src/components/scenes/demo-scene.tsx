import {
  AbsoluteFill,
  Freeze,
  OffthreadVideo,
  Sequence,
  staticFile,
  useCurrentFrame,
} from "remotion";

import { Backdrop, SceneContent } from "@/components/backdrop";
import {
  DEMO_CONTENT_FRAMES,
  DEMO_DURATION,
  DEMO_PUSH_START,
  EDITOR_DEMO_FILE,
  HANDOFF_FRAMES,
} from "@/lib/constants";
import { demoSlideRect, previewSlideRect, rectCenter } from "@/lib/layout";
import { PANEL_RADIUS, panelShadow } from "@/lib/theme";
import { progress } from "@/lib/timing";

function DemoVideo() {
  return (
    <OffthreadVideo
      src={staticFile(EDITOR_DEMO_FILE)}
      style={{ width: "100%", height: "100%", objectFit: "cover" }}
    />
  );
}

function DemoPushIn() {
  const time = useCurrentFrame();
  const push = progress(time, DEMO_PUSH_START, HANDOFF_FRAMES - 2);
  const settle = 1 - push;
  const from = previewSlideRect(1);
  const to = demoSlideRect();
  const start = from.zoom / to.zoom;
  const scale = start + (1 - start) * push;
  const origin = rectCenter(to);
  const target = rectCenter(from);

  return (
    <AbsoluteFill
      style={{
        opacity: progress(time, 0, 8),
        overflow: "hidden",
        borderRadius: (PANEL_RADIUS / scale) * settle,
        boxShadow: panelShadow(settle),
        transformOrigin: `${origin.x}px ${origin.y}px`,
        transform: [
          `translate(${(target.x - origin.x) * settle}px, ${(target.y - origin.y) * settle}px)`,
          `scale(${scale})`,
        ].join(" "),
      }}
    >
      <Freeze frame={0}>
        <DemoVideo />
      </Freeze>
    </AbsoluteFill>
  );
}

export function DemoScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={DEMO_DURATION}>
        <Sequence durationInFrames={HANDOFF_FRAMES} layout="none">
          <DemoPushIn />
        </Sequence>
        <Sequence from={HANDOFF_FRAMES} durationInFrames={DEMO_CONTENT_FRAMES} layout="none">
          <DemoVideo />
        </Sequence>
      </SceneContent>
    </AbsoluteFill>
  );
}
