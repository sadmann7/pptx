import { Composition } from "remotion";

import { calculateLaunchMetadata, DEFAULT_LAUNCH_PROPS, Launch } from "@/components/launch";
import { LAUNCH_DURATION, VIDEO_H, VIDEO_W } from "@/lib/constants";
import "@/styles/globals.css";

export function RemotionRoot() {
  return (
    <Composition
      id="launch"
      component={Launch}
      durationInFrames={LAUNCH_DURATION}
      width={VIDEO_W}
      height={VIDEO_H}
      fps={30}
      defaultProps={DEFAULT_LAUNCH_PROPS}
      calculateMetadata={calculateLaunchMetadata}
    />
  );
}
