import { Composition } from "remotion";

import { LAUNCH_DURATION, Launch } from "@/components/launch";
import "@/styles/globals.css";

export function RemotionRoot() {
  return (
    <Composition
      id="launch"
      component={Launch}
      durationInFrames={LAUNCH_DURATION}
      width={1920}
      height={1080}
      fps={30}
    />
  );
}
