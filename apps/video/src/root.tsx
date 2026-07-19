import "./index.css";
import { Composition } from "remotion";

import { Demo, DEMO_DURATION_IN_FRAMES, FPS } from "./demo";

export function RemotionRoot() {
  return (
    <Composition
      id="demo"
      component={Demo}
      durationInFrames={DEMO_DURATION_IN_FRAMES}
      fps={FPS}
      width={1920}
      height={1080}
    />
  );
}
