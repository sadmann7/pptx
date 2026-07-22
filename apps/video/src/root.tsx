import { Composition } from "remotion";

import { Demo, DEMO_DURATION_IN_FRAMES } from "./demo";
import "./index.css";
import { LaunchFilm } from "./launch-film";

const video = { width: 1920, height: 1080, fps: 30 } as const;

export function RemotionRoot() {
  return (
    <>
      <Composition
        id="demo"
        component={Demo}
        durationInFrames={DEMO_DURATION_IN_FRAMES}
        {...video}
      />
      <Composition id="launch" component={LaunchFilm} durationInFrames={480} {...video} />
    </>
  );
}
