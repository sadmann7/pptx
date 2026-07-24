import { Composition } from "remotion";

import { LAUNCH_DURATION, LaunchFilm } from "@/compositions/launch-film";
import "@/styles/globals.css";

export function RemotionRoot() {
  return (
    <Composition
      id="launch"
      component={LaunchFilm}
      durationInFrames={LAUNCH_DURATION}
      width={1920}
      height={1080}
      fps={30}
    />
  );
}
