import { Composition, staticFile, type CalculateMetadataFunction } from "remotion";

import {
  DEFAULT_LAUNCH_PROPS,
  EDITOR_DEMO_FILE,
  LAUNCH_DURATION,
  Launch,
  launchDuration,
  type LaunchProps,
} from "@/components/launch";
import "@/styles/globals.css";

export const calculateLaunchMetadata: CalculateMetadataFunction<LaunchProps> = async () => {
  let hasEditorDemo = false;
  try {
    const response = await fetch(staticFile(EDITOR_DEMO_FILE), { method: "HEAD" });
    hasEditorDemo = response.ok;
  } catch {
    hasEditorDemo = false;
  }

  return {
    props: { hasEditorDemo },
    durationInFrames: launchDuration(hasEditorDemo),
  };
};

export function RemotionRoot() {
  return (
    <Composition
      id="launch"
      component={Launch}
      durationInFrames={LAUNCH_DURATION}
      width={1920}
      height={1080}
      fps={30}
      defaultProps={DEFAULT_LAUNCH_PROPS}
      calculateMetadata={calculateLaunchMetadata}
    />
  );
}
