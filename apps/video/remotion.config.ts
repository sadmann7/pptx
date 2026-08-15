import { Config } from "@remotion/cli/config";
import { enableTailwind } from "@remotion/tailwind-v4";
import path from "node:path";

Config.setVideoImageFormat("jpeg");
Config.setOverwriteOutput(true);
// Keep a cheap silent track; some platforms drop videos with no audio stream.
Config.setAudioBitrate("32K");
Config.overrideWebpackConfig((config) => {
  const withTailwind = enableTailwind(config);

  return {
    ...withTailwind,
    resolve: {
      ...withTailwind.resolve,
      alias: {
        ...withTailwind.resolve?.alias,
        "@": path.resolve(process.cwd(), "src"),
      },
    },
  };
});
