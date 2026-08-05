import { Config } from "@remotion/cli/config";
import { enableTailwind } from "@remotion/tailwind-v4";
import path from "node:path";

Config.setVideoImageFormat("jpeg");
Config.setOverwriteOutput(true);
/**
 * Nothing here produces audio, but the track is kept because platforms X among
 * them can mishandle a file with no audio stream at all. At the default bitrate
 * that silence was 317 kb/s, around a tenth of the file, so it is encoded as
 * cheaply as the muxer allows instead of being dropped.
 */
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
