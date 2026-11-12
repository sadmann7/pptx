import { Config } from "@remotion/cli/config";
import { enableTailwind } from "@remotion/tailwind-v4";
import path from "node:path";

// PowerPoint and Chromium dither long gradient backgrounds, and the default
// JPEG quality of 80 quantizes that dither into visible bands.
Config.setVideoImageFormat("jpeg");
Config.setJpegQuality(100);
Config.setOverwriteOutput(true);
// The showcase's 3D transform hands slides to Chromium's compositor, whose
// default software backend draws gradients without dithering; ANGLE dithers.
Config.setChromiumOpenGlRenderer("angle");
// At the default preset, h264 flattens the near-static dark ramps into bands,
// and a finer quantizer costs twice the file for the same result.
Config.setX264Preset("slower");
// Remotion muxes a silent AAC track into the mp4 at ~317 kbps by default.
Config.setAudioBitrate("32K");
// The default uses 8 of 24 threads, and frame capture scales with more.
Config.setConcurrency("75%");
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
