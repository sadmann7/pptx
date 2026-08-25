import { Config } from "@remotion/cli/config";
import { enableTailwind } from "@remotion/tailwind-v4";
import path from "node:path";

// Slide backgrounds are long, low-contrast gradients that PowerPoint dithers,
// and so does Chromium, by alternating between two adjacent levels per pixel.
// The default quality of 80 quantizes that dither away and leaves the ramp
// visibly banded; 100 keeps it, and stays far cheaper to write than png frames.
Config.setVideoImageFormat("jpeg");
Config.setJpegQuality(100);
Config.setOverwriteOutput(true);
// The showcase board sits under a 3D transform, which hands its slides to
// Chromium's compositor. The default software backend draws gradients there
// without dithering, so those same ramps quantize into 4px bands; ANGLE dithers.
Config.setChromiumOpenGlRenderer("angle");
// h264 spends its bits on the moving foreground and flattens the near-static
// dark ramps into bands. Buying the ramps a finer quantizer costs twice the file
// for the same result as simply letting x264 search harder at the default crf.
Config.setX264Preset("slower");
// Remotion muxes a silent AAC track into the mp4 even though nothing here plays
// audio, and it defaults to ~317 kbps of silence. This keeps that track cheap.
Config.setAudioBitrate("32K");
// Defaults to 8 of the 24 threads this renders on; frame capture is the bulk of
// the wall clock and scales with it.
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
