import { AbsoluteFill } from "remotion";

/**
 * Shared background for all scenes.
 * Dark gradient with subtle blue/purple tones to avoid flat-black monotony.
 */
export function SceneBg() {
  return (
    <AbsoluteFill
      style={{
        background:
          "radial-gradient(ellipse 120% 80% at 50% 20%, #1a1a2e 0%, #0f0f1a 40%, #09090b 100%)",
      }}
    />
  );
}
