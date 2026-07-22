import { AbsoluteFill } from "remotion";

const INK = "#070806";

/**
 * Shared background for all scenes.
 * Dark radial glow + subtle grid overlay + film-grain noise.
 * Matches the GPT launch-kit aesthetic.
 */
export function SceneBg() {
  return (
    <AbsoluteFill
      style={{
        overflow: "hidden",
        background: `radial-gradient(circle at 73% 26%, rgba(59,130,246,.10), transparent 29%), ${INK}`,
      }}
    >
      {/* Grid overlay */}
      <AbsoluteFill
        style={{
          opacity: 0.14,
          backgroundImage:
            "linear-gradient(rgba(59,130,246,.25) 1px, transparent 1px), linear-gradient(90deg, rgba(59,130,246,.25) 1px, transparent 1px)",
          backgroundSize: "72px 72px",
          maskImage: "linear-gradient(to bottom, transparent, black 20%, black 80%, transparent)",
        }}
      />
      {/* Film grain noise */}
      <AbsoluteFill
        style={{
          opacity: 0.04,
          backgroundImage:
            'url("data:image/svg+xml,%3Csvg viewBox=%270 0 180 180%27 xmlns=%27http://www.w3.org/2000/svg%27%3E%3Cfilter id=%27n%27%3E%3CfeTurbulence type=%27fractalNoise%27 baseFrequency=%271.1%27 numOctaves=%273%27 stitchTiles=%27stitch%27/%3E%3C/filter%3E%3Crect width=%27100%25%27 height=%27100%25%27 filter=%27url(%23n)%27 opacity=%27.8%27/%3E%3C/svg%3E")',
          pointerEvents: "none",
        }}
      />
    </AbsoluteFill>
  );
}
