import type { CSSProperties } from "react";

import { BlurOutUp } from "@pptx/ui/components/remocn/blur-out-up";
import { FocusBlurResolve } from "@pptx/ui/components/remocn/focus-blur-resolve";
import { SoftBlurIn } from "@pptx/ui/components/remocn/soft-blur-in";
import { Typewriter } from "@pptx/ui/components/remocn/typewriter";
import { AbsoluteFill, Easing, interpolate, Sequence, useCurrentFrame } from "remotion";

import { PptxCard } from "@/components/pptx-card";
import { geistSans } from "@/lib/fonts";

const C = {
  ink: "#0e1117",
  line: "#2a3040",
  accent: "#a78bfa",
  white: "#f5f6f8",
  muted: "#8b92a8",
};

const clamp = {
  extrapolateLeft: "clamp",
  extrapolateRight: "clamp",
} as const;
const ease = Easing.bezier(0.22, 1, 0.36, 1);

const progress = (frame: number, from: number, to: number) =>
  interpolate(frame, [from, to], [0, 1], { ...clamp, easing: ease });

const fadeWindow = (frame: number, duration: number, edge = 16) =>
  interpolate(frame, [0, edge, duration - edge, duration], [0, 1, 1, 0], clamp);

const font: CSSProperties = {
  fontFamily: geistSans,
};

function Noise() {
  return (
    <AbsoluteFill
      style={{
        opacity: 0.045,
        backgroundImage:
          'url("data:image/svg+xml,%3Csvg viewBox=%270 0 180 180%27 xmlns=%27http://www.w3.org/2000/svg%27%3E%3Cfilter id=%27n%27%3E%3CfeTurbulence type=%27fractalNoise%27 baseFrequency=%271.1%27 numOctaves=%273%27 stitchTiles=%27stitch%27/%3E%3C/filter%3E%3Crect width=%27100%25%27 height=%27100%25%27 filter=%27url(%23n)%27 opacity=%27.8%27/%3E%3C/svg%3E")',
        pointerEvents: "none",
      }}
    />
  );
}

function Backdrop() {
  return (
    <AbsoluteFill
      style={{
        ...font,
        color: C.white,
        overflow: "hidden",
        background: [
          "radial-gradient(ellipse 80% 60% at 20% 80%, rgba(124,58,237,.28), transparent)",
          "radial-gradient(ellipse 70% 50% at 80% 20%, rgba(59,130,246,.25), transparent)",
          "radial-gradient(ellipse 60% 40% at 50% 50%, rgba(236,72,153,.15), transparent)",
          C.ink,
        ].join(", "),
      }}
    >
      <AbsoluteFill
        style={{
          opacity: 0.12,
          backgroundImage:
            "linear-gradient(rgba(255,255,255,.06) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,.06) 1px, transparent 1px)",
          backgroundSize: "72px 72px",
          maskImage: "linear-gradient(to bottom, transparent, black 20%, black 80%, transparent)",
        }}
      />
      <Noise />
    </AbsoluteFill>
  );
}

const SLIDE_W = 960;
const SLIDE_H = 540;

interface Spotlight {
  file: string;
  slideIndex: number;
  caption: string;
}

const SPOTLIGHTS: Spotlight[] = [
  {
    file: "editorial-forest.pptx",
    slideIndex: 1,
    caption: "Pixel-perfect rendering",
  },
  {
    file: "sakura-chroma.pptx",
    slideIndex: 0,
    caption: "Gradients, shapes & effects",
  },
  {
    file: "bold-poster.pptx",
    slideIndex: 0,
    caption: "Typography that holds up",
  },
  {
    file: "emerald-editorial.pptx",
    slideIndex: 1,
    caption: "Complex layouts, intact",
  },
  {
    file: "biennale-yellow.pptx",
    slideIndex: 0,
    caption: "Any theme, any style",
  },
];

function SpotlightScene({ spotlight, duration }: { spotlight: Spotlight; duration: number }) {
  const frame = useCurrentFrame();
  const t = progress(frame, 0, duration - 1);
  const intro = progress(frame, 0, 26);
  const opacity = fadeWindow(frame, duration, 12);

  const cardW = 1360;
  const cardH = (cardW / SLIDE_W) * SLIDE_H;

  return (
    <AbsoluteFill>
      <Backdrop />
      <AbsoluteFill style={{ opacity }}>
        <div
          style={{
            position: "absolute",
            left: (1920 - cardW) / 2,
            top: (1080 - cardH) / 2 - 24,
            width: cardW,
            height: cardH,
            borderRadius: 12,
            overflow: "hidden",
            boxShadow: "0 60px 130px rgba(0,0,0,.55)",
            transformStyle: "preserve-3d",
            transform: `perspective(1900px) translate3d(${interpolate(t, [0, 1], [26, -26])}px, ${interpolate(t, [0, 1], [14, -10])}px, 0) rotateX(${interpolate(t, [0, 1], [1.8, -0.8])}deg) rotateY(${interpolate(t, [0, 1], [-4, 2])}deg) scale(${interpolate(intro, [0, 1], [0.95, 1])})`,
          }}
        >
          <div
            style={{
              width: SLIDE_W,
              height: SLIDE_H,
              transform: `scale(${cardW / SLIDE_W})`,
              transformOrigin: "top left",
            }}
          >
            <PptxCard
              file={spotlight.file}
              width={SLIDE_W}
              height={SLIDE_H}
              slideIndex={spotlight.slideIndex}
            />
          </div>
        </div>

        <div style={{ position: "absolute", left: 0, right: 0, bottom: 44, height: 100 }}>
          <Sequence from={10} durationInFrames={duration - 10} layout="none">
            <BlurOutUp text={spotlight.caption} fontSize={52} fontWeight={800} color={C.white} />
          </Sequence>
        </div>
      </AbsoluteFill>
    </AbsoluteFill>
  );
}

function TitleCard({ duration }: { duration: number }) {
  const frame = useCurrentFrame();
  const opacity = fadeWindow(frame, duration, 14);

  return (
    <AbsoluteFill>
      <Backdrop />
      <AbsoluteFill
        style={{
          ...font,
          alignItems: "center",
          justifyContent: "center",
          opacity,
        }}
      >
        <SoftBlurIn
          text="PowerPoint in the Browser"
          fontSize={96}
          fontWeight={800}
          color={C.white}
        />
      </AbsoluteFill>
    </AbsoluteFill>
  );
}

const SNAPS = [
  "Parse. Render. Edit. Re-export.",
  "Tables, charts, shapes, images.",
  "TypeScript-first. Zero runtime.",
];

const SNAP_BEAT = 55;
const FEATURES_DURATION = SNAPS.length * SNAP_BEAT;

function FeaturesScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      {SNAPS.map((snap, i) => (
        <Sequence key={snap} from={i * SNAP_BEAT} durationInFrames={SNAP_BEAT} layout="none">
          <FocusBlurResolve text={snap} fontSize={72} fontWeight={600} color={C.white} />
        </Sequence>
      ))}
    </AbsoluteFill>
  );
}

const CTA_DURATION = 90;

function CtaScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <Typewriter
        text="npm i @diceui/pptx"
        fontSize={80}
        fontWeight={700}
        color={C.white}
        charsPerSecond={16}
      />
    </AbsoluteFill>
  );
}

const SPOTLIGHT_DURATION = 85;
const OVERLAP = 8;

export function Launch() {
  const spotlightStart = 75 - OVERLAP;
  const featuresStart = spotlightStart + SPOTLIGHTS.length * (SPOTLIGHT_DURATION - OVERLAP);
  const ctaStart = featuresStart + FEATURES_DURATION - OVERLAP;

  return (
    <AbsoluteFill style={{ background: C.ink }}>
      <Sequence from={0} durationInFrames={75} premountFor={20}>
        <TitleCard duration={75} />
      </Sequence>
      {SPOTLIGHTS.map((spotlight, index) => (
        <Sequence
          key={spotlight.caption}
          from={spotlightStart + index * (SPOTLIGHT_DURATION - OVERLAP)}
          durationInFrames={SPOTLIGHT_DURATION}
          premountFor={30}
        >
          <SpotlightScene spotlight={spotlight} duration={SPOTLIGHT_DURATION} />
        </Sequence>
      ))}
      <Sequence from={featuresStart} durationInFrames={FEATURES_DURATION} premountFor={20}>
        <FeaturesScene />
      </Sequence>
      <Sequence from={ctaStart} durationInFrames={CTA_DURATION} premountFor={20}>
        <CtaScene />
      </Sequence>
    </AbsoluteFill>
  );
}

export const LAUNCH_DURATION =
  75 -
  OVERLAP +
  SPOTLIGHTS.length * (SPOTLIGHT_DURATION - OVERLAP) +
  FEATURES_DURATION -
  OVERLAP +
  CTA_DURATION;
