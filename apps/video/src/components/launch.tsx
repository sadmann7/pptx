import type { CSSProperties } from "react";

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
  description: string;
}

const SPOTLIGHTS: Spotlight[] = [
  {
    file: "editorial-forest.pptx",
    slideIndex: 1,
    caption: "Pixel-perfect rendering",
    description: "Every shape, gradient, and layout lands exactly where PowerPoint put it.",
  },
  {
    file: "sakura-chroma.pptx",
    slideIndex: 0,
    caption: "Gradients, shapes & effects",
    description: "Linear fills, radial blends, shadows, and transparency. All preserved.",
  },
  {
    file: "bold-poster.pptx",
    slideIndex: 0,
    caption: "Typography that holds up",
    description: "Font weights, spacing, and text boxes render faithfully in the browser.",
  },
  {
    file: "emerald-editorial.pptx",
    slideIndex: 1,
    caption: "Complex layouts, intact",
    description: "Grouped shapes, nested containers, and multi-column slides stay intact.",
  },
];

function SpotlightScene({
  spotlight,
  duration,
  index,
}: {
  spotlight: Spotlight;
  duration: number;
  index: number;
}) {
  const frame = useCurrentFrame();
  const intro = progress(frame, 0, 28);
  const captionIntro = progress(frame, 10, 38);
  const descIntro = progress(frame, 18, 46);
  const opacity = fadeWindow(frame, duration, 14);

  const isLeft = index % 2 === 0;

  const cardW = 1050;
  const cardH = (cardW / SLIDE_W) * SLIDE_H;
  const cardMargin = 80;
  const cardX = isLeft ? cardMargin : 1920 - cardW - cardMargin;
  const cardSlideX = interpolate(intro, [0, 1], [isLeft ? -30 : 30, 0]);
  const cardSlideY = interpolate(intro, [0, 1], [16, 0]);
  const cardScale = interpolate(intro, [0, 1], [0.97, 1]);

  const captionX = isLeft ? cardMargin + cardW + 60 : 0;
  const captionW = 1920 - cardW - cardMargin - 60;
  const captionSlide = interpolate(captionIntro, [0, 1], [20, 0]);
  const descSlide = interpolate(descIntro, [0, 1], [14, 0]);

  return (
    <AbsoluteFill>
      <Backdrop />
      <AbsoluteFill style={{ opacity }}>
        <div
          style={{
            position: "absolute",
            left: cardX,
            top: (1080 - cardH) / 2,
            width: cardW,
            height: cardH,
            borderRadius: 12,
            overflow: "hidden",
            boxShadow: "0 40px 100px rgba(0,0,0,.45)",
            opacity: intro,
            transform: `translate(${cardSlideX}px, ${cardSlideY}px) scale(${cardScale})`,
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

        <div
          style={{
            position: "absolute",
            left: captionX,
            top: 0,
            width: captionW,
            height: 1080,
            display: "flex",
            flexDirection: "column",
            justifyContent: "center",
            padding: "0 40px",
            textAlign: isLeft ? "left" : "right",
          }}
        >
          <span
            style={{
              ...font,
              fontSize: 52,
              fontWeight: 800,
              color: C.white,
              letterSpacing: -1.2,
              lineHeight: 1.15,
              opacity: captionIntro,
              transform: `translateY(${captionSlide}px)`,
            }}
          >
            {spotlight.caption}
          </span>
          <span
            style={{
              ...font,
              fontSize: 22,
              fontWeight: 400,
              color: C.muted,
              lineHeight: 1.5,
              marginTop: 16,
              opacity: descIntro,
              transform: `translateY(${descSlide}px)`,
            }}
          >
            {spotlight.description}
          </span>
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
          <SpotlightScene spotlight={spotlight} duration={SPOTLIGHT_DURATION} index={index} />
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
