import type { CSSProperties, ReactNode } from "react";

import { FocusBlurResolve } from "@pptx/ui/components/remocn/focus-blur-resolve";
import { SoftBlurIn } from "@pptx/ui/components/remocn/soft-blur-in";
import { Typewriter } from "@pptx/ui/components/remocn/typewriter";
import { linearTiming, TransitionSeries } from "@remotion/transitions";
import { fade } from "@remotion/transitions/fade";
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

/** Length of the crossfade between sections. */
const FADE_DURATION = 12;
/** How long a section's content takes to clear the crossfade at either edge. */
const CONTENT_FADE = 10;
/**
 * Held at the head of a section so its content starts once the crossfade has
 * finished rather than animating behind it. Every section animates its own
 * entrance, so this waits rather than fading and the gap between sections
 * stays as short as the crossfade itself.
 */
const CONTENT_LEAD = FADE_DURATION;

/**
 * A section is its content plus the lead-in that clears the incoming
 * crossfade and the outgoing crossfade it has to be gone by. Sized this way
 * the content fades out over its own last frames and reaches zero exactly as
 * the crossfade starts, with no stretch of bare backdrop in between.
 */
const sectionDuration = (contentFrames: number) => CONTENT_LEAD + contentFrames + FADE_DURATION;

const progress = (frame: number, from: number, to: number) =>
  interpolate(frame, [from, to], [0, 1], { ...clamp, easing: ease });

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

/**
 * Fades a section's content out before the crossfade into the next section
 * begins. Every section paints the same backdrop, so a crossfade has only the
 * content on top of it to blend: leave content up on both sides and the
 * outgoing section shows through the incoming one rather than dissolving into
 * it, which is what put the features text across the last slide.
 *
 * Only the exit needs handling here. Content is held back on the way in by
 * `CONTENT_LEAD`, which keeps it off screen until the crossfade is over.
 */
function SceneContent({
  durationInFrames,
  fadeOut = true,
  children,
}: {
  durationInFrames: number;
  fadeOut?: boolean;
  children: ReactNode;
}) {
  const frame = useCurrentFrame();
  const exitStart = durationInFrames - FADE_DURATION - CONTENT_FADE;
  const opacity = fadeOut
    ? interpolate(frame, [exitStart, exitStart + CONTENT_FADE], [1, 0], clamp)
    : 1;

  return <AbsoluteFill style={{ opacity }}>{children}</AbsoluteFill>;
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
    caption: "True to the original",
    description: "Slides look exactly the way they do in PowerPoint.",
  },
  {
    file: "sakura-chroma.pptx",
    slideIndex: 0,
    caption: "Rich visuals",
    description: "Gradients, shadows, and transparency come through untouched.",
  },
  {
    file: "make-something-strange.pptx",
    slideIndex: 5,
    caption: "Charts",
    description: "Data visualizations render natively, no images involved.",
  },
  {
    file: "make-something-strange.pptx",
    slideIndex: 6,
    caption: "Tables",
    description: "Styled cells, borders, and layouts stay just as designed.",
  },
  {
    file: "bold-poster.pptx",
    slideIndex: 0,
    caption: "Typography",
    description: "Weights, spacing, and alignment carry over precisely.",
  },
  {
    file: "emerald-editorial.pptx",
    slideIndex: 1,
    caption: "Complex layouts",
    description: "Grouped shapes and multi-column designs hold together.",
  },
];

const SPOTLIGHT_DURATION = 85;

function SpotlightScene({ spotlight, index }: { spotlight: Spotlight; index: number }) {
  const frame = useCurrentFrame();
  const intro = progress(frame, 0, 28);
  const captionIntro = progress(frame, 10, 38);
  const descIntro = progress(frame, 18, 46);

  // Fade content out before the hard cut to the next spotlight so its
  // entrance doesn't overlap with lingering visuals from this one. The cuts
  // between spotlights are the only edges this covers; the crossfades at
  // either end of the run are handled by the enclosing SceneContent.
  const exitFade = interpolate(
    frame,
    [SPOTLIGHT_DURATION - 12, SPOTLIGHT_DURATION - 2],
    [1, 0],
    clamp,
  );
  const contentOpacity = Math.min(intro, exitFade);

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
      <AbsoluteFill style={{ opacity: contentOpacity }}>
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
              fontSize: 64,
              fontWeight: 800,
              color: C.white,
              letterSpacing: -1.6,
              lineHeight: 1.12,
              opacity: captionIntro,
              transform: `translateY(${captionSlide}px)`,
            }}
          >
            {spotlight.caption}
          </span>
          <span
            style={{
              ...font,
              fontSize: 28,
              fontWeight: 400,
              color: C.muted,
              lineHeight: 1.5,
              marginTop: 18,
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

// ── Title ───────────────────────────────────────────────────────────────────

const TITLE_DURATION = 60;

function TitleCard() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={TITLE_DURATION}>
        <AbsoluteFill
          style={{
            ...font,
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <SoftBlurIn
            text="PowerPoint in the browser."
            fontSize={104}
            fontWeight={800}
            color={C.white}
          />
        </AbsoluteFill>
      </SceneContent>
    </AbsoluteFill>
  );
}

// ── Features snaps ──────────────────────────────────────────────────────────

const SNAPS = [
  "Parse. Render. Edit. Re-export.",
  "Tables, charts, shapes, images.",
  "TypeScript-first. Zero runtime.",
];

/**
 * Trimmed from 55 so the run keeps the length it had before the crossfades
 * stopped overlapping it. Each snap used to lose frames to the fade at either
 * end of the section; giving them back in full read as a drag.
 */
const SNAP_BEAT = 47;
const FEATURES_DURATION = sectionDuration(SNAPS.length * SNAP_BEAT);

function FeaturesScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={FEATURES_DURATION}>
        {SNAPS.map((snap, i) => (
          <Sequence
            key={snap}
            from={CONTENT_LEAD + i * SNAP_BEAT}
            durationInFrames={SNAP_BEAT}
            layout="none"
          >
            <FocusBlurResolve text={snap} fontSize={84} fontWeight={600} color={C.white} />
          </Sequence>
        ))}
      </SceneContent>
    </AbsoluteFill>
  );
}

// ── CTA ─────────────────────────────────────────────────────────────────────

const CTA_DURATION = CONTENT_LEAD + 90;

function CtaScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      {/* Closes the video, so there is no crossfade to clear on the way out. */}
      <SceneContent durationInFrames={CTA_DURATION} fadeOut={false}>
        {/* Held past the crossfade so no keystrokes land while it is running. */}
        <Sequence from={CONTENT_LEAD} layout="none">
          <Typewriter
            text="npm i @diceui/pptx"
            fontSize={80}
            fontWeight={700}
            color={C.white}
            charsPerSecond={16}
          />
        </Sequence>
      </SceneContent>
    </AbsoluteFill>
  );
}

// ── Composition ─────────────────────────────────────────────────────────────

const fadeTiming = linearTiming({ durationInFrames: FADE_DURATION });

const SPOTLIGHTS_TOTAL = sectionDuration(SPOTLIGHTS.length * SPOTLIGHT_DURATION);

export function Launch() {
  return (
    <TransitionSeries>
      <TransitionSeries.Sequence durationInFrames={TITLE_DURATION}>
        <TitleCard />
      </TransitionSeries.Sequence>

      <TransitionSeries.Transition presentation={fade()} timing={fadeTiming} />

      {/* All spotlights in one sequence with hard cuts between them, each
          managing its own fade across those cuts. The run shares one backdrop
          so the envelope has only the cards and captions to fade. */}
      <TransitionSeries.Sequence durationInFrames={SPOTLIGHTS_TOTAL}>
        <AbsoluteFill>
          <Backdrop />
          <SceneContent durationInFrames={SPOTLIGHTS_TOTAL}>
            {SPOTLIGHTS.map((spotlight, index) => (
              <Sequence
                key={`${spotlight.file}-${spotlight.slideIndex}`}
                from={CONTENT_LEAD + index * SPOTLIGHT_DURATION}
                durationInFrames={SPOTLIGHT_DURATION}
              >
                <SpotlightScene spotlight={spotlight} index={index} />
              </Sequence>
            ))}
          </SceneContent>
        </AbsoluteFill>
      </TransitionSeries.Sequence>

      <TransitionSeries.Transition presentation={fade()} timing={fadeTiming} />

      <TransitionSeries.Sequence durationInFrames={FEATURES_DURATION}>
        <FeaturesScene />
      </TransitionSeries.Sequence>

      <TransitionSeries.Transition presentation={fade()} timing={fadeTiming} />

      <TransitionSeries.Sequence durationInFrames={CTA_DURATION}>
        <CtaScene />
      </TransitionSeries.Sequence>
    </TransitionSeries>
  );
}

const SECTION_FADES = 3;
export const LAUNCH_DURATION =
  TITLE_DURATION +
  SPOTLIGHTS_TOTAL +
  FEATURES_DURATION +
  CTA_DURATION -
  SECTION_FADES * FADE_DURATION;
