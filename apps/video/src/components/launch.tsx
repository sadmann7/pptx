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
  /** Centre of the card on the board, in board units (one unit is one slide px). */
  x: number;
  y: number;
  /** How close the camera sits while this card is the subject. */
  zoom: number;
}

/**
 * The cards sit on one board that a camera moves across, rather than each
 * taking a turn in the same spot. Positions step on both axes so every move
 * is diagonal, and the gaps are sized so the neighbouring cards stay just
 * inside the frame edges: that parallax is what carries the run between
 * subjects, in place of the fade to bare backdrop it used to cut through.
 */
const SPOTLIGHTS: Spotlight[] = [
  {
    file: "editorial-forest-editable.pptx",
    slideIndex: 1,
    caption: "True to the original",
    description: "Slides look exactly the way they do in PowerPoint.",
    x: 0,
    y: -120,
    zoom: 1.06,
  },
  {
    file: "pocket-machines-sakura-chroma.pptx",
    slideIndex: 0,
    caption: "Rich visuals",
    description: "Gradients, shadows, and transparency come through untouched.",
    x: 1420,
    y: 280,
    zoom: 1.02,
  },
  {
    file: "adventure-club-pin-and-paper.pptx",
    slideIndex: 3,
    caption: "Charts",
    description: "Data visualizations render natively, no images involved.",
    x: 2820,
    y: -230,
    zoom: 1.08,
  },
  {
    file: "make-something-strange-creative-mode.pptx",
    slideIndex: 6,
    caption: "Tables",
    description: "Styled cells, borders, and layouts stay just as designed.",
    x: 4240,
    y: 240,
    zoom: 1.03,
  },
  {
    file: "after-the-needle-drops-mat.pptx",
    slideIndex: 0,
    caption: "Typography",
    description: "Weights, spacing, and alignment carry over precisely.",
    x: 5640,
    y: -270,
    zoom: 1.07,
  },
  {
    file: "side-quest-club-block-frame.pptx",
    slideIndex: 4,
    caption: "Complex layouts",
    description: "Grouped shapes and multi-column designs hold together.",
    x: 7060,
    y: 160,
    zoom: 1.04,
  },
];

const SPOTLIGHT_DURATION = 85;
/** Frames of a beat spent gliding to the next card. The rest is the hold. */
const CAMERA_MOVE = 30;
const CAMERA_HOLD = SPOTLIGHT_DURATION - CAMERA_MOVE;
const SHOWCASE_FRAMES = SPOTLIGHTS.length * SPOTLIGHT_DURATION;

/** How far ahead of the camera settling a caption starts arriving. */
const CAPTION_LEAD = 8;

/** Keeps the text column readable where cards pass behind it. */
const TEXT_ON_BOARD = "0 2px 14px rgba(14,17,23,.85), 0 0 44px rgba(14,17,23,.75)";

/** Where on screen the subject of the shot sits, leaving the left for text. */
const FOCUS_X = 1290;
const FOCUS_Y = 540;

/** Slow out of the old card, slow into the new one, quick in between. */
const cameraEase = Easing.bezier(0.65, 0, 0.25, 1);

interface Camera {
  x: number;
  y: number;
  zoom: number;
}

/**
 * Where the camera rests on `index`, with `t` running 0 to 1 across the hold.
 * The drift is deliberately small: enough that no frame is frozen, not enough
 * to read as a move of its own.
 */
function restingCamera(index: number, t: number): Camera {
  const spotlight = SPOTLIGHTS[index];
  if (!spotlight) throw new Error(`no spotlight at ${index}`);
  const direction = index % 2 === 0 ? 1 : -1;

  return {
    x: spotlight.x + t * 26 * direction,
    y: spotlight.y + t * 14,
    zoom: spotlight.zoom * (1 + t * 0.014),
  };
}

/** Beat index plus how far the camera has travelled towards the next one. */
function focusAt(frame: number): number {
  const index = Math.min(SPOTLIGHTS.length - 1, Math.floor(frame / SPOTLIGHT_DURATION));
  const local = frame - index * SPOTLIGHT_DURATION;
  if (index === SPOTLIGHTS.length - 1 || local < CAMERA_HOLD) return index;
  return index + cameraEase((local - CAMERA_HOLD) / CAMERA_MOVE);
}

function cameraAt(frame: number): Camera {
  const focus = focusAt(frame);
  const index = Math.floor(focus);
  const travel = focus - index;
  if (travel === 0) {
    // Resting. The last card has no move left to make, so it drifts for the
    // whole beat rather than only the hold.
    const local = frame - index * SPOTLIGHT_DURATION;
    const span = index === SPOTLIGHTS.length - 1 ? SPOTLIGHT_DURATION : CAMERA_HOLD;
    return restingCamera(index, Math.min(1, local / span));
  }

  const from = restingCamera(index, 1);
  const to = restingCamera(index + 1, 0);
  return {
    x: from.x + (to.x - from.x) * travel,
    y: from.y + (to.y - from.y) * travel,
    zoom: from.zoom + (to.zoom - from.zoom) * travel,
  };
}

function Board({ frame }: { frame: number }) {
  const camera = cameraAt(frame);
  const focus = focusAt(frame);

  // A shallow turn that peaks mid-move and is gone by the time the camera
  // settles, so a move reads as the camera swinging round to the next card
  // rather than the board sliding under a fixed lens.
  const swing = Math.sin((focus % 1) * Math.PI) * 1.6;

  return (
    <AbsoluteFill style={{ perspective: 2600 }}>
      <AbsoluteFill
        style={{
          transformOrigin: "0 0",
          transform: [
            `translate(${FOCUS_X}px, ${FOCUS_Y}px)`,
            `rotateY(${-swing}deg)`,
            `scale(${camera.zoom})`,
            `translate(${-camera.x}px, ${-camera.y}px)`,
          ].join(" "),
        }}
      >
        {SPOTLIGHTS.map((spotlight, index) => {
          // Cards recede as the camera leaves them, so the subject is always
          // the one in focus even while two are on screen at once. The curve
          // is steeper than linear so the card being left behind gives up the
          // frame early in the move rather than competing through the whole
          // of it.
          const emphasis = (1 - Math.min(1, Math.abs(index - focus))) ** 1.4;
          const blur = (1 - emphasis) * 4;

          return (
            <div
              key={`${spotlight.file}-${spotlight.slideIndex}`}
              style={{
                position: "absolute",
                left: spotlight.x - SLIDE_W / 2,
                top: spotlight.y - SLIDE_H / 2,
                width: SLIDE_W,
                height: SLIDE_H,
                borderRadius: 10,
                overflow: "hidden",
                boxShadow: "0 40px 100px rgba(0,0,0,.45)",
                opacity: 0.26 + 0.74 * emphasis,
                filter: blur > 0.05 ? `blur(${blur}px)` : undefined,
              }}
            >
              <PptxCard
                file={spotlight.file}
                width={SLIDE_W}
                height={SLIDE_H}
                slideIndex={spotlight.slideIndex}
              />
            </div>
          );
        })}
      </AbsoluteFill>
    </AbsoluteFill>
  );
}

/** Words rise and sharpen in sequence, so a caption lands as a phrase. */
function StaggeredWords({
  text,
  frame,
  delay,
  style,
}: {
  text: string;
  frame: number;
  delay: number;
  style: CSSProperties;
}) {
  return (
    <span style={{ display: "inline-flex", flexWrap: "wrap", ...style }}>
      {text.split(" ").map((word, index) => {
        const start = delay + index * 1.6;
        const reveal = progress(frame, start, start + 17);

        return (
          <span
            key={`${word}-${index}`}
            style={{
              display: "inline-block",
              // The gap is margin rather than a space inside the span, so a
              // wrapped line starts flush instead of indented by it.
              marginRight: "0.27em",
              opacity: reveal,
              transform: `translateY(${(1 - reveal) * 14}px)`,
              filter: reveal < 1 ? `blur(${(1 - reveal) * 5}px)` : undefined,
            }}
          >
            {word}
          </span>
        );
      })}
    </span>
  );
}

/**
 * Captions are stacked in one slot and cross-fade, so the text changes on the
 * spot while the camera does the travelling.
 */
function Captions({ frame }: { frame: number }) {
  return (
    <div style={{ position: "relative", height: 260 }}>
      {SPOTLIGHTS.map((spotlight, index) => {
        const start = index * SPOTLIGHT_DURATION;
        const local = frame - start;
        const isLast = index === SPOTLIGHTS.length - 1;
        // Gone by the time the camera is properly under way. Text sitting over
        // a card that is sliding out from under it is the one thing that reads
        // as clutter here, so the column empties first and the move gets the
        // frame to itself.
        const exit = isLast
          ? 1
          : interpolate(local, [CAMERA_HOLD - 2, CAMERA_HOLD + 8], [1, 0], clamp);

        // Started just before the camera settles, so the column is never empty
        // on arrival. The lead is short enough that the words are still rising
        // as the card comes to rest, which ties the two together.
        const reveal = local + CAPTION_LEAD;
        if (reveal < 0 || exit === 0) return null;

        return (
          <div
            key={`${spotlight.file}-${spotlight.slideIndex}`}
            style={{ position: "absolute", inset: 0, opacity: exit }}
          >
            <StaggeredWords
              text={spotlight.caption}
              frame={reveal}
              delay={0}
              style={{
                ...font,
                fontSize: 62,
                fontWeight: 800,
                color: C.white,
                letterSpacing: -1.6,
                lineHeight: 1.12,
                // Cards pass behind the column, dimmed and blurred but still
                // bright enough on the lighter decks to eat into the text. A
                // shadow carries that on the few hundred pixels where they
                // overlap, without a scrim muting the card itself.
                textShadow: TEXT_ON_BOARD,
              }}
            />
            <div style={{ height: 18 }} />
            <StaggeredWords
              text={spotlight.description}
              frame={reveal}
              delay={10}
              style={{
                ...font,
                fontSize: 27,
                fontWeight: 400,
                color: C.muted,
                lineHeight: 1.5,
                textShadow: TEXT_ON_BOARD,
              }}
            />
          </div>
        );
      })}
    </div>
  );
}

function ShowcaseScene() {
  const frame = useCurrentFrame();

  // Every other section animates its own way in, since content is held back
  // past the crossfade rather than fading across it. Without this the board
  // arrived in one frame, which is what made the cut from the title read as a
  // jump. The camera settles the last of the way in as it appears.
  const entrance = progress(frame, 0, 20);

  return (
    <AbsoluteFill style={{ opacity: entrance }}>
      <AbsoluteFill style={{ transform: `scale(${interpolate(entrance, [0, 1], [0.985, 1])})` }}>
        <Board frame={frame} />
      </AbsoluteFill>
      <div
        style={{
          position: "absolute",
          left: 120,
          top: 0,
          width: 620,
          height: 1080,
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
        }}
      >
        <Captions frame={frame} />
      </div>
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

const SPOTLIGHTS_TOTAL = sectionDuration(SHOWCASE_FRAMES);

export function Launch() {
  return (
    <TransitionSeries>
      <TransitionSeries.Sequence durationInFrames={TITLE_DURATION}>
        <TitleCard />
      </TransitionSeries.Sequence>

      <TransitionSeries.Transition presentation={fade()} timing={fadeTiming} />

      {/* One continuous shot: every card is mounted on the board for the whole
          run and the camera moves between them, so there are no cuts for the
          cards to fade across. The run shares one backdrop, leaving the
          envelope only the board and the text column to fade. */}
      <TransitionSeries.Sequence durationInFrames={SPOTLIGHTS_TOTAL}>
        <AbsoluteFill>
          <Backdrop />
          <SceneContent durationInFrames={SPOTLIGHTS_TOTAL}>
            <Sequence from={CONTENT_LEAD} layout="none">
              <ShowcaseScene />
            </Sequence>
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
