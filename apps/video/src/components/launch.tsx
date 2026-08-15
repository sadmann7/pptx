import * as React from "react";

import { FocusBlurResolve } from "@pptx/ui/components/remocn/focus-blur-resolve";
import { SoftBlurIn } from "@pptx/ui/components/remocn/soft-blur-in";
import { Typewriter } from "@pptx/ui/components/remocn/typewriter";
import { Video } from "@remotion/media";
import { linearTiming, TransitionSeries } from "@remotion/transitions";
import { fade } from "@remotion/transitions/fade";
import { AbsoluteFill, Easing, interpolate, Sequence, staticFile, useCurrentFrame } from "remotion";

import { EditorPreview, type EditorPreviewReveal } from "@/components/editor-preview";
import { PptxCard } from "@/components/pptx-card";
import { geistMono, geistSans } from "@/lib/fonts";

const C = {
  ink: "#0e1117",
  white: "#f5f6f8",
  muted: "#8b92a8",
};

const clamp = {
  extrapolateLeft: "clamp",
  extrapolateRight: "clamp",
} as const;
const ease = Easing.bezier(0.22, 1, 0.36, 1);
const cameraEase = Easing.bezier(0.7, 0, 0.25, 1);

const FADE_DURATION = 12;
const CONTENT_FADE = 10;
/** Wait out the incoming crossfade; each scene animates its own entrance. */
const CONTENT_LEAD = FADE_DURATION;

const TITLE_DURATION = 60;
const SPOTLIGHT_DURATION = 62;
const CAMERA_MOVE = 20;
const CAMERA_HOLD = SPOTLIGHT_DURATION - CAMERA_MOVE;
const CAPTION_LEAD = 8;
const SNAP_BEAT = 47;
const COMPOSITION_CONTENT_FRAMES = 132;
const CTA_DURATION = CONTENT_LEAD + 90;
const DEMO_CONTENT_FRAMES = 210;

const SLIDE_W = 960;
const SLIDE_H = 540;
const FOCUS_X = 1290;
const FOCUS_Y = 540;
const TEXT_ON_BOARD = "0 2px 14px rgba(14,17,23,.85), 0 0 44px rgba(14,17,23,.75)";

const font: React.CSSProperties = {
  fontFamily: geistSans,
};

const fadeTiming = linearTiming({ durationInFrames: FADE_DURATION });

function sectionDuration(contentFrames: number) {
  return CONTENT_LEAD + contentFrames + FADE_DURATION;
}

function progress(frame: number, from: number, to: number) {
  return interpolate(frame, [from, to], [0, 1], { ...clamp, easing: ease });
}

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

/** Fade content out before the section crossfade so two scenes don't stack. */
function SceneContent({
  durationInFrames,
  fadeOut = true,
  children,
}: {
  durationInFrames: number;
  fadeOut?: boolean;
  children: React.ReactNode;
}) {
  const frame = useCurrentFrame();
  const exitStart = durationInFrames - FADE_DURATION - CONTENT_FADE;
  const opacity = fadeOut
    ? interpolate(frame, [exitStart, exitStart + CONTENT_FADE], [1, 0], clamp)
    : 1;

  return <AbsoluteFill style={{ opacity }}>{children}</AbsoluteFill>;
}

interface Spotlight {
  file: string;
  slideIndex: number;
  caption: string;
  x: number;
  y: number;
  zoom: number;
}

const SPOTLIGHTS: Spotlight[] = [
  {
    file: "editorial-forest-editable.pptx",
    slideIndex: 1,
    caption: "As in PowerPoint",
    x: 0,
    y: -120,
    zoom: 1.06,
  },
  {
    file: "pocket-machines-sakura-chroma.pptx",
    slideIndex: 0,
    caption: "Native gradients",
    x: 1420,
    y: 280,
    zoom: 1.02,
  },
  {
    file: "adventure-club-pin-and-paper.pptx",
    slideIndex: 3,
    caption: "Charts, not images",
    x: 2820,
    y: -230,
    zoom: 1.08,
  },
  {
    file: "make-something-strange-creative-mode.pptx",
    slideIndex: 6,
    caption: "Tables and borders",
    x: 4240,
    y: 240,
    zoom: 1.03,
  },
  {
    file: "after-the-needle-drops-mat.pptx",
    slideIndex: 0,
    caption: "Embedded fonts",
    x: 5640,
    y: -270,
    zoom: 1.07,
  },
];

const SHOWCASE_FRAMES = SPOTLIGHTS.length * SPOTLIGHT_DURATION;
const SPOTLIGHTS_TOTAL = sectionDuration(SHOWCASE_FRAMES);

interface Camera {
  x: number;
  y: number;
  zoom: number;
}

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

interface StaggeredWordsProps extends React.ComponentProps<"span"> {
  text: string;
  frame: number;
  delay?: number;
}

function StaggeredWords({ text, frame, delay = 0, style, ...props }: StaggeredWordsProps) {
  return (
    <span style={{ display: "inline-flex", flexWrap: "wrap", ...style }} {...props}>
      {text.split(" ").map((word, index) => {
        const start = delay + index * 1.6;
        const reveal = progress(frame, start, start + 17);

        return (
          <span
            key={`${word}-${index}`}
            style={{
              display: "inline-block",
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

interface CaptionsProps extends React.ComponentProps<"div"> {
  frame: number;
}

function Captions({ frame, style, ...props }: CaptionsProps) {
  return (
    <div style={{ position: "relative", height: 260, ...style }} {...props}>
      {SPOTLIGHTS.map((spotlight, index) => {
        const start = index * SPOTLIGHT_DURATION;
        const local = frame - start;
        const isLast = index === SPOTLIGHTS.length - 1;
        const exit = isLast
          ? 1
          : interpolate(local, [CAMERA_HOLD - 2, CAMERA_HOLD + 8], [1, 0], clamp);
        const reveal = local + CAPTION_LEAD;
        if (reveal < 0 || exit === 0) return null;

        return (
          <div
            key={`${spotlight.file}-${spotlight.slideIndex}`}
            style={{
              position: "absolute",
              inset: 0,
              display: "flex",
              alignItems: "center",
              opacity: exit,
            }}
          >
            <StaggeredWords
              text={spotlight.caption}
              frame={reveal}
              delay={0}
              style={{
                ...font,
                fontSize: 76,
                fontWeight: 800,
                color: C.white,
                letterSpacing: -2,
                lineHeight: 1.12,
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

const SNAPS = [
  "Parse. Render. Edit. Re-export.",
  "Tables, charts, shapes, images.",
  "TypeScript-first. No server required.",
];

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

const COMPOSITION_LINES = [
  "<Presentation.Root>",
  "  <Presentation.ThumbnailList />",
  "  <Presentation.Viewport>",
  "    <Presentation.Slide>",
  "      <Presentation.Selection />",
  "    </Presentation.Slide>",
  "  </Presentation.Viewport>",
  "</Presentation.Root>",
];

const COMPOSITION_DURATION = sectionDuration(COMPOSITION_CONTENT_FRAMES);

/**
 * The tag colour is github-dark's, matching the docs site. The namespace is
 * dimmed rather than sharing it: `Presentation` repeats on all sixteen tags and
 * would otherwise drown out the five component names, which are the point.
 */
const CODE_COLORS = {
  tag: "#7ee787",
  namespace: "#7d8590",
  plain: "#5b636d",
} as const;

const TOKEN_PATTERN = /(?<namespace>Presentation)|(?<tag>[A-Z][\w$]*)|(?<plain>[^A-Z]+)/gu;

function tokenize(line: string): { text: string; kind: keyof typeof CODE_COLORS }[] {
  return Array.from(line.matchAll(TOKEN_PATTERN), (match) => {
    const groups = match.groups ?? {};
    const kind = groups.namespace ? "namespace" : groups.tag ? "tag" : "plain";
    return { text: match[0], kind };
  });
}

function CodeLine({ line, reveal }: { line: string; reveal: number }) {
  return (
    <div
      style={{
        opacity: reveal,
        transform: `translateY(${(1 - reveal) * 10}px)`,
      }}
    >
      {tokenize(line).map((token, index) => (
        <span key={`${index}-${token.text}`} style={{ color: CODE_COLORS[token.kind] }}>
          {token.text}
        </span>
      ))}
    </div>
  );
}

/** Frames between one line appearing and the next. */
const LINE_STAGGER = 9;
const LINE_REVEAL = 15;
/** The preview trails the line that introduces it, so the code reads as cause. */
const PREVIEW_LAG = 5;

const lineStart = (index: number) => index * LINE_STAGGER;

function CompositionCode({ time }: { time: number }) {
  return (
    <pre
      style={{
        margin: 0,
        fontFamily: geistMono,
        fontSize: 34,
        fontWeight: 500,
        lineHeight: 1.6,
        letterSpacing: -0.4,
      }}
    >
      {COMPOSITION_LINES.map((line, index) => (
        <CodeLine
          key={line}
          line={line}
          reveal={progress(time, lineStart(index), lineStart(index) + LINE_REVEAL)}
        />
      ))}
    </pre>
  );
}

const PREVIEW_FILE = "the-good-room-soft-editorial.pptx";
const PREVIEW_SLIDE = 0;
const PREVIEW_RAIL_W = 130;
const PREVIEW_PAD = 26;
const PREVIEW_W = 980;
/** Sized so the 16:9 slide fills the canvas exactly, leaving no dead band. */
const PREVIEW_H =
  Math.round(((PREVIEW_W - PREVIEW_RAIL_W - PREVIEW_PAD * 2) * 9) / 16) + PREVIEW_PAD * 2;

/**
 * Which line of the snippet each part of the preview belongs to, so the editor
 * assembles itself in step with the code that declares it.
 */
const PREVIEW_STAGES = {
  chrome: 0,
  rail: 1,
  canvas: 2,
  slide: 3,
  selection: 4,
} as const;

type PreviewStageEntries = {
  [K in keyof typeof PREVIEW_STAGES]: [K, (typeof PREVIEW_STAGES)[K]];
}[keyof typeof PREVIEW_STAGES][];

function previewStageEntries(): PreviewStageEntries {
  return Object.entries(PREVIEW_STAGES) as PreviewStageEntries;
}

function buildPreviewReveal(time: number): EditorPreviewReveal {
  const reveal: EditorPreviewReveal = {
    chrome: 0,
    rail: 0,
    canvas: 0,
    slide: 0,
    selection: 0,
  };
  for (const [part, line] of previewStageEntries()) {
    const from = lineStart(line) + PREVIEW_LAG;
    reveal[part] = progress(time, from, from + LINE_REVEAL);
  }
  return reveal;
}

function CompositionScene() {
  const frame = useCurrentFrame();
  const time = frame - CONTENT_LEAD;
  const reveal = buildPreviewReveal(time);

  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={COMPOSITION_DURATION}>
        <AbsoluteFill
          style={{
            ...font,
            flexDirection: "row",
            alignItems: "center",
            justifyContent: "center",
            gap: 72,
          }}
        >
          <CompositionCode time={time} />
          <EditorPreview
            file={PREVIEW_FILE}
            slideIndex={PREVIEW_SLIDE}
            width={PREVIEW_W}
            height={PREVIEW_H}
            railWidth={PREVIEW_RAIL_W}
            padding={PREVIEW_PAD}
            reveal={reveal}
          />
        </AbsoluteFill>
      </SceneContent>
    </AbsoluteFill>
  );
}

/** Drop `public/editor-demo.mp4` in to insert the editing scene. */
export const EDITOR_DEMO_FILE = "editor-demo.mp4";

const DEMO_DURATION = sectionDuration(DEMO_CONTENT_FRAMES);

function DemoScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={DEMO_DURATION}>
        <Sequence from={CONTENT_LEAD} durationInFrames={DEMO_CONTENT_FRAMES} layout="none">
          <Video
            src={staticFile(EDITOR_DEMO_FILE)}
            style={{ width: "100%", height: "100%", objectFit: "cover" }}
          />
        </Sequence>
      </SceneContent>
    </AbsoluteFill>
  );
}

function CtaScene() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={CTA_DURATION} fadeOut={false}>
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

export type LaunchProps = {
  hasEditorDemo: boolean;
} & Record<string, unknown>;

export const DEFAULT_LAUNCH_PROPS: LaunchProps = {
  hasEditorDemo: false,
};

export function launchDuration(hasEditorDemo: boolean): number {
  const fades = hasEditorDemo ? 5 : 4;
  return (
    TITLE_DURATION +
    SPOTLIGHTS_TOTAL +
    (hasEditorDemo ? DEMO_DURATION : 0) +
    FEATURES_DURATION +
    COMPOSITION_DURATION +
    CTA_DURATION -
    fades * FADE_DURATION
  );
}

export const LAUNCH_DURATION = launchDuration(false);

export function Launch({ hasEditorDemo = false }: LaunchProps) {
  const scenes = [
    <TransitionSeries.Sequence key="title" durationInFrames={TITLE_DURATION}>
      <TitleCard />
    </TransitionSeries.Sequence>,
    <TransitionSeries.Sequence key="showcase" durationInFrames={SPOTLIGHTS_TOTAL}>
      <AbsoluteFill>
        <Backdrop />
        <SceneContent durationInFrames={SPOTLIGHTS_TOTAL}>
          <Sequence from={CONTENT_LEAD} layout="none">
            <ShowcaseScene />
          </Sequence>
        </SceneContent>
      </AbsoluteFill>
    </TransitionSeries.Sequence>,
    ...(hasEditorDemo
      ? [
          <TransitionSeries.Sequence key="demo" durationInFrames={DEMO_DURATION}>
            <DemoScene />
          </TransitionSeries.Sequence>,
        ]
      : []),
    <TransitionSeries.Sequence key="features" durationInFrames={FEATURES_DURATION}>
      <FeaturesScene />
    </TransitionSeries.Sequence>,
    <TransitionSeries.Sequence key="composition" durationInFrames={COMPOSITION_DURATION}>
      <CompositionScene />
    </TransitionSeries.Sequence>,
    <TransitionSeries.Sequence key="cta" durationInFrames={CTA_DURATION}>
      <CtaScene />
    </TransitionSeries.Sequence>,
  ];

  return (
    <TransitionSeries>
      {scenes.flatMap((scene, index) =>
        index === 0
          ? [scene]
          : [
              <TransitionSeries.Transition
                key={`cut-${index}`}
                presentation={fade()}
                timing={fadeTiming}
              />,
              scene,
            ],
      )}
    </TransitionSeries>
  );
}
