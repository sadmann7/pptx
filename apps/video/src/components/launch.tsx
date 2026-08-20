import * as React from "react";

import { FocusBlurResolve } from "@pptx/ui/components/remocn/focus-blur-resolve";
import { SoftBlurIn } from "@pptx/ui/components/remocn/soft-blur-in";
import { Typewriter } from "@pptx/ui/components/remocn/typewriter";
import { Video } from "@remotion/media";
import type {
  TransitionPresentation,
  TransitionPresentationComponentProps,
} from "@remotion/transitions";
import { linearTiming, TransitionSeries } from "@remotion/transitions";
import { fade } from "@remotion/transitions/fade";
import { AbsoluteFill, Easing, interpolate, Sequence, staticFile, useCurrentFrame } from "remotion";
import type { ThemedToken } from "shiki/core";

import { EditorPreview } from "@/components/editor-preview";
import { PptxCard } from "@/components/pptx-card";
import { fontVars, geistMono, geistSans } from "@/lib/fonts";
import { useHighlightedLines } from "@/lib/highlight";
import { panelShadow } from "@/lib/theme";

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
const COMPOSITION_CONTENT_FRAMES = 170;
const CTA_DURATION = CONTENT_LEAD + 90;
const DEMO_CONTENT_FRAMES = 210;

const SLIDE_W = 960;
const SLIDE_H = 540;
const FOCUS_X = 1290;
const FOCUS_Y = 540;
const TEXT_ON_BOARD = "0 2px 14px rgba(14,17,23,.85), 0 0 44px rgba(14,17,23,.75)";
const VIDEO_W = 1920;
const VIDEO_H = 1080;
const HANDOFF_FRAMES = 36;
/** How long the flying card keeps its lift off the board. */
const FLUSH_FRAMES = 12;
const COMPOSITION_GAP = 72;
const PREVIEW_RAIL_W = 130;
const PREVIEW_PAD = 26;
const PREVIEW_W = 980;
/** Sized so the 16:9 slide fills the canvas exactly, leaving no dead band. */
const PREVIEW_H =
  Math.round(((PREVIEW_W - PREVIEW_RAIL_W - PREVIEW_PAD * 2) * 9) / 16) + PREVIEW_PAD * 2;

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
const SPOTLIGHTS_TOTAL = CONTENT_LEAD + SHOWCASE_FRAMES + HANDOFF_FRAMES;

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
  if (frame >= SHOWCASE_FRAMES) {
    const t = progress(frame, SHOWCASE_FRAMES, SHOWCASE_FRAMES + HANDOFF_FRAMES);
    const from = restingCamera(SPOTLIGHTS.length - 1, 1);
    const to = landingCamera();
    return {
      x: from.x + (to.x - from.x) * t,
      y: from.y + (to.y - from.y) * t,
      zoom: from.zoom + (to.zoom - from.zoom) * t,
    };
  }

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
  // Lift and rounding are gone by the time the card reaches the editor, which
  // the camera does well before the handoff ends. A slide on a canvas does not
  // float, so carrying the shadow the rest of the way just leaves a glow
  // sitting inside the window until the cut.
  const flush = progress(frame, SHOWCASE_FRAMES, SHOWCASE_FRAMES + FLUSH_FRAMES);
  const othersOut = progress(frame, SHOWCASE_FRAMES - 18, SHOWCASE_FRAMES + 6);

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
          const isLast = index === SPOTLIGHTS.length - 1;
          const emphasis = (1 - Math.min(1, Math.abs(index - focus))) ** 1.4;
          const blur = (1 - emphasis) * 4;
          // The last card is left at full opacity and simply unmounts with the
          // scene. By then it is nested exactly over the editor's own copy of
          // the same slide, so the cut is invisible, where a crossfade would
          // put both on fractional opacity and dim the pair for its duration.
          const opacity = (0.26 + 0.74 * emphasis) * (isLast ? 1 : 1 - othersOut);

          return (
            <div
              key={`${spotlight.file}-${spotlight.slideIndex}`}
              style={{
                position: "absolute",
                left: spotlight.x - SLIDE_W / 2,
                top: spotlight.y - SLIDE_H / 2,
                width: SLIDE_W,
                height: SLIDE_H,
                borderRadius: isLast ? 10 * (1 - flush) : 10,
                overflow: "hidden",
                boxShadow: isLast
                  ? `0 ${40 * (1 - flush)}px ${100 * (1 - flush)}px rgba(0,0,0,${0.45 * (1 - flush)})`
                  : "0 40px 100px rgba(0,0,0,.45)",
                opacity,
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
          ? interpolate(frame, [SHOWCASE_FRAMES - 22, SHOWCASE_FRAMES - 4], [1, 0], clamp)
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

/**
 * The showcase sits above the composition for the handoff, so its backdrop has
 * to get out of the way. Dropping it in one step is invisible because the scene
 * underneath is already painting the same gradient at full opacity, whereas
 * fading it would tint every pixel for the length of the ramp.
 */
function ShowcaseBackdrop() {
  const frame = useCurrentFrame();
  if (frame >= SPOTLIGHTS_TOTAL - HANDOFF_FRAMES) return null;
  return <Backdrop />;
}

function TitleCard() {
  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={TITLE_DURATION}>
        <AbsoluteFill
          style={{
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

const RAIL_LINE_ID = "rail";

const COMPOSITION_LINES: Record<string, string> = {
  "root-open": "<Presentation.Root>",
  [RAIL_LINE_ID]: "  <Presentation.ThumbnailList />",
  "viewport-open": "  <Presentation.Viewport>",
  "slide-open": "    <Presentation.Slide>",
  selection: "      <Presentation.Selection />",
  "slide-close": "    </Presentation.Slide>",
  "viewport-close": "  </Presentation.Viewport>",
  "root-close": "</Presentation.Root>",
};

/** A slide on its own, then the rail dropped in as a sibling of the canvas. */
const WITHOUT_RAIL = [
  "root-open",
  "viewport-open",
  "slide-open",
  "selection",
  "slide-close",
  "viewport-close",
  "root-close",
];

const WITH_RAIL = [
  "root-open",
  RAIL_LINE_ID,
  "viewport-open",
  "slide-open",
  "selection",
  "slide-close",
  "viewport-close",
  "root-close",
];

const COMPOSITION_DURATION = sectionDuration(COMPOSITION_CONTENT_FRAMES);

/** Highlighted as one document so each line is tokenised in context. */
const COMPOSITION_SOURCE = WITH_RAIL.map((id) => COMPOSITION_LINES[id]).join("\n");

const CODE_FONT_SIZE = 34;
const CODE_LINE_H = Math.round(CODE_FONT_SIZE * 1.6);
/**
 * Every line is positioned rather than laid out, so none of them give the block
 * a width. In a monospace face the longest line is exactly this many characters.
 */
const CODE_COLUMNS = Math.max(...Object.values(COMPOSITION_LINES).map((line) => line.length));
const CODE_PANEL_PAD_X = 40;
const CODE_PANEL_W = CODE_PANEL_PAD_X * 2 + Math.round(CODE_COLUMNS * CODE_FONT_SIZE * 0.6);

/**
 * Camera that places the last showcase slide on the composition preview's
 * empty canvas, so the two copies occupy the same pixels at the cut.
 */
function landingCamera(): Camera {
  const spotlight = SPOTLIGHTS.at(-1);
  if (!spotlight) throw new Error("no last spotlight");

  const rowWidth = CODE_PANEL_W + COMPOSITION_GAP + PREVIEW_W;
  const previewLeft = (VIDEO_W - rowWidth) / 2 + CODE_PANEL_W + COMPOSITION_GAP;
  const previewTop = (VIDEO_H - PREVIEW_H) / 2;
  const slideHeight = PREVIEW_H - PREVIEW_PAD * 2;
  const slideWidth = (slideHeight * 16) / 9;
  const slideLeft = previewLeft + (PREVIEW_W - slideWidth) / 2;
  const slideTop = previewTop + PREVIEW_PAD;
  const zoom = slideWidth / SLIDE_W;

  return {
    x: spotlight.x - (slideLeft + slideWidth / 2 - FOCUS_X) / zoom,
    y: spotlight.y - (slideTop + slideHeight / 2 - FOCUS_Y) / zoom,
    zoom,
  };
}

function CodeLine({ tokens, row, mark, opacity }: CodeLineProps) {
  return (
    <div
      style={{
        position: "absolute",
        top: 0,
        left: 0,
        whiteSpace: "pre",
        height: CODE_LINE_H,
        lineHeight: `${CODE_LINE_H}px`,
        opacity,
        // Lifted while the lines below make room, so it does not collide with them.
        zIndex: mark > 0 ? 1 : 0,
        transform: `translateY(${row * CODE_LINE_H}px)`,
      }}
    >
      <span
        style={{
          position: "absolute",
          inset: "0 -20px",
          borderRadius: 8,
          // GitHub Dark paints React components blue, so the insert marker uses
          // that same token rather than the green added-line wash.
          backgroundColor: `rgba(13,17,23,${0.94 * mark})`,
          backgroundImage: `linear-gradient(rgba(121,192,255,${0.16 * mark}), rgba(121,192,255,${0.16 * mark}))`,
          boxShadow: `inset 3px 0 0 rgba(121,192,255,${mark})`,
        }}
      />
      {tokens.map((token, index) => (
        <span
          key={`${index}-${token.content}`}
          style={{ position: "relative", color: token.color }}
        >
          {token.content}
        </span>
      ))}
    </div>
  );
}

interface CodeLineProps {
  tokens: ThemedToken[];
  row: number;
  mark: number;
  opacity: number;
}

function CompositionCode({ shift }: { shift: number }) {
  const lines = useHighlightedLines(COMPOSITION_SOURCE);

  return (
    <div
      style={{
        position: "relative",
        width: `${CODE_COLUMNS}ch`,
        height: WITH_RAIL.length * CODE_LINE_H,
        fontFamily: geistMono,
        fontSize: CODE_FONT_SIZE,
        fontWeight: 500,
        letterSpacing: -0.4,
      }}
    >
      {WITH_RAIL.map((id, to) => {
        const tokens = lines?.[to];
        if (!tokens) return null;
        const from = id === RAIL_LINE_ID ? to : WITHOUT_RAIL.indexOf(id);
        return (
          <CodeLine
            key={id}
            tokens={tokens}
            row={from + (to - from) * shift}
            mark={id === RAIL_LINE_ID ? Math.sin(shift * Math.PI) : 0}
            opacity={id === RAIL_LINE_ID ? shift : 1}
          />
        );
      })}
    </div>
  );
}

const PREVIEW_FILE = "after-the-needle-drops-mat.pptx";
const PREVIEW_SLIDE = 0;

/** Late enough that the whole snippet has been on screen long enough to read. */
const MOVE_START = 72;
const MOVE_FRAMES = 26;

/** Frames between one miniature starting its fade and the next one starting. */
const RAIL_POP_CADENCE = 4;

/**
 * Surface behind the snippet. Same material and height as the editor window, so
 * the two columns read as a pair rather than as text floating on the gradient.
 */
function CodePanel({ reveal, children }: { reveal: number; children: React.ReactNode }) {
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        width: CODE_PANEL_W,
        height: PREVIEW_H,
        padding: `0 ${CODE_PANEL_PAD_X}px`,
        boxSizing: "border-box",
        borderRadius: 18,
        background: "rgba(255,255,255,.04)",
        border: "1px solid rgba(255,255,255,.12)",
        boxShadow: panelShadow(reveal),
        opacity: reveal,
      }}
    >
      {children}
    </div>
  );
}

/**
 * A full-bleed slide and the same window with a rail, each mounted for the
 * whole scene and settled before it is seen. Cutting between them is the only
 * way the rail can appear without throwing its IntersectionObserver.
 */
function PreviewSwap({
  reveal,
  swap,
  railReveal,
}: {
  reveal: number;
  swap: number;
  railReveal: number;
}) {
  const common = {
    file: PREVIEW_FILE,
    slideIndex: PREVIEW_SLIDE,
    width: PREVIEW_W,
    height: PREVIEW_H,
    railWidth: PREVIEW_RAIL_W,
    padding: PREVIEW_PAD,
  };
  const layer: React.CSSProperties = { position: "absolute", inset: 0 };

  return (
    <div style={{ position: "relative", width: PREVIEW_W, height: PREVIEW_H, flex: "0 0 auto" }}>
      <EditorPreview {...common} showRail={false} reveal={reveal * (1 - swap)} style={layer} />
      <EditorPreview {...common} reveal={reveal * swap} railReveal={railReveal} style={layer} />
    </div>
  );
}

function CompositionScene() {
  const time = useCurrentFrame();
  const reveal = progress(time, 8, 28);
  const shift = progress(time, MOVE_START, MOVE_START + MOVE_FRAMES);
  const swap = progress(time, MOVE_START + MOVE_FRAMES - 4, MOVE_START + MOVE_FRAMES);
  const railReveal = Math.max(0, (time - (MOVE_START + MOVE_FRAMES)) / RAIL_POP_CADENCE);

  return (
    <AbsoluteFill>
      <Backdrop />
      <SceneContent durationInFrames={COMPOSITION_DURATION} fadeOut={false}>
        <AbsoluteFill
          style={{
            flexDirection: "row",
            alignItems: "center",
            justifyContent: "center",
            gap: COMPOSITION_GAP,
          }}
        >
          <CodePanel reveal={reveal}>
            <CompositionCode shift={shift} />
          </CodePanel>
          <PreviewSwap reveal={reveal} swap={swap} railReveal={railReveal} />
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

/**
 * Both scenes stay fully opaque and the showcase stays on top, so the card
 * still flying lands over an editor that is already fully drawn. Fading either
 * scene would be worse than a cut: a fractional opacity puts the backdrop
 * gradient on its own compositing surface, which Chromium renders about a luma
 * darker, and the step back at the end of the ramp reads as a flash.
 */
function persistPresentation(): TransitionPresentation<Record<string, never>> {
  return { component: PersistPresentation, props: {} };
}

function PersistPresentation({
  children,
  presentationDirection,
}: TransitionPresentationComponentProps<Record<string, never>>) {
  return (
    <AbsoluteFill style={{ zIndex: presentationDirection === "exiting" ? 1 : 0 }}>
      {children}
    </AbsoluteFill>
  );
}

const handoffTiming = linearTiming({ durationInFrames: HANDOFF_FRAMES });

export type LaunchProps = {
  hasEditorDemo: boolean;
} & Record<string, unknown>;

export const DEFAULT_LAUNCH_PROPS: LaunchProps = {
  hasEditorDemo: false,
};

export function launchDuration(hasEditorDemo: boolean): number {
  const shortFades = hasEditorDemo ? 4 : 3;
  return (
    TITLE_DURATION +
    SPOTLIGHTS_TOTAL +
    (hasEditorDemo ? DEMO_DURATION : 0) +
    FEATURES_DURATION +
    COMPOSITION_DURATION +
    CTA_DURATION -
    shortFades * FADE_DURATION -
    (hasEditorDemo ? FADE_DURATION : HANDOFF_FRAMES)
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
        <ShowcaseBackdrop />
        <SceneContent durationInFrames={SPOTLIGHTS_TOTAL} fadeOut={false}>
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
    // Composition before Features: the showcase has just proved what it
    // renders, so "how you build it" follows while that is fresh, and the
    // summary claims run into the install rather than splitting the two
    // scenes that actually show something.
    <TransitionSeries.Sequence key="composition" durationInFrames={COMPOSITION_DURATION}>
      <CompositionScene />
    </TransitionSeries.Sequence>,
    <TransitionSeries.Sequence key="features" durationInFrames={FEATURES_DURATION}>
      <FeaturesScene />
    </TransitionSeries.Sequence>,
    <TransitionSeries.Sequence key="cta" durationInFrames={CTA_DURATION}>
      <CtaScene />
    </TransitionSeries.Sequence>,
  ];

  return (
    <AbsoluteFill style={{ ...fontVars, fontFamily: geistSans }}>
      <TransitionSeries>
        {scenes.flatMap((scene, index) => {
          if (index === 0) return [scene];
          const from = scenes[index - 1]?.key;
          const handoff = from === "showcase" && scene.key === "composition";
          return [
            handoff ? (
              <TransitionSeries.Transition
                key={`cut-${index}`}
                presentation={persistPresentation()}
                timing={handoffTiming}
              />
            ) : (
              <TransitionSeries.Transition
                key={`cut-${index}`}
                presentation={fade()}
                timing={fadeTiming}
              />
            ),
            scene,
          ];
        })}
      </TransitionSeries>
    </AbsoluteFill>
  );
}
