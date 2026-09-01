import type * as React from "react";

import { AbsoluteFill, Sequence, interpolate, useCurrentFrame } from "remotion";

import { Backdrop, SceneContent } from "@/components/backdrop";
import { PptxCard } from "@/components/pptx-card";
import {
  CAMERA_HOLD,
  CAMERA_MOVE,
  CAPTION_LEAD,
  CONTENT_LEAD,
  FLUSH_FRAMES,
  FOCUS_X,
  FOCUS_Y,
  HANDOFF_FRAMES,
  SHOWCASE_FRAMES,
  SLIDE_H,
  SLIDE_W,
  SPOTLIGHTS,
  SPOTLIGHTS_TOTAL,
  SPOTLIGHT_DURATION,
  TEXT_ON_BOARD,
} from "@/lib/constants";
import { previewSlideRect } from "@/lib/layout";
import { color } from "@/lib/theme";
import { cameraEase, clamp, progress } from "@/lib/timing";

interface Camera {
  x: number;
  y: number;
  zoom: number;
}

function landingCamera(): Camera {
  const spotlight = SPOTLIGHTS.at(-1);
  if (!spotlight) throw new Error("no last spotlight");

  // The rail has not landed yet when the showcase hands off.
  const slide = previewSlideRect(0);

  return {
    x: spotlight.x - (slide.left + slide.width / 2 - FOCUS_X) / slide.zoom,
    y: spotlight.y - (slide.top + slide.height / 2 - FOCUS_Y) / slide.zoom,
    zoom: slide.zoom,
  };
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

const captionStyle: React.CSSProperties = {
  fontSize: 76,
  fontWeight: 800,
  color: color.white,
  letterSpacing: -2,
  lineHeight: 1.12,
  textShadow: TEXT_ON_BOARD,
};

function Captions({ frame }: { frame: number }) {
  return (
    <div style={{ position: "relative", height: 260 }}>
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
            <StaggeredWords text={spotlight.caption} frame={reveal} style={captionStyle} />
          </div>
        );
      })}
    </div>
  );
}

function ShowcaseBoard() {
  const frame = useCurrentFrame();
  const entrance = progress(frame, 0, 20);

  return (
    <AbsoluteFill style={{ opacity: entrance }}>
      <AbsoluteFill style={{ scale: interpolate(entrance, [0, 1], [0.985, 1]) }}>
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

function ShowcaseBackdrop() {
  const frame = useCurrentFrame();
  if (frame >= SPOTLIGHTS_TOTAL - HANDOFF_FRAMES) return null;
  return <Backdrop />;
}

export function ShowcaseScene() {
  return (
    <AbsoluteFill>
      <ShowcaseBackdrop />
      <SceneContent durationInFrames={SPOTLIGHTS_TOTAL} fadeOut={false}>
        <Sequence from={CONTENT_LEAD} layout="none">
          <ShowcaseBoard />
        </Sequence>
      </SceneContent>
    </AbsoluteFill>
  );
}
