import type * as React from "react";

import { AbsoluteFill, useCurrentFrame } from "remotion";
import type { ThemedToken } from "shiki/core";

import { Backdrop, SceneContent } from "@/components/backdrop";
import { EditorPreview } from "@/components/editor-preview";
import { useHighlightedLines } from "@/hooks/use-highlighted-lines";
import {
  CODE_COLUMNS,
  CODE_FONT_SIZE,
  CODE_LINE_H,
  CODE_PANEL_PAD_X,
  CODE_PANEL_W,
  COMPOSITION_DURATION,
  COMPOSITION_GAP,
  COMPOSITION_SOURCE,
  HANDOFF_FADE_FRAMES,
  HANDOFF_FRAMES,
  MOVE_FRAMES,
  MOVE_START,
  PREVIEW_FILE,
  PREVIEW_H,
  PREVIEW_PAD,
  PREVIEW_RAIL_W,
  PREVIEW_SLIDE,
  PREVIEW_W,
  RAIL_IN_FRAMES,
  RAIL_LINE_ID,
  RAIL_POP_CADENCE,
  WITHOUT_RAIL,
  WITH_RAIL,
} from "@/lib/constants";
import { geistMono } from "@/lib/fonts";
import { panelBox } from "@/lib/theme";
import { progress } from "@/lib/timing";

interface CodePanelProps extends React.ComponentProps<"div"> {
  reveal: number;
}

function CodePanel({ reveal, children, style, ...props }: CodePanelProps) {
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        width: CODE_PANEL_W,
        height: PREVIEW_H,
        padding: `0 ${CODE_PANEL_PAD_X}px`,
        boxSizing: "border-box",
        opacity: reveal,
        ...panelBox(reveal),
        ...style,
      }}
      {...props}
    >
      {children}
    </div>
  );
}

interface CodeLineProps {
  tokens: ThemedToken[];
  row: number;
  mark: number;
  opacity: number;
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
        zIndex: mark > 0 ? 1 : 0,
        transform: `translateY(${row * CODE_LINE_H}px)`,
      }}
    >
      <span
        style={{
          position: "absolute",
          inset: "0 -20px",
          borderRadius: 8,
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

interface CompositionCodeProps {
  shift: number;
}

function CompositionCode({ shift }: CompositionCodeProps) {
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

function CompositionBackdrop({ handoff }: { handoff: boolean }) {
  const time = useCurrentFrame();
  if (handoff && time >= COMPOSITION_DURATION - HANDOFF_FRAMES) return null;
  return <Backdrop />;
}

interface CompositionSceneProps {
  /**
   * When `true`, the scene dissolves into the editor demo instead of fading out,
   * leaving the demo's matching slide in place of the preview panel.
   */
  handoff?: boolean;
}

export function CompositionScene({ handoff = false }: CompositionSceneProps) {
  const time = useCurrentFrame();
  const reveal = progress(time, 8, 28);
  const exitStart = COMPOSITION_DURATION - HANDOFF_FRAMES;
  const exit = handoff ? progress(time, exitStart, exitStart + HANDOFF_FADE_FRAMES) : 0;
  const shift = progress(time, MOVE_START, MOVE_START + MOVE_FRAMES);
  const railLanded = MOVE_START + MOVE_FRAMES;
  const railIn = progress(time, railLanded, railLanded + RAIL_IN_FRAMES);
  const railReveal = Math.max(0, (time - (railLanded + RAIL_POP_CADENCE)) / RAIL_POP_CADENCE);

  return (
    <AbsoluteFill>
      <CompositionBackdrop handoff={handoff} />
      <SceneContent durationInFrames={COMPOSITION_DURATION} fadeOut={false}>
        <AbsoluteFill
          style={{
            flexDirection: "row",
            alignItems: "center",
            justifyContent: "center",
            gap: COMPOSITION_GAP,
            opacity: 1 - exit,
          }}
        >
          <CodePanel reveal={reveal}>
            <CompositionCode shift={shift} />
          </CodePanel>
          <EditorPreview
            file={PREVIEW_FILE}
            slideIndex={PREVIEW_SLIDE}
            width={PREVIEW_W}
            height={PREVIEW_H}
            railWidth={PREVIEW_RAIL_W}
            padding={PREVIEW_PAD}
            reveal={reveal}
            railIn={railIn}
            railReveal={railReveal}
            style={{ flex: "0 0 auto" }}
          />
        </AbsoluteFill>
      </SceneContent>
    </AbsoluteFill>
  );
}
