import * as React from "react";
import type { CSSProperties } from "react";

import type { PresentationStore } from "@diceui/pptx";
import { Presentation } from "@diceui/pptx";
import {
  AbsoluteFill,
  continueRender,
  delayRender,
  Easing,
  interpolate,
  staticFile,
  useCurrentFrame,
} from "remotion";

import type { CursorKeyframe } from "@/components/cursor";
import { Cursor } from "@/components/cursor";
import { SceneBg } from "@/components/scene-bg";
import { geistSans } from "@/fonts";

const FONT_VARS = {
  "--font-geist-sans": geistSans,
} as CSSProperties;

const VIEWPORT_W = 1440;
const VIEWPORT_H = 810;
const VIEWPORT_LEFT = (1920 - VIEWPORT_W) / 2;
const VIEWPORT_TOP = 100;

// Scale factor: cursor coords are in composition space (1920x1080),
// edits need slide-space coords. The viewport is 1440x810 placed at
// VIEWPORT_LEFT, VIEWPORT_TOP. The slide itself is 1270x714 (editorial-forest)
// but autoFit scales it to fill 1440x810.
const SLIDE_W = 1270;
const SLIDE_H = 714;
const SCALE_X = SLIDE_W / VIEWPORT_W;
const SCALE_Y = SLIDE_H / VIEWPORT_H;

function toSlideX(screenX: number) {
  return (screenX - VIEWPORT_LEFT) * SCALE_X;
}
function toSlideY(screenY: number) {
  return (screenY - VIEWPORT_TOP) * SCALE_Y;
}

// --- Phase 1: Drag (frames 10–90) ---
// Cursor moves to a shape and drags it to the right.
const DRAG_START_SCREEN = { x: 500, y: 380 };
const DRAG_END_SCREEN = { x: 900, y: 320 };

// --- Phase 2: Resize (frames 100–170) ---
const RESIZE_TARGET_SCREEN = { x: 1100, y: 500 };
const RESIZE_END_SCREEN = { x: 1250, y: 580 };

// --- Phase 3: Text edit (frames 180–235) ---
const TEXT_TARGET_SCREEN = { x: 600, y: 280 };

const CURSOR_KEYFRAMES: CursorKeyframe[] = [
  // Appear
  { frame: 10, x: 960, y: 600 },
  // Move to drag target
  { frame: 30, x: DRAG_START_SCREEN.x, y: DRAG_START_SCREEN.y },
  // Press and drag
  { frame: 40, x: DRAG_START_SCREEN.x, y: DRAG_START_SCREEN.y, pressed: true },
  { frame: 80, x: DRAG_END_SCREEN.x, y: DRAG_END_SCREEN.y, pressed: true },
  { frame: 85, x: DRAG_END_SCREEN.x, y: DRAG_END_SCREEN.y },
  // Move to resize handle
  { frame: 105, x: RESIZE_TARGET_SCREEN.x, y: RESIZE_TARGET_SCREEN.y },
  // Press and resize
  { frame: 115, x: RESIZE_TARGET_SCREEN.x, y: RESIZE_TARGET_SCREEN.y, pressed: true },
  { frame: 155, x: RESIZE_END_SCREEN.x, y: RESIZE_END_SCREEN.y, pressed: true },
  { frame: 160, x: RESIZE_END_SCREEN.x, y: RESIZE_END_SCREEN.y },
  // Move to text
  { frame: 185, x: TEXT_TARGET_SCREEN.x, y: TEXT_TARGET_SCREEN.y },
  { frame: 192, x: TEXT_TARGET_SCREEN.x, y: TEXT_TARGET_SCREEN.y, pressed: true },
  { frame: 195, x: TEXT_TARGET_SCREEN.x, y: TEXT_TARGET_SCREEN.y },
  // Typing
  { frame: 235, x: TEXT_TARGET_SCREEN.x, y: TEXT_TARGET_SCREEN.y },
];

const ACTION_LABELS: { start: number; end: number; text: string }[] = [
  { start: 30, end: 90, text: "Drag shapes" },
  { start: 100, end: 165, text: "Resize elements" },
  { start: 180, end: 240, text: "Edit text inline" },
];

export const INTERACTION_FRAMES = 240;

interface NodeInfo {
  id: string;
  nodeType: string;
  position: { x: number; y: number };
  size: { w: number; h: number };
  hasText: boolean;
}

export function InteractionScene() {
  const frame = useCurrentFrame();
  const [data, setData] = React.useState<ArrayBuffer | null>(null);
  const storeRef = React.useRef<PresentationStore | null>(null);
  const nodesRef = React.useRef<NodeInfo[]>([]);
  const slideIdRef = React.useRef<string | null>(null);
  const lastEditFrame = React.useRef(-1);
  const [handle] = React.useState(() => delayRender("load interaction deck"));

  React.useEffect(() => {
    let cancelled = false;
    fetch(staticFile("editable.pptx"))
      .then((r) => r.arrayBuffer())
      .then((buf) => {
        if (!cancelled) setData(buf);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  // Drive edits based on frame
  React.useEffect(() => {
    const store = storeRef.current;
    if (!store) return;
    const nodes = nodesRef.current;
    const slideId = slideIdRef.current;
    if (!slideId || nodes.length === 0) return;

    // Only apply edits going forward, once per frame
    if (frame <= lastEditFrame.current) return;
    lastEditFrame.current = frame;

    // Phase 1: Drag — move node[0] from its original position toward the right
    const dragNode = nodes[0];
    if (dragNode && frame >= 40 && frame <= 80) {
      const progress = (frame - 40) / 40;
      const dx = (toSlideX(DRAG_END_SCREEN.x) - toSlideX(DRAG_START_SCREEN.x)) * progress;
      const dy = (toSlideY(DRAG_END_SCREEN.y) - toSlideY(DRAG_START_SCREEN.y)) * progress;
      store.edit({
        type: "setNodeTransform",
        slideId,
        nodeId: dragNode.id,
        position: {
          x: dragNode.position.x + dx,
          y: dragNode.position.y + dy,
        },
      });
    }

    // Phase 2: Resize — grow node[1]
    const resizeNode = nodes.length > 1 ? nodes[1] : undefined;
    if (resizeNode && frame >= 115 && frame <= 155) {
      const progress = (frame - 115) / 40;
      const extraW = 80 * progress;
      const extraH = 50 * progress;
      store.edit({
        type: "setNodeTransform",
        slideId,
        nodeId: resizeNode.id,
        size: {
          w: resizeNode.size.w + extraW,
          h: resizeNode.size.h + extraH,
        },
      });
    }

    // Phase 3: Text edit — progressively type into the first text node
    const textNode = nodes.find((n) => n.hasText);
    if (textNode && frame >= 200 && frame <= 230) {
      const fullText = "PowerPoint in the Browser";
      const charsTyped = Math.floor(((frame - 200) / 30) * fullText.length);
      const text = fullText.substring(0, charsTyped);
      store.edit({
        type: "setTextRun",
        slideId,
        nodeId: textNode.id,
        paragraphIndex: 0,
        runIndex: 0,
        text,
      });
    }
  }, [frame]);

  const containerOpacity = interpolate(frame, [0, 15], [0, 1], {
    extrapolateRight: "clamp",
    easing: Easing.bezier(0.16, 1, 0.3, 1),
  });

  const activeLabel = ACTION_LABELS.find((l) => frame >= l.start && frame < l.end);
  const labelOpacity = activeLabel
    ? interpolate(frame, [activeLabel.start, activeLabel.start + 12], [0, 1], {
        extrapolateLeft: "clamp",
        extrapolateRight: "clamp",
      })
    : 0;

  return (
    <AbsoluteFill style={FONT_VARS}>
      <SceneBg />
      <div
        style={{
          position: "absolute",
          left: VIEWPORT_LEFT,
          top: VIEWPORT_TOP,
          opacity: containerOpacity,
          borderRadius: 14,
          overflow: "hidden",
          boxShadow: "0 40px 100px rgba(0,0,0,0.6)",
        }}
      >
        {data && (
          <Presentation.Root
            file={data}
            readOnly={false}
            onLoad={(store) => {
              storeRef.current = store;
              // Discover nodes on the first slide
              const slide = store.getActiveSlide();
              if (slide) {
                slideIdRef.current = slide.id;
                nodesRef.current = slide.nodes.map((n) => ({
                  id: n.id,
                  nodeType: n.nodeType,
                  position: { ...n.position },
                  size: { ...n.size },
                  hasText:
                    n.nodeType === "shape" &&
                    (n as any).textBody?.paragraphs?.length > 0 &&
                    (n as any).textBody?.paragraphs?.[0]?.runs?.length > 0,
                }));
              }
              continueRender(handle);
            }}
          >
            <Presentation.Viewport
              autoFit
              style={{ width: VIEWPORT_W, height: VIEWPORT_H, overflow: "hidden" }}
            >
              <Presentation.Slide />
            </Presentation.Viewport>
          </Presentation.Root>
        )}
      </div>

      <Cursor keyframes={CURSOR_KEYFRAMES} appearFrame={10} />

      {activeLabel && (
        <div
          style={{
            position: "absolute",
            bottom: 80,
            left: 0,
            right: 0,
            textAlign: "center",
            fontSize: 24,
            fontFamily: geistSans,
            fontWeight: 600,
            color: "#3b82f6",
            letterSpacing: "0.06em",
            textTransform: "uppercase",
            opacity: labelOpacity,
          }}
        >
          {activeLabel.text}
        </div>
      )}
    </AbsoluteFill>
  );
}
