import { PANEL_BORDER_WIDTH } from "@/lib/theme";
import { FADE_DURATION } from "@/lib/timing";

export const CONTENT_LEAD = FADE_DURATION;

export const TITLE_DURATION = 60;
export const SPOTLIGHT_DURATION = 62;
export const CAMERA_MOVE = 20;
export const CAMERA_HOLD = SPOTLIGHT_DURATION - CAMERA_MOVE;
export const CAPTION_LEAD = 8;
export const SNAP_BEAT = 47;
export const COMPOSITION_CONTENT_FRAMES = 170;
export const DEMO_CONTENT_FRAMES = 430; // 14.33s * 30fps = 429.9 frames

export const SLIDE_W = 960;
export const SLIDE_H = 540;
export const FOCUS_X = 1290;
export const FOCUS_Y = 540;
export const TEXT_ON_BOARD = "0 2px 14px rgba(14,17,23,.85), 0 0 44px rgba(14,17,23,.75)";
export const VIDEO_W = 1920;
export const VIDEO_H = 1080;
export const HANDOFF_FRAMES = 36;
export const FLUSH_FRAMES = 12;
export const COMPOSITION_GAP = 72;
export const PREVIEW_RAIL_W = 130;
export const PREVIEW_PAD = 26;
export const PREVIEW_W = 980;
export const PREVIEW_CANVAS_W = PREVIEW_W - PANEL_BORDER_WIDTH * 2 - PREVIEW_RAIL_W;
export const PREVIEW_H =
  Math.round(((PREVIEW_CANVAS_W - PREVIEW_PAD * 2) * SLIDE_H) / SLIDE_W) +
  PREVIEW_PAD * 2 +
  PANEL_BORDER_WIDTH * 2;

export function sectionDuration(contentFrames: number) {
  return CONTENT_LEAD + contentFrames + FADE_DURATION;
}

export interface Spotlight {
  file: string;
  slideIndex: number;
  caption: string;
  x: number;
  y: number;
  zoom: number;
}

export const SPOTLIGHTS: Spotlight[] = [
  {
    file: "the-city-after-asphalt-editorial-forest.pptx",
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

export const SHOWCASE_FRAMES = SPOTLIGHTS.length * SPOTLIGHT_DURATION;
export const SPOTLIGHTS_TOTAL = CONTENT_LEAD + SHOWCASE_FRAMES + HANDOFF_FRAMES;

export const SNAPS = [
  "Parse. Render. Edit. Re-export.",
  "Tables, charts, shapes, images.",
  "TypeScript-first. No server required.",
];

export const FEATURES_DURATION = sectionDuration(SNAPS.length * SNAP_BEAT);
export const COMPOSITION_DURATION = sectionDuration(COMPOSITION_CONTENT_FRAMES);
export const CTA_DURATION = CONTENT_LEAD + 90;
export const DEMO_DURATION = sectionDuration(DEMO_CONTENT_FRAMES);

export const RAIL_LINE_ID = "rail";

export const COMPOSITION_LINES: Record<string, string> = {
  "root-open": "<Presentation.Root>",
  [RAIL_LINE_ID]: "  <Presentation.ThumbnailList />",
  "viewport-open": "  <Presentation.Viewport>",
  "slide-open": "    <Presentation.Slide>",
  selection: "      <Presentation.Selection />",
  "slide-close": "    </Presentation.Slide>",
  "viewport-close": "  </Presentation.Viewport>",
  "root-close": "</Presentation.Root>",
};

export const WITHOUT_RAIL = [
  "root-open",
  "viewport-open",
  "slide-open",
  "selection",
  "slide-close",
  "viewport-close",
  "root-close",
];

export const WITH_RAIL = [
  "root-open",
  RAIL_LINE_ID,
  "viewport-open",
  "slide-open",
  "selection",
  "slide-close",
  "viewport-close",
  "root-close",
];

export const COMPOSITION_SOURCE = WITH_RAIL.map((id) => COMPOSITION_LINES[id]).join("\n");

export const CODE_FONT_SIZE = 34;
export const CODE_LINE_H = Math.round(CODE_FONT_SIZE * 1.6);
export const CODE_COLUMNS = Math.max(
  ...Object.values(COMPOSITION_LINES).map((line) => line.length),
);
export const CODE_PANEL_PAD_X = 40;
export const CODE_PANEL_W = CODE_PANEL_PAD_X * 2 + Math.round(CODE_COLUMNS * CODE_FONT_SIZE * 0.6);

export const PREVIEW_FILE = "after-the-needle-drops-mat.pptx";
export const PREVIEW_SLIDE = 0;
export const MOVE_START = 72;
export const MOVE_FRAMES = 26;
export const RAIL_IN_FRAMES = 20;
export const RAIL_POP_CADENCE = 4;

export const EDITOR_DEMO_FILE = "editor-demo.mp4";

export function launchDuration(hasEditorDemo: boolean): number {
  const fades = hasEditorDemo ? 4 : 3;

  return (
    TITLE_DURATION +
    SPOTLIGHTS_TOTAL +
    COMPOSITION_DURATION +
    (hasEditorDemo ? DEMO_DURATION : 0) +
    FEATURES_DURATION +
    CTA_DURATION -
    fades * FADE_DURATION -
    HANDOFF_FRAMES
  );
}

export const LAUNCH_DURATION = launchDuration(false);
