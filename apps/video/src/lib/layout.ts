import {
  CODE_PANEL_W,
  COMPOSITION_GAP,
  DEMO_SLIDE,
  PREVIEW_CANVAS_W,
  PREVIEW_H,
  PREVIEW_PAD,
  PREVIEW_RAIL_W,
  PREVIEW_W,
  SLIDE_H,
  SLIDE_W,
  VIDEO_H,
  VIDEO_W,
} from "@/lib/constants";
import { PANEL_BORDER_WIDTH } from "@/lib/theme";

export interface SlideRect {
  left: number;
  top: number;
  width: number;
  height: number;
  /** Slide scale relative to `SLIDE_W` by `SLIDE_H`. */
  zoom: number;
}

export function rectCenter(rect: SlideRect) {
  return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
}

/**
 * Rect of the slide inside the composition scene's preview panel.
 *
 * `railIn` mirrors `EditorPreview`'s rail progress: the canvas sits half a rail
 * further left until the thumbnail rail lands.
 */
export function previewSlideRect(railIn: number): SlideRect {
  const rowWidth = CODE_PANEL_W + COMPOSITION_GAP + PREVIEW_W;
  const previewLeft = (VIDEO_W - rowWidth) / 2 + CODE_PANEL_W + COMPOSITION_GAP;
  const previewTop = (VIDEO_H - PREVIEW_H) / 2;
  const canvasHeight = PREVIEW_H - PANEL_BORDER_WIDTH * 2;
  const zoom = Math.min(
    (PREVIEW_CANVAS_W - PREVIEW_PAD * 2) / SLIDE_W,
    (canvasHeight - PREVIEW_PAD * 2) / SLIDE_H,
  );
  const width = SLIDE_W * zoom;
  const height = SLIDE_H * zoom;

  return {
    width,
    height,
    zoom,
    left:
      previewLeft +
      PANEL_BORDER_WIDTH +
      (PREVIEW_RAIL_W * (1 + railIn)) / 2 +
      (PREVIEW_CANVAS_W - width) / 2,
    top: previewTop + PANEL_BORDER_WIDTH + (canvasHeight - height) / 2,
  };
}

/** Rect of the slide inside the editor demo recording, in video pixels. */
export function demoSlideRect(): SlideRect {
  const zoom = DEMO_SLIDE.width / SLIDE_W;

  return {
    left: DEMO_SLIDE.left,
    top: DEMO_SLIDE.top,
    width: DEMO_SLIDE.width,
    height: SLIDE_H * zoom,
    zoom,
  };
}
