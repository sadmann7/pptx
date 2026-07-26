/**
 * Node-side image comparison for the ground-truth oracle specs.
 *
 * Screenshots and PowerPoint-exported PNGs are decoded with sharp, resized to
 * common dimensions, and scored with SSIM (structural similarity; 1 = pixel
 * identical). Fuzzy scoring is deliberate: fonts, antialiasing, and text
 * metrics legitimately differ between PowerPoint and the browser, so a
 * pixel-counting comparison would never pass. SSIM's local windows tolerate a
 * glyph landing a pixel off while still catching structural breakage.
 *
 * Three numbers are recorded per slide, because SSIM alone cannot gate this.
 *
 * - `overall`: SSIM over the whole image. Sensitive to systematic drift, blunt
 *   about local damage, since most of a slide is background that matches
 *   perfectly and dilutes the mean.
 * - `worstTile`: the lowest SSIM over an 8x6 grid. Localises damage the mean
 *   absorbs, e.g. a table drawn a few pixels off.
 * - `ink`: per region of a 4x3 grid, how far the painted share of that region
 *   is from PowerPoint's, as a proportion of PowerPoint's.
 *
 * The ink measure exists because SSIM scores *blank* higher than *misplaced*:
 * two images with text half a line apart are locally anti-correlated (near
 * zero), while text against blank background merely lacks correlation (~0.5).
 * A renderer change that collapsed a paragraph into unreadable 8px text
 * therefore *improved* both SSIM numbers while destroying the slide.
 *
 * Its shape was settled by fault injection across six exported slides — hide a
 * text block, shrink one to half, hide the largest graphic, shift the slide
 * 8px, recolour the largest shape — which `pnpm faults` re-runs, and which any
 * change to these tolerances or the grid should be re-checked against:
 *
 * - Proportional, not absolute. Text paints only a few percent of a region, so
 *   an absolute coverage gap sized to one slide missed a whole text block
 *   vanishing on a sparser one.
 * - Per region against that region's own baseline, not the worst region. A max
 *   is masked by whichever region permanently sits worst, which swallowed two
 *   of the injected faults.
 * - 4x3, not 8x6. Text drifting a few pixels across a tile edge moves a small
 *   tile's ink as much as content vanishing does; at 4x3 the faults stood
 *   0.47-0.92 above control while re-runs moved 0.000.
 *
 * Known blind spots, from that same run: on an ink-dense slide, losing a
 * *small* element barely moves any region (0.003 on one of the six) and SSIM
 * does not catch it either; hiding a large background graphic can also slip
 * under every tolerance. Positional faults are the SSIM pair's job, not this
 * one's.
 */

import ssim from "@blazediff/ssim/ssim";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import sharp from "sharp";

const SPECS_DIR = dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, "$1"));

export const GROUND_TRUTH_DIR = join(SPECS_DIR, "..", "fixtures", "ground-truth");
const BASELINES_DIR = join(SPECS_DIR, "oracle-baselines");

/** Whole-image score drops beyond this tolerance fail the oracle spec. */
export const SCORE_TOLERANCE = 0.01;

/**
 * Worst-tile drops beyond this tolerance fail. Looser than the whole-image
 * tolerance because a single tile is a twentieth of the pixels, so the same
 * antialiasing difference moves it further; still far tighter than what the
 * mean can detect, since a tile is mostly the thing that broke.
 */
export const TILE_TOLERANCE = 0.03;

/**
 * How much further one region's ink may drift from PowerPoint's than its
 * baseline records. Re-running the same slide on the same machine moves this
 * 0.000, and the injected content-loss faults moved it 0.47 and up, so this
 * sits an order of magnitude below what it must catch while leaving room for
 * the font substitution that costs Linux ~0.015 SSIM. Recording baselines on a
 * new platform is the check on that headroom.
 */
export const INK_TOLERANCE = 0.15;

/**
 * Absolute floor, independent of baselines. Catches catastrophic breakage
 * (blank slide, missing chart) even on platforms with no recorded baseline
 * or with a badly-recorded one.
 *
 * The exported decks set the bar here: their worst slides score ~0.74 on
 * Windows, where line spacing accumulates a few pixels of drift down a
 * paragraph and the charts choose their own axis ticks — both of which SSIM
 * penalises across the whole image even though the slide reads correctly.
 * Linux scores ~0.015 lower again (0.06 on the worst slide) for want of the
 * decks' fonts, bottoming out at 0.705, so 0.65 is as high as the floor can
 * sit without turning runner variance into a failure; raising it means fixing
 * those two gaps first. Per-slide regressions are caught by the baseline
 * comparison, not by this floor.
 */
export const SCORE_FLOOR = 0.65;

/** Decodes to the RGBA bytes the SSIM implementation expects. */
async function decodeTo(buffer: Buffer, width: number, height: number): Promise<Uint8ClampedArray> {
  const data = await sharp(buffer)
    .resize(width, height, { fit: "fill" })
    .ensureAlpha()
    .raw()
    .toBuffer();
  return new Uint8ClampedArray(data);
}

/**
 * Tile grid the worst-tile score is computed over. 8x6 puts a tile at roughly
 * a sixth of the slide's height, which is large enough that text antialiasing
 * averages out and small enough that one broken region cannot hide behind the
 * background around it.
 */
const TILE_COLUMNS = 8;
const TILE_ROWS = 6;

export interface OracleScore {
  /** SSIM over the whole image. */
  overall: number;
  /** The lowest per-tile SSIM, which localises damage the mean absorbs. */
  worstTile: number;
  /** Per 4x3 region, its painted share's proportional gap from PowerPoint's. */
  ink: number[];
}

/** A pixel this far from the slide's background tone counts as painted. */
const INK_THRESHOLD = 24;

/** Grid the ink regions are cut on. Coarse on purpose — see the file header. */
const INK_COLUMNS = 4;
const INK_ROWS = 3;

/**
 * Share of a region PowerPoint must paint before its ink is compared. Below
 * this the region is background holding a stray glyph edge, where a
 * proportional comparison reads antialiasing as total loss. Content appearing
 * in a region PowerPoint leaves empty is the SSIM pair's to catch.
 */
const MIN_REGION_INK = 0.02;

/** Share of painted pixels per region, against the image's own background tone. */
async function inkPerRegion(buffer: Buffer, width: number, height: number): Promise<number[]> {
  const pixels = await sharp(buffer)
    .resize(width, height, { fit: "fill" })
    .greyscale()
    .raw()
    .toBuffer();

  // Slides are mostly background, so its tone is the median sample. Assuming
  // white would read a dark-themed slide as entirely painted.
  const background = Uint8Array.prototype.slice.call(pixels).sort()[Math.floor(pixels.length / 2)];

  const regionWidth = Math.floor(width / INK_COLUMNS);
  const regionHeight = Math.floor(height / INK_ROWS);
  const regions: number[] = [];
  for (let row = 0; row < INK_ROWS; row++) {
    for (let column = 0; column < INK_COLUMNS; column++) {
      let painted = 0;
      for (let y = 0; y < regionHeight; y++) {
        const rowStart = (row * regionHeight + y) * width + column * regionWidth;
        for (let x = 0; x < regionWidth; x++) {
          if (Math.abs(pixels[rowStart + x] - background) > INK_THRESHOLD) painted++;
        }
      }
      regions.push(painted / (regionWidth * regionHeight));
    }
  }
  return regions;
}

/** Copies one tile's RGBA bytes out of a full-image buffer. */
function cropTile(
  source: Uint8ClampedArray,
  imageWidth: number,
  x: number,
  y: number,
  width: number,
  height: number,
): Uint8ClampedArray {
  const tile = new Uint8ClampedArray(width * height * 4);
  for (let row = 0; row < height; row++) {
    const start = ((y + row) * imageWidth + x) * 4;
    tile.set(source.subarray(start, start + width * 4), row * width * 4);
  }
  return tile;
}

/** Scores a Playwright screenshot against a ground-truth PNG (0..1 each). */
export async function scoreAgainstGroundTruth(
  screenshot: Buffer,
  groundTruthPath: string,
): Promise<OracleScore> {
  const groundTruth = readFileSync(groundTruthPath);
  const meta = await sharp(groundTruth).metadata();
  const width = meta.width ?? 1280;
  const height = meta.height ?? 720;

  const [actual, expected] = await Promise.all([
    decodeTo(screenshot, width, height),
    decodeTo(groundTruth, width, height),
  ]);

  const tileWidth = Math.floor(width / TILE_COLUMNS);
  const tileHeight = Math.floor(height / TILE_ROWS);
  let worstTile = 1;
  for (let row = 0; row < TILE_ROWS; row++) {
    for (let column = 0; column < TILE_COLUMNS; column++) {
      const x = column * tileWidth;
      const y = row * tileHeight;
      const score = ssim(
        cropTile(actual, width, x, y, tileWidth, tileHeight),
        cropTile(expected, width, x, y, tileWidth, tileHeight),
        undefined,
        tileWidth,
        tileHeight,
      );
      if (score < worstTile) worstTile = score;
    }
  }

  const [actualInk, expectedInk] = await Promise.all([
    inkPerRegion(screenshot, width, height),
    inkPerRegion(groundTruth, width, height),
  ]);
  const ink = expectedInk.map((expected, index) =>
    expected < MIN_REGION_INK ? 0 : Math.abs(actualInk[index] - expected) / expected,
  );

  return { overall: ssim(actual, expected, undefined, width, height), worstTile, ink };
}

// ---------------------------------------------------------------------------
// Score baselines: one JSON file per (deck, slide, browser, OS) so parallel
// workers never write the same file, and diffs stay reviewable.
//
// Platform is part of the key because text rendering differs per OS (Linux
// lacks Calibri and antialiases differently), shifting SSIM by ~0.01-0.03
// even when the renderer is unchanged.
// ---------------------------------------------------------------------------

function baselinePath(deck: string, slide: number, project: string): string {
  return join(BASELINES_DIR, process.platform, `${deck}-${slide}-${project}.json`);
}

export function readScoreBaseline(
  deck: string,
  slide: number,
  project: string,
): OracleScore | null {
  const path = baselinePath(deck, slide, project);
  if (!existsSync(path)) return null;
  const recorded = JSON.parse(readFileSync(path, "utf8")) as Partial<OracleScore>;
  if (
    recorded.overall === undefined ||
    recorded.worstTile === undefined ||
    recorded.ink?.length !== INK_COLUMNS * INK_ROWS
  ) {
    // Recorded before the current metrics existed; re-record to assert again.
    return null;
  }
  return { overall: recorded.overall, worstTile: recorded.worstTile, ink: recorded.ink };
}

export function writeScoreBaseline(
  deck: string,
  slide: number,
  project: string,
  score: OracleScore,
): void {
  mkdirSync(join(BASELINES_DIR, process.platform), { recursive: true });
  const rounded = {
    overall: Number(score.overall.toFixed(4)),
    worstTile: Number(score.worstTile.toFixed(4)),
    ink: score.ink.map((drift) => Number(drift.toFixed(4))),
  } satisfies OracleScore;
  writeFileSync(baselinePath(deck, slide, project), `${JSON.stringify(rounded, null, 2)}\n`);
}

/** Set ORACLE_UPDATE=1 to (re)record baselines instead of asserting them. */
export const isUpdateMode = process.env.ORACLE_UPDATE === "1";
