/**
 * Node-side image comparison for the oracle specs.
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
 * Its shape was settled by fault injection across six exported slides (hide a
 * text block, shrink one to half, hide the largest graphic, shift the slide
 * 8px, recolour the largest shape), which `pnpm faults` re-runs, and which any
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

export const POWERPOINT_DIR = join(SPECS_DIR, "..", "fixtures", "powerpoint");
const BASELINES_DIR = join(SPECS_DIR, "oracle-baselines");

export const SCORE_TOLERANCE = 0.01;

// Looser than SCORE_TOLERANCE because a tile is a twentieth of the pixels, so the
// same antialiasing difference moves it further.
export const TILE_TOLERANCE = 0.03;

// Reruns move ink 0.000 and injected content-loss faults 0.47 and up, so this
// sits well below what it must catch with room for Linux's font substitution.
export const INK_TOLERANCE = 0.15;

// Catches a blank slide or missing chart regardless of baselines. The exported
// decks' worst slides score ~0.74 on Windows and 0.705 on Linux without the
// fonts, so 0.65 is as high as it can sit without runner variance failing it.
export const SCORE_FLOOR = 0.65;

async function decodeTo(buffer: Buffer, width: number, height: number): Promise<Uint8ClampedArray> {
  const data = await sharp(buffer)
    .resize(width, height, { fit: "fill" })
    .ensureAlpha()
    .raw()
    .toBuffer();
  return new Uint8ClampedArray(data);
}

// A tile a sixth of the slide high is large enough for text antialiasing to
// average out and small enough that a broken region cannot hide in background.
const TILE_COLUMNS = 8;
const TILE_ROWS = 6;

export interface OracleScore {
  overall: number;
  worstTile: number;
  ink: number[];
}

const INK_THRESHOLD = 24;

const INK_COLUMNS = 4;
const INK_ROWS = 3;

// Below this share a region is background holding a stray glyph edge, where a
// proportional comparison reads antialiasing as total loss.
const MIN_REGION_INK = 0.02;

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

export async function scoreAgainstPowerPoint(
  screenshot: Buffer,
  powerPointPath: string,
): Promise<OracleScore> {
  const powerPointPng = readFileSync(powerPointPath);
  const meta = await sharp(powerPointPng).metadata();
  const width = meta.width ?? 1280;
  const height = meta.height ?? 720;

  const [actual, expected] = await Promise.all([
    decodeTo(screenshot, width, height),
    decodeTo(powerPointPng, width, height),
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
    inkPerRegion(powerPointPng, width, height),
  ]);
  const ink = expectedInk.map((expected, index) =>
    expected < MIN_REGION_INK ? 0 : Math.abs(actualInk[index] - expected) / expected,
  );

  return { overall: ssim(actual, expected, undefined, width, height), worstTile, ink };
}

// One file per deck, slide, browser and OS, so parallel workers never write the
// same file. Text renders differently per OS (Linux lacks Calibri), moving SSIM
// by ~0.01-0.03.
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
    // A baseline missing a current metric asserts nothing until re-recorded.
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

export const isUpdateMode = process.env.ORACLE_UPDATE === "1";
