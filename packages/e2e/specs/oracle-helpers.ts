/**
 * Node-side image comparison for the ground-truth oracle specs.
 *
 * Screenshots and PowerPoint-exported PNGs are decoded with sharp, resized to
 * common dimensions, and scored with SSIM (structural similarity; 1 = pixel
 * identical). Fuzzy scoring is deliberate: fonts, antialiasing, and text
 * metrics legitimately differ between PowerPoint and the browser, so
 * pixel-exact comparison would never pass.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import sharp from "sharp";
import { ssim } from "ssim.js";

const SPECS_DIR = dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, "$1"));

export const GROUND_TRUTH_DIR = join(SPECS_DIR, "..", "fixtures", "ground-truth");
const BASELINES_DIR = join(SPECS_DIR, "oracle-baselines");

/** Score drops beyond this tolerance fail the oracle spec. */
export const SCORE_TOLERANCE = 0.01;

interface RawImage {
  data: Uint8ClampedArray;
  width: number;
  height: number;
}

async function decodeTo(buffer: Buffer, width: number, height: number): Promise<RawImage> {
  const { data } = await sharp(buffer)
    .resize(width, height, { fit: "fill" })
    .ensureAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  return { data: new Uint8ClampedArray(data), width, height };
}

/** SSIM between a Playwright screenshot and a ground-truth PNG (0..1). */
export async function ssimAgainstGroundTruth(
  screenshot: Buffer,
  groundTruthPath: string,
): Promise<number> {
  const groundTruth = readFileSync(groundTruthPath);
  const meta = await sharp(groundTruth).metadata();
  const width = meta.width ?? 1280;
  const height = meta.height ?? 720;

  const [actual, expected] = await Promise.all([
    decodeTo(screenshot, width, height),
    decodeTo(groundTruth, width, height),
  ]);

  return ssim(actual, expected).mssim;
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

export function readScoreBaseline(deck: string, slide: number, project: string): number | null {
  const path = baselinePath(deck, slide, project);
  if (!existsSync(path)) return null;
  return (JSON.parse(readFileSync(path, "utf8")) as { ssim: number }).ssim;
}

export function writeScoreBaseline(
  deck: string,
  slide: number,
  project: string,
  score: number,
): void {
  mkdirSync(join(BASELINES_DIR, process.platform), { recursive: true });
  writeFileSync(
    baselinePath(deck, slide, project),
    `${JSON.stringify({ ssim: Number(score.toFixed(4)) }, null, 2)}\n`,
  );
}

/** Set ORACLE_UPDATE=1 to (re)record baselines instead of asserting them. */
export const isUpdateMode = process.env.ORACLE_UPDATE === "1";
