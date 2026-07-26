/**
 * Pixel-level reads of a screenshot.
 *
 * A hairline that looks wrong is usually a coverage problem rather than a
 * colour one: it lands between device pixels and is spread across two of them
 * at partial alpha, which reads as a washed-out line tinted by whatever it sits
 * on. Only the actual channel values across the line tell that apart from a
 * genuinely wrong colour, so the border specs assert on them directly.
 */
import sharp from "sharp";

export interface Pixel {
  r: number;
  g: number;
  b: number;
}

export interface PixelGrid {
  width: number;
  height: number;
  at(x: number, y: number): Pixel;
  /** Pixels down a vertical line, `from`..`to` inclusive. */
  column(x: number, from: number, to: number): Pixel[];
  /** Pixels across a horizontal line, `from`..`to` inclusive. */
  row(y: number, from: number, to: number): Pixel[];
}

/** Decodes a PNG screenshot into an addressable RGB grid. */
export async function readPixels(png: Buffer): Promise<PixelGrid> {
  const { data, info } = await sharp(png).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  const { width, height, channels } = info;

  const at = (x: number, y: number): Pixel => {
    if (x < 0 || y < 0 || x >= width || y >= height) {
      throw new Error(`pixel (${x}, ${y}) is outside the ${width}x${height} image`);
    }
    const offset = (y * width + x) * channels;
    return { r: data[offset], g: data[offset + 1], b: data[offset + 2] };
  };

  const range = (from: number, to: number) =>
    Array.from({ length: to - from + 1 }, (_, i) => from + i);

  return {
    width,
    height,
    at,
    column: (x, from, to) => range(from, to).map((y) => at(x, y)),
    row: (y, from, to) => range(from, to).map((x) => at(x, y)),
  };
}

/** Perceptual-ish distance between two colours, 0 when identical. */
export function colorDistance(pixel: Pixel, hex: string): number {
  const value = Number.parseInt(hex.replace("#", ""), 16);
  return Math.max(
    Math.abs(pixel.r - ((value >> 16) & 0xff)),
    Math.abs(pixel.g - ((value >> 8) & 0xff)),
    Math.abs(pixel.b - (value & 0xff)),
  );
}

/** `rgb(r,g,b)` for assertion messages. */
export function formatPixel({ r, g, b }: Pixel): string {
  return `rgb(${r},${g},${b})`;
}
