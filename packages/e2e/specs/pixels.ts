/**
 * A hairline between device pixels spreads across two at partial alpha, and
 * only the channel values across it tell that apart from a wrong colour.
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
  column(x: number, from: number, to: number): Pixel[];
  row(y: number, from: number, to: number): Pixel[];
}

export async function readPixels(png: Buffer): Promise<PixelGrid> {
  const { data, info } = await sharp(png)
    .removeAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  const { width, height, channels } = info;

  const at = (x: number, y: number): Pixel => {
    if (x < 0 || y < 0 || x >= width || y >= height) {
      throw new Error(
        `pixel (${x}, ${y}) is outside the ${width}x${height} image`,
      );
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

export function colorDistance(pixel: Pixel, hex: string): number {
  const value = Number.parseInt(hex.replace("#", ""), 16);
  return Math.max(
    Math.abs(pixel.r - ((value >> 16) & 0xff)),
    Math.abs(pixel.g - ((value >> 8) & 0xff)),
    Math.abs(pixel.b - (value & 0xff)),
  );
}

export function formatPixel({ r, g, b }: Pixel): string {
  return `rgb(${r},${g},${b})`;
}
