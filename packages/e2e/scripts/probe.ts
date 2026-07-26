/**
 * Reads a screenshot back at the pixel level, or magnifies a crop of it.
 *
 * Prints the channel values along a row or column so a washed-out hairline
 * (partial coverage spread over two device pixels) can be told apart from a
 * genuinely wrong border colour:
 *
 *   pnpm probe out/table-borders.pptx-0.png --col 416 --y 180:200
 *   pnpm probe out/table-borders.pptx-0.png --row 300 --x 90:1060
 *   pnpm probe out/table-borders.pptx-0.png --crop 400,180,40,30 --zoom 9
 */
import { mkdir, readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import sharp from "sharp";

import { readPixels } from "../specs/pixels";
import { parseArgs, parseRange } from "./args";

const args = parseArgs(process.argv.slice(2), {
  numbers: ["row", "col", "zoom"],
  strings: ["x", "y", "crop", "out"],
});

const source = args.positional[0];
if (!source) {
  throw new Error("usage: pnpm probe <screenshot.png> [--row N | --col N | --crop x,y,w,h]");
}
const png = await readFile(resolve(source));
const pixels = await readPixels(png);
console.log(`${source} ${pixels.width}x${pixels.height}`);

const { row, col, zoom } = args.numbers;

if (row !== undefined || col !== undefined) {
  const alongX = row !== undefined;
  const [from, to] = alongX
    ? parseRange(args.strings.x, 0, pixels.width - 1)
    : parseRange(args.strings.y, 0, pixels.height - 1);
  const samples = alongX ? pixels.row(row, from, to) : pixels.column(col as number, from, to);

  console.log(alongX ? `row ${row}` : `col ${col}`);
  samples.forEach((pixel, index) => {
    const at = String(from + index).padStart(5);
    const channels = [pixel.r, pixel.g, pixel.b].map((v) => String(v).padStart(3)).join(",");
    console.log(`  ${alongX ? "x" : "y"}=${at} : ${channels}`);
  });
}

if (args.strings.crop) {
  const [x, y, width, height] = args.strings.crop.split(",").map(Number);
  const factor = zoom ?? 1;
  const out = resolve(args.strings.out ?? "out/crop.png");
  await mkdir(dirname(out), { recursive: true });
  await sharp(png)
    .extract({ left: x, top: y, width, height })
    // Nearest-neighbour: a magnified crop has to show the pixels as captured.
    .resize(width * factor, height * factor, { kernel: "nearest" })
    .toFile(out);
  console.log(`wrote ${out} (${width}x${height} @${factor}x)`);
}
