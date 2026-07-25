import { readFile, mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { chromium } from "playwright";

import { parseArgs, parseRange } from "./args.mjs";

/**
 * Read a screenshot back at the pixel level.
 *
 * A hairline that looks wrong is usually a coverage problem rather than a
 * colour one: it lands between device pixels and gets spread across two of them
 * at partial alpha, which reads as a washed-out line tinted by whatever it sits
 * on. Printing the actual channel values across the line is the only reliable
 * way to tell that apart from a genuinely wrong colour.
 *
 *   node probe.mjs out/slide7.png --col 740 --y 470:490
 *   node probe.mjs out/slide7.png --row 480 --x 330:1400
 *   node probe.mjs out/slide7.png --crop 700,470,90,40 --zoom 9 --out out/corner.png
 *
 * Decoding happens in a headless page (the browser is already a dependency
 * here), so this works the same on every platform.
 */
const args = parseArgs(process.argv.slice(2), {
  numbers: ["row", "col", "zoom"],
  strings: ["x", "y", "crop", "out"],
});

const source = args._?.[0];
if (!source) throw new Error("usage: node probe.mjs <screenshot.png> [--row N | --col N | --crop]");
const dataUrl = `data:image/png;base64,${(await readFile(resolve(source))).toString("base64")}`;

const browser = await chromium.launch();
const page = await browser.newPage();
await page.setContent("<canvas id=c></canvas>");
const size = await page.evaluate(async (url) => {
  const image = new Image();
  image.src = url;
  await image.decode();
  const canvas = document.getElementById("c");
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;
  canvas.getContext("2d").drawImage(image, 0, 0);
  globalThis.__image = image;
  return { width: image.naturalWidth, height: image.naturalHeight };
}, dataUrl);
console.log(`${source} ${size.width}x${size.height}`);

if (args.row !== undefined || args.col !== undefined) {
  const along =
    args.row !== undefined
      ? parseRange(args.x ?? `0:${size.width - 1}`)
      : parseRange(args.y ?? `0:${size.height - 1}`);
  const samples = await page.evaluate(
    ({ row, col, from, to }) => {
      const context = document.getElementById("c").getContext("2d");
      const out = [];
      for (let i = from; i <= to; i++) {
        const x = row !== undefined ? i : col;
        const y = row !== undefined ? row : i;
        const [r, g, b] = context.getImageData(x, y, 1, 1).data;
        out.push({ i, r, g, b });
      }
      return out;
    },
    { row: args.row, col: args.col, from: along.from, to: along.to },
  );
  const axis = args.row !== undefined ? "x" : "y";
  const fixed = args.row !== undefined ? `row ${args.row}` : `col ${args.col}`;
  console.log(fixed);
  for (const { i, r, g, b } of samples) {
    console.log(
      `  ${axis}=${String(i).padStart(5)} : ${[r, g, b].map((v) => String(v).padStart(3)).join(",")}`,
    );
  }
}

if (args.crop) {
  const [x, y, width, height] = args.crop.split(",").map(Number);
  const zoom = args.zoom ?? 1;
  const out = resolve(args.out ?? "out/crop.png");
  const cropped = await page.evaluate(
    ({ x, y, width, height, zoom }) => {
      const canvas = document.createElement("canvas");
      canvas.width = width * zoom;
      canvas.height = height * zoom;
      const context = canvas.getContext("2d");
      // Nearest-neighbour: a magnified crop must show the pixels as captured.
      context.imageSmoothingEnabled = false;
      context.drawImage(globalThis.__image, x, y, width, height, 0, 0, canvas.width, canvas.height);
      return canvas.toDataURL("image/png");
    },
    { x, y, width, height, zoom },
  );
  await mkdir(dirname(out), { recursive: true });
  await writeFile(out, Buffer.from(cropped.split(",")[1], "base64"));
  console.log(`wrote ${out} (${width}x${height} @${zoom}x)`);
}

await browser.close();
