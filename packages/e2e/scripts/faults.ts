import { chromium } from "@playwright/test";
/**
 * Measures what the oracle's three numbers can actually catch, by injecting
 * known faults into a rendered slide and scoring the damaged render against
 * PowerPoint's export the way oracle.spec.ts does.
 *
 * Every fault is compared against a control render of the same slide, so a row
 * reads "this much of the fault showed up in each number". A fault no measure
 * reacts to is a hole in the gate; a control run that trips anything means a
 * tolerance is below the noise floor.
 *
 * This is what settled the ink measure's shape (see oracle.ts) and it should be
 * re-run whenever those tolerances or the grid change, since a gate tuned only
 * against the bug that motivated it proves nothing.
 *
 * Start the harness first (`pnpm harness`), then:
 *   pnpm faults
 *   pnpm faults --slide adventure-club-pin-and-paper:4 --fault hide-text
 */
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import sharp from "sharp";

import {
  POWERPOINT_DIR,
  INK_TOLERANCE,
  type OracleScore,
  SCORE_TOLERANCE,
  scoreAgainstPowerPoint,
  TILE_TOLERANCE,
} from "../specs/oracle";
import { parseArgs } from "./args";

/**
 * Faults chosen to span the ways a render breaks: content lost, content
 * shrunk, a graphic lost, everything shifted, and a fill recoloured. The first
 * two are the classes SSIM is known to score *better* on than a correct
 * render, which is the whole reason a coverage measure exists.
 */
const FAULTS = [
  "control",
  "control-2",
  "hide-text",
  "shrink-text",
  "drop-graphic",
  "shift-8px",
  "recolor",
] as const;

/**
 * Slides the tolerances were sized against: text-heavy, chart, photo-backed,
 * table, textured, and gradient-heavy, so no single measure can look good by
 * suiting one kind of slide.
 */
const DEFAULT_SLIDES = [
  "the-good-room-soft-editorial:1",
  "geometry-of-attention-cartesian:5",
  "make-something-strange-creative-mode:0",
  "adventure-club-pin-and-paper:4",
  "internet-with-texture-broadside:2",
  "pocket-machines-sakura-chroma:3",
];

const args = parseArgs(process.argv.slice(2), {
  numbers: ["port", "width", "height"],
  lists: ["slide", "fault"],
});

const port = args.numbers.port ?? 5000;
const requestedFaults = args.lists.fault?.length ? args.lists.fault : [...FAULTS];
const faults = FAULTS.filter(
  (fault) => fault.startsWith("control") || requestedFaults.includes(fault),
);
const slides = (args.lists.slide?.length ? args.lists.slide : DEFAULT_SLIDES).map((entry) => {
  const [deck, index] = entry.split(":");
  if (!deck) throw new Error(`--slide expects "deck:index", got "${entry}"`);
  return { deck, slide: Number(index ?? 0) };
});

/** Worst movement of any one region away from its control value. */
function worstRegionMoved(score: OracleScore, control: OracleScore): number {
  return score.ink.reduce(
    (worst, drift, region) => Math.max(worst, drift - (control.ink[region] ?? 0)),
    0,
  );
}

/** Which measures a fault moved past the tolerances oracle.spec.ts asserts. */
function caughtBy(score: OracleScore, control: OracleScore): string {
  const caught: string[] = [];
  if (control.overall - score.overall > SCORE_TOLERANCE) caught.push("overall");
  if (control.worstTile - score.worstTile > TILE_TOLERANCE) caught.push("tile");
  if (worstRegionMoved(score, control) > INK_TOLERANCE) caught.push("ink");
  return caught.join("+");
}

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: args.numbers.width ?? 1400, height: args.numbers.height ?? 900 },
});
page.on("pageerror", (error) => console.error("[pageerror]", error.message));

for (const { deck, slide } of slides) {
  const powerPointPath = join(POWERPOINT_DIR, deck, `slide-${slide}.png`);
  if (!existsSync(powerPointPath)) {
    throw new Error(`Missing PowerPoint export ${powerPointPath}. Run "pnpm oracle:export".`);
  }
  const { width = 0, height = 0 } = await sharp(readFileSync(powerPointPath)).metadata();
  let control: OracleScore | undefined;
  let controlShot: Buffer | undefined;

  console.log(`\n${deck} slide ${slide + 1} (${width}x${height})`);
  for (const fault of faults) {
    const url = `http://localhost:${port}/?file=${encodeURIComponent(`${deck}.pptx`)}&slide=${slide}`;
    await page.goto(url, { waitUntil: "load" });
    await page.waitForFunction(
      () => window.__renderDone === true || window.__renderError !== undefined,
      null,
      { timeout: 60_000 },
    );
    const renderError = await page.evaluate(() => window.__renderError);
    if (renderError) throw new Error(renderError);
    // Fonts land after first paint and move every measure, so let them settle.
    await page.waitForFunction(() => document.fonts.status === "loaded", null, { timeout: 30_000 });

    const applied = await page.evaluate((kind) => {
      // Everything stays inline: named functions do not survive the
      // transpiler's rewrite when the body is serialized into the page.
      const root = document.querySelector<HTMLElement>("#slide-container");
      if (!root) throw new Error("no #slide-container");
      if (kind.startsWith("control")) return true;

      const all = [...root.querySelectorAll<HTMLElement>("*")];
      const areas = new Map<Element, number>();
      for (const element of all) {
        const rect = element.getBoundingClientRect();
        areas.set(element, rect.width * rect.height);
      }
      const text = all
        .filter((element) =>
          [...element.childNodes].some(
            (node) => node.nodeType === 3 && (node.textContent ?? "").trim().length > 8,
          ),
        )
        .sort((a, b) => (areas.get(b) ?? 0) - (areas.get(a) ?? 0))[0];

      if (kind === "hide-text") {
        if (!text) return false;
        text.style.visibility = "hidden";
        return true;
      }

      if (kind === "shrink-text") {
        if (!text) return false;
        // transform is ignored on inline boxes, so force a box first.
        text.style.display = "inline-block";
        text.style.transformOrigin = "top left";
        text.style.transform = "scale(0.5)";
        return true;
      }

      if (kind === "drop-graphic") {
        const graphic = all
          .filter((element) => {
            const painted =
              element.tagName === "IMG" ||
              element.tagName === "svg" ||
              element.tagName === "CANVAS" ||
              getComputedStyle(element).backgroundImage !== "none";
            return painted && (areas.get(element) ?? 0) > 10_000;
          })
          .sort((a, b) => (areas.get(b) ?? 0) - (areas.get(a) ?? 0))[0];
        if (!graphic) return false;
        graphic.style.visibility = "hidden";
        return true;
      }

      if (kind === "shift-8px") {
        for (const child of root.children) {
          (child as HTMLElement).style.transform = "translate(8px, 8px)";
        }
        return true;
      }

      const filled = all
        .filter((element) => {
          const background = getComputedStyle(element).backgroundColor;
          const opaque = background !== "rgba(0, 0, 0, 0)" && background !== "transparent";
          return opaque && (areas.get(element) ?? 0) > 20_000;
        })
        .sort((a, b) => (areas.get(b) ?? 0) - (areas.get(a) ?? 0))[0];
      if (!filled) return false;
      filled.style.backgroundColor = "#c0392b";
      return true;
    }, fault);

    if (!applied) {
      console.log(`  ${fault.padEnd(13)} skipped: this slide has nothing to break that way`);
      continue;
    }

    const screenshot = await page.locator("#slide-container").screenshot();
    const score = await scoreAgainstPowerPoint(screenshot, powerPointPath);
    control ??= score;
    controlShot ??= screenshot;

    // A fault that changed no pixels says nothing about the gate, so keep it
    // out of the "missed" column it would otherwise land in.
    if (!fault.startsWith("control") && screenshot.equals(controlShot)) {
      console.log(`  ${fault.padEnd(13)} no pixel change: the fault did not land on this slide`);
      continue;
    }

    const caught = caughtBy(score, control);
    console.log(
      `  ${fault.padEnd(13)} overall ${score.overall.toFixed(3)}  ` +
        `tile ${score.worstTile.toFixed(3)}  ` +
        `region moved ${worstRegionMoved(score, control).toFixed(3)}  ` +
        `-> ${fault.startsWith("control") ? (caught ? `TRIPPED ${caught}` : "quiet") : caught || "MISSED"}`,
    );
  }
}

await browser.close();
