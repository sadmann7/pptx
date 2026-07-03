/**
 * Stage timing against a real .pptx from disk (not a synthetic fixture).
 * Skipped unless PROFILE_DECK points at a file.
 *
 * Usage:
 *   $env:PROFILE_DECK="C:\path\to\deck.pptx"
 *   pnpm -F "@diceui/pptx-parser" exec vitest run src/tests/profile-deck.test.ts
 */
import fs from "node:fs";
import { describe, expect, it } from "vitest";

import { decodeEmbeddedFont } from "../fonts/font-decode";
import { collectPriorityTypefaces } from "../fonts/font-injector";
import { buildPresentation, materializeSlideNodes } from "../model/presentation";
import { parseZip } from "../ooxml/zip-parser";

const deckPath = process.env.PROFILE_DECK;

describe.skipIf(!deckPath || !fs.existsSync(deckPath))("real deck stage timing", () => {
  it("times each pipeline stage", async () => {
    const nodeBuffer = fs.readFileSync(deckPath!);
    const buffer = nodeBuffer.buffer.slice(
      nodeBuffer.byteOffset,
      nodeBuffer.byteOffset + nodeBuffer.byteLength,
    ) as ArrayBuffer;

    const lines: string[] = [`\n===== ${deckPath} =====`];
    const time = async <T>(label: string, fn: () => T | Promise<T>): Promise<T> => {
      const t0 = performance.now();
      const result = await fn();
      lines.push(`${(performance.now() - t0).toFixed(0).padStart(6)}ms  ${label}`);
      return result;
    };

    const files = await time("parseZip", () => parseZip(buffer));
    const pres = await time("buildPresentation (lazy)", () =>
      buildPresentation(files, { lazy: true }),
    );

    // Font prioritization: which parts must decode before first paint?
    const layoutPath = pres.slideToLayout.get(0);
    const masterPath = layoutPath ? pres.layoutToMaster.get(layoutPath) : undefined;
    const priority = collectPriorityTypefaces(pres, [
      pres.slides[0]?.sourceXml,
      layoutPath ? files.slideLayouts.get(layoutPath) : undefined,
      masterPath ? files.slideMasters.get(masterPath) : undefined,
    ]);
    const priorityPaths = new Set<string>();
    for (const entry of pres.embeddedFonts ?? []) {
      if (!priority?.has(entry.typeface)) continue;
      for (const variant of [entry.regular, entry.bold, entry.italic, entry.boldItalic]) {
        if (variant) priorityPaths.add(variant.path);
      }
    }
    await time(`decode ${priorityPaths.size} priority font parts (serial)`, () => {
      for (const path of priorityPaths) {
        const bytes = files.fonts.get(path);
        if (bytes) decodeEmbeddedFont(bytes);
      }
    });

    await time("materialize first slide", () => {
      materializeSlideNodes(pres, pres.slides[0]);
    });
    await time(`decode all ${files.fonts.size} embedded fonts (serial)`, () => {
      for (const [, bytes] of files.fonts) decodeEmbeddedFont(bytes);
    });
    await time("materialize remaining slides", () => {
      for (const slide of pres.slides) materializeSlideNodes(pres, slide);
    });

    lines.push(`slides: ${pres.slides.length}, font parts: ${files.fonts.size}`);
    fs.appendFileSync(`${process.env.TEMP ?? "/tmp"}/pptx-profile.txt`, lines.join("\n") + "\n");
    expect(pres.slides.length).toBeGreaterThan(0);
  }, 120000);
});
