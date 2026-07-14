/**
 * Head-to-head benchmark: our parser vs @aiden0z/pptx-renderer@1.2.3.
 *
 * Run with: pnpm -F "@diceui/pptx-core" bench
 *
 * Both sides receive the identical ArrayBuffer produced by generateDeck(),
 * so the comparison is purely parser throughput, not fixture differences.
 *
 * jsdom is used instead of happy-dom because the upstream package calls
 * Element.lookupNamespaceURI() unconditionally: a method happy-dom doesn't
 * implement. Our package patches this with optional chaining; upstream doesn't.
 *
 * @vitest-environment jsdom
 */
import {
  buildPresentation as upstreamBuild,
  materializeSlideNodes as upstreamMaterialize,
  parseZip as upstreamParseZip,
  parseZipLazyMedia as upstreamParseZipLazyMedia,
} from "@aiden0z/pptx-renderer";
import { bench, describe } from "vitest";

import {
  buildPresentation as ourBuild,
  materializeSlide as ourMaterialize,
} from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import { DECK_SPECS, generateDeck } from "../fixtures/bench-decks";

const mediumBuffer = await generateDeck(DECK_SPECS.medium);
const largeBuffer = await generateDeck(DECK_SPECS.large);

// Pre-parse files for stages that only benchmark a single step
const mediumFilesOurs = await readPptx(mediumBuffer);
const largeFilesOurs = await readPptx(largeBuffer);
const mediumFilesUpstream = await upstreamParseZip(mediumBuffer);
const largeFilesUpstream = await upstreamParseZip(largeBuffer);

// ─── readPptx ────────────────────────────────────────────────────────────────

describe("readPptx: medium (20 slides)", () => {
  bench("ours", async () => {
    await readPptx(mediumBuffer);
  });
  bench("upstream", async () => {
    await upstreamParseZip(mediumBuffer);
  });
});

describe("readPptx: large (100 slides)", () => {
  bench("ours", async () => {
    await readPptx(largeBuffer);
  });
  bench("upstream", async () => {
    await upstreamParseZip(largeBuffer);
  });
});

describe("readPptx lazyMedia: large (100 slides)", () => {
  bench("ours", async () => {
    await readPptx(largeBuffer, { lazyMedia: true });
  });
  bench("upstream", async () => {
    await upstreamParseZipLazyMedia(largeBuffer);
  });
});

// ─── buildPresentation ───────────────────────────────────────────────────────

describe("buildPresentation: medium", () => {
  bench("ours", () => {
    ourBuild(mediumFilesOurs);
  });
  bench("upstream", () => {
    upstreamBuild(mediumFilesUpstream);
  });
});

describe("buildPresentation: large", () => {
  bench("ours", () => {
    ourBuild(largeFilesOurs);
  });
  bench("upstream", () => {
    upstreamBuild(largeFilesUpstream);
  });
});

// ─── end-to-end load (buffer → first slide ready) ────────────────────────────

describe("time-to-first-slide: medium", () => {
  bench("ours", async () => {
    const files = await readPptx(mediumBuffer);
    const pres = ourBuild(files, { lazySlides: true });
    ourMaterialize(pres, pres.slides[0]);
  });
  bench("upstream", async () => {
    const files = await upstreamParseZip(mediumBuffer);
    const pres = upstreamBuild(files, { lazySlides: true });
    upstreamMaterialize(pres, pres.slides[0]);
  });
});

describe("time-to-first-slide: large", () => {
  bench("ours", async () => {
    const files = await readPptx(largeBuffer);
    const pres = ourBuild(files, { lazySlides: true });
    ourMaterialize(pres, pres.slides[0]);
  });
  bench("upstream", async () => {
    const files = await upstreamParseZip(largeBuffer);
    const pres = upstreamBuild(files, { lazySlides: true });
    upstreamMaterialize(pres, pres.slides[0]);
  });
});

// ─── full-deck materialization ────────────────────────────────────────────────
// Skipped for upstream: jsdom's DOM node allocations for 100 slides × repeated
// rounds exhaust the default Node heap (~4 GB). The per-slide cost can be
// inferred from time-to-first-slide above times slide count.
describe("full-deck materialize: large (100 slides)", () => {
  bench("ours", () => {
    const pres = ourBuild(largeFilesOurs);
    for (const slide of pres.slides) ourMaterialize(pres, slide);
  });
});
