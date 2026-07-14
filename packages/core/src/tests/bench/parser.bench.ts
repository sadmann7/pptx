/**
 * Stage-level benchmarks for the parse/render pipeline.
 *
 * Run with: pnpm -F "@diceui/pptx-core" bench
 *
 * Each stage is measured in isolation against pre-built inputs so the numbers
 * attribute time to the right phase:
 *   readPptx          : unzip + part extraction
 *   buildPresentation : XML parsing + model construction
 *   materialize       : node materialization + placeholder inheritance
 *   renderSlide       : DOM/SVG generation for one slide
 *   renderAllSlides   : full-deck rendering
 *   search            : text index + query
 *   serialize         : JSON export
 */
import { bench, describe } from "vitest";

import type { PresentationData } from "../../model/presentation";
import {
  buildPresentation,
  materializeAllSlides,
  materializeSlide,
} from "../../model/presentation";
import { serializePresentation } from "../../model/serialize";
import { buildTextIndex, searchText } from "../../model/text-search";
import type { PptxFiles } from "../../ooxml/zip";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { DECK_SPECS, generateDeck } from "../fixtures/bench-decks";

interface Workload {
  buffer: ArrayBuffer;
  files: PptxFiles;
  /** Fully built + materialized, for render/search benches. */
  presentation: PresentationData;
}

async function prepare(spec: (typeof DECK_SPECS)[keyof typeof DECK_SPECS]): Promise<Workload> {
  const buffer = await generateDeck(spec);
  const files = await readPptx(buffer);
  const presentation = buildPresentation(files);
  materializeAllSlides(presentation);
  return { buffer, files, presentation };
}

const small = await prepare(DECK_SPECS.small);
const medium = await prepare(DECK_SPECS.medium);
const large = await prepare(DECK_SPECS.large);

describe("readPptx", () => {
  bench("small (5 slides)", async () => {
    await readPptx(small.buffer);
  });
  bench("medium (20 slides)", async () => {
    await readPptx(medium.buffer);
  });
  bench("large (100 slides)", async () => {
    await readPptx(large.buffer);
  });
});

describe("buildPresentation", () => {
  bench("small", () => {
    buildPresentation(small.files);
  });
  bench("medium", () => {
    buildPresentation(medium.files);
  });
  bench("large", () => {
    buildPresentation(large.files);
  });
});

describe("materializeAllSlides", () => {
  // Materialization caches per slide object, so rebuild the model each round;
  // the rebuild cost is reported separately above and subtracted mentally.
  bench("medium (incl. rebuild)", () => {
    const pres = buildPresentation(medium.files);
    materializeAllSlides(pres);
  });
  bench("large (incl. rebuild)", () => {
    const pres = buildPresentation(large.files);
    materializeAllSlides(pres);
  });
});

describe("renderSlide", () => {
  bench("one slide (medium deck)", () => {
    renderSlide(medium.presentation, medium.presentation.slides[0]).dispose();
  });
  bench("all slides (medium deck, 20)", () => {
    for (const slide of medium.presentation.slides) {
      renderSlide(medium.presentation, slide).dispose();
    }
  });
});

describe("end-to-end load (buffer → materialized model)", () => {
  bench("medium", async () => {
    const files = await readPptx(medium.buffer);
    const pres = buildPresentation(files);
    for (const slide of pres.slides) materializeSlide(pres, slide);
  });
});

describe("time-to-first-slide (buffer → first slide ready)", () => {
  bench("large eager (parse all slides up front)", async () => {
    const files = await readPptx(large.buffer);
    const pres = buildPresentation(files);
    materializeSlide(pres, pres.slides[0]);
  });
  bench("large lazy (parse only first slide)", async () => {
    const files = await readPptx(large.buffer);
    const pres = buildPresentation(files, { lazySlides: true });
    materializeSlide(pres, pres.slides[0]);
  });
});

describe("text search", () => {
  const index = buildTextIndex(large.presentation);
  bench("buildTextIndex (large)", () => {
    buildTextIndex(large.presentation);
  });
  bench("searchText (large index)", () => {
    searchText(index, "revenue");
  });
});

describe("serializePresentation", () => {
  bench("large", () => {
    serializePresentation(large.presentation);
  });
});
