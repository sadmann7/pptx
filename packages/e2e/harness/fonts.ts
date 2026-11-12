/**
 * Embedded-font decode harness.
 *
 * Loads a deck and decodes its embedded fonts once, while a frame monitor
 * watches the main thread. font-perf.spec.ts drives it once per mode, in a
 * fresh page each time, to see what the Worker pool is worth: it buys a main
 * thread that stays free more than it buys throughput, so the frame numbers
 * matter as much as the wall clock.
 *
 * Each mode gets its own page load because the first run leaves its faces in
 * `document.fonts`, which makes a second registration in the same document
 * cheaper than the one the user pays for.
 *
 * Query params:
 *   file   deck to load, served from fixtures/ (or decks/ for local decks)
 *   mode   which path decodes the parts (default "worker"):
 *            worker  the shipped pool
 *            main    the shipped fallback, which hands the event loop a turn
 *                    between parts; `Worker` is hidden to select it
 *            sliced  the fallback without a worker to fall back from: decodes
 *                    here and only yields once a slice runs long, which is what
 *                    dropping the worker would mean
 *   parts  multiplies the deck's font parts by this factor, to see how the two
 *          modes diverge on decks heavier than the fixtures (default 1)
 *   mtx    replaces the deck's font parts with MTX-compressed ones, which is
 *          what PowerPoint itself writes; every fixture deck embeds plain EOT,
 *          where decoding is a header strip and costs nothing
 *   build  "dist" to use the built font entry (inlined blob worker, minified),
 *          which is what consumers get; "src" (default) is what Vite spawns
 *          from source, where the worker is a module the browser fetches
 */
import type { EmbeddedFontEntry, PresentationData } from "@diceui/pptx-core";
import { buildPresentation, readPptx } from "@diceui/pptx-core";
import type { EmbeddedFontError } from "@diceui/pptx-core/fonts";

const params = new URLSearchParams(location.search);
const file = params.get("file");
const modeParam = params.get("mode");
const mode: BenchMode = modeParam === "main" || modeParam === "sliced" ? modeParam : "worker";
const partsFactor = Math.max(1, Number.parseInt(params.get("parts") ?? "1", 10));
const useMtx = params.get("mtx") === "1";
const useBuild = params.get("build") === "dist";

type FontsModule = typeof import("@diceui/pptx-core/fonts");

// Loaded through a plain URL, since bundling it would stop the blob worker and
// minified decoder being timed exactly as published.
async function loadFontsModule(): Promise<FontsModule> {
  if (!useBuild) return import("@diceui/pptx-core/fonts");
  // The specifier is a URL the dev server answers, not a module TypeScript can
  // resolve, so it goes through a variable to keep it out of resolution.
  const url = "/core-dist/fonts.mjs";
  const module = (await import(/* @vite-ignore */ url)) as Partial<FontsModule>;
  if (!module.loadEmbeddedFonts) {
    throw new Error("/core-dist/fonts.mjs is missing; run pnpm -F @diceui/pptx-core build");
  }
  return module as FontsModule;
}

// MTX stands in for the payload PowerPoint writes, which no deck on disk contains.
const MTX_PARTS = [
  "/fonts/InstrumentSansSemiBold-regular.fntdata",
  "/fonts/SpaceGroteskSemiBold-bold.fntdata",
];

const VARIANTS = [
  { key: "regular", weight: "normal", style: "normal" },
  { key: "bold", weight: "bold", style: "normal" },
  { key: "italic", weight: "normal", style: "italic" },
  { key: "boldItalic", weight: "bold", style: "italic" },
] as const;

const VARIANT_KEYS = VARIANTS.map((variant) => variant.key);

const SLICE_MS = 5;

const FRAME_BUDGET_MS = 1000 / 60;

// Real decks embed only 3 to 7 parts, so multiplying reaches the deck sizes the
// loader was written for. The bytes are shared, so only the decode work repeats.
function multiplyFontParts(presentation: PresentationData, factor: number): void {
  const base = presentation.embeddedFonts;
  if (!base || factor <= 1) return;

  const added: EmbeddedFontEntry[] = [];
  for (let copy = 1; copy < factor; copy++) {
    for (const entry of base) {
      const clone: EmbeddedFontEntry = { ...entry, typeface: `${entry.typeface} Copy${copy}` };
      for (const key of VARIANT_KEYS) {
        const variant = entry[key];
        if (!variant) continue;
        const bytes = presentation.fonts.get(variant.path);
        if (!bytes) continue;
        const path = `${variant.path}?copy=${copy}`;
        presentation.fonts.set(path, bytes);
        clone[key] = { ...variant, path };
      }
      added.push(clone);
    }
  }
  base.push(...added);
}

// The MTX bytes carry their own EOT header, so the deck's fontKey no longer
// applies and is dropped.
async function applyMtxFontParts(presentation: PresentationData): Promise<void> {
  const payloads: Uint8Array[] = [];
  for (const url of MTX_PARTS) {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`fetch ${url}: HTTP ${response.status}`);
    payloads.push(new Uint8Array(await response.arrayBuffer()));
  }

  let index = 0;
  for (const entry of presentation.embeddedFonts ?? []) {
    for (const key of VARIANT_KEYS) {
      const variant = entry[key];
      if (!variant) continue;
      const payload = payloads[index++ % payloads.length];
      if (!payload) continue;
      presentation.fonts.set(variant.path, payload);
      entry[key] = { path: variant.path };
    }
  }
}

function fontPartPaths(presentation: PresentationData): string[] {
  const paths = new Set<string>();
  for (const entry of presentation.embeddedFonts ?? []) {
    for (const key of VARIANT_KEYS) {
      const path = entry[key]?.path;
      if (path) paths.add(path);
    }
  }
  return [...paths];
}

function typefacesForPath(presentation: PresentationData, path: string): string[] {
  const typefaces = new Set<string>();
  for (const entry of presentation.embeddedFonts ?? []) {
    for (const key of VARIANT_KEYS) {
      if (entry[key]?.path === path) typefaces.add(entry.typeface);
    }
  }
  return [...typefaces];
}

function repointVariants(presentation: PresentationData, from: string, to: string): void {
  for (const entry of presentation.embeddedFonts ?? []) {
    for (const key of VARIANT_KEYS) {
      const variant = entry[key];
      if (variant?.path === from) entry[key] = { ...variant, path: to };
    }
  }
}

// One part per failure depth: a dropped part never reaches the decoder, an empty
// or garbage one is rejected by it, and an sfnt signature over junk passes it to
// be refused by the browser's FontFace.
function breakFontParts(presentation: PresentationData): BrokenFontPart[] {
  const paths = fontPartPaths(presentation);
  const broken: BrokenFontPart[] = [];
  const record = (path: string, kind: BrokenFontPart["kind"], stage: BrokenFontPart["stage"]) => {
    broken.push({ path, kind, stage, typefaces: typefacesForPath(presentation, path) });
  };

  const missing = paths[0];
  if (missing) {
    presentation.fonts.delete(missing);
    record(missing, "missing", "missing");
  }

  // Present but empty, which the deck did carry, so it is the decoder's to
  // reject rather than a part the package never had.
  const empty = paths[1];
  if (empty) {
    presentation.fonts.set(empty, new Uint8Array(0));
    record(empty, "empty", "decode");
  }

  // Long enough to hold an EOT header, so the container parse rejects it on
  // its own terms (zeroed sizes) rather than for being too short to read.
  const undecodable = paths[2];
  if (undecodable) {
    presentation.fonts.set(undecodable, new Uint8Array(200));
    record(undecodable, "undecodable", "decode");
  }

  // A TrueType sfnt version and nothing else behind it: the decoder passes it
  // through as an already-raw font, then FontFace finds no tables. A second
  // family is pointed at the same part first, since the bytes are what failed
  // and both families lose them together.
  const unregisterable = paths[3];
  const donor = paths[4];
  if (unregisterable) {
    if (donor) repointVariants(presentation, donor, unregisterable);
    const stub = new Uint8Array(64);
    stub.set([0x00, 0x01, 0x00, 0x00]);
    presentation.fonts.set(unregisterable, stub);
    record(unregisterable, "unregisterable", "register");
  }

  return broken;
}

function fontPartStats(presentation: PresentationData): { parts: number; bytes: number } {
  const paths = new Set<string>();
  for (const entry of presentation.embeddedFonts ?? []) {
    for (const key of VARIANT_KEYS) {
      const path = entry[key]?.path;
      if (path && presentation.fonts.get(path)?.length) paths.add(path);
    }
  }
  let bytes = 0;
  for (const path of paths) bytes += presentation.fonts.get(path)?.length ?? 0;
  return { parts: paths.size, bytes };
}

interface MainThreadStats {
  frames: number;
  worstFrameMs: number;
  blockedMs: number;
  longTasks: number;
  worstTaskMs: number;
  totalTaskMs: number;
}

function startMainThreadMonitor(): () => MainThreadStats {
  let frames = 0;
  let worstFrameMs = 0;
  let blockedMs = 0;
  let longTasks = 0;
  let worstTaskMs = 0;
  let totalTaskMs = 0;
  let running = true;
  let last = performance.now();

  const onFrame = (now: number): void => {
    const delta = now - last;
    last = now;
    frames++;
    if (delta > worstFrameMs) worstFrameMs = delta;
    if (delta > FRAME_BUDGET_MS) blockedMs += delta - FRAME_BUDGET_MS;
    if (running) requestAnimationFrame(onFrame);
  };
  requestAnimationFrame(onFrame);

  let observer: PerformanceObserver | undefined;
  try {
    observer = new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) {
        longTasks++;
        totalTaskMs += entry.duration;
        if (entry.duration > worstTaskMs) worstTaskMs = entry.duration;
      }
    });
    observer.observe({ entryTypes: ["longtask"] });
  } catch {
    // No long-task support (Firefox); the frame numbers carry the result.
  }

  return () => {
    running = false;
    observer?.disconnect();
    return { frames, worstFrameMs, blockedMs, longTasks, worstTaskMs, totalTaskMs };
  };
}

// Yields only once a slice runs long, since yielding after every part as the
// shipped fallback does costs about as much as decoding a cheap part.
async function decodeSliced(
  presentation: PresentationData,
  decodeEmbeddedFont: FontsModule["decodeEmbeddedFont"],
): Promise<void> {
  interface Part {
    bytes: Uint8Array;
    fontKey?: string;
    faces: { typeface: string; weight: string; style: string }[];
  }

  const parts = new Map<string, Part>();
  for (const entry of presentation.embeddedFonts ?? []) {
    for (const { key, weight, style } of VARIANTS) {
      const variant = entry[key];
      const bytes = variant && presentation.fonts.get(variant.path);
      if (!variant || !bytes?.length) continue;
      const face = { typeface: entry.typeface, weight, style };
      const part = parts.get(variant.path);
      if (part) part.faces.push(face);
      else parts.set(variant.path, { bytes, fontKey: variant.fontKey, faces: [face] });
    }
  }

  const registrations: Promise<unknown>[] = [];
  let sliceStart = performance.now();
  for (const part of parts.values()) {
    const decoded = decodeEmbeddedFont(part.bytes, part.fontKey);
    if (decoded) {
      const buffer = decoded.buffer.slice(
        decoded.byteOffset,
        decoded.byteOffset + decoded.byteLength,
      ) as ArrayBuffer;
      for (const face of part.faces) {
        const fontFace = new FontFace(face.typeface, buffer, {
          weight: face.weight,
          style: face.style,
        });
        registrations.push(
          fontFace
            .load()
            .then((loaded) => document.fonts.add(loaded))
            .catch(() => {
              // Invalid font data: skip this variant, text falls back.
            }),
        );
      }
    }
    if (performance.now() - sliceStart >= SLICE_MS) {
      await new Promise((resolve) => setTimeout(resolve));
      sliceStart = performance.now();
    }
  }

  await Promise.all(registrations);
}

// Without text to reflow, `document.fonts.add` costs almost nothing and the
// main thread work would be understated in both modes.
function renderProbeText(presentation: PresentationData): void {
  const probe = document.getElementById("font-probe");
  if (!probe) return;
  probe.innerHTML = "";
  for (const entry of presentation.embeddedFonts ?? []) {
    const line = document.createElement("p");
    line.style.fontFamily = `"${entry.typeface}", serif`;
    line.style.fontSize = "24px";
    line.textContent = `${entry.typeface} — the quick brown fox jumps over the lazy dog 0123456789`;
    probe.appendChild(line);
  }
}

async function main(): Promise<void> {
  if (!file) {
    window.__renderError = "missing ?file= query param";
    return;
  }

  try {
    const response = await fetch(`/${file}`);
    if (!response.ok) throw new Error(`fetch ${file}: HTTP ${response.status}`);
    const presentation = buildPresentation(await readPptx(await response.arrayBuffer()));
    const { decodeEmbeddedFont, findPriorityTypefaces, loadEmbeddedFonts } =
      await loadFontsModule();

    if (useMtx) await applyMtxFontParts(presentation);
    multiplyFontParts(presentation, partsFactor);
    renderProbeText(presentation);

    const { parts, bytes } = fontPartStats(presentation);
    window.__fontParts = parts;
    window.__fontBytes = bytes;

    window.__benchDecodeParts = () => {
      const timings: { path: string; bytes: number; ms: number }[] = [];
      for (const entry of presentation.embeddedFonts ?? []) {
        for (const key of VARIANT_KEYS) {
          const variant = entry[key];
          const partBytes = variant && presentation.fonts.get(variant.path);
          if (!variant || !partBytes?.length) continue;
          if (timings.some((timing) => timing.path === variant.path)) continue;
          const start = performance.now();
          decodeEmbeddedFont(partBytes, variant.fontKey);
          timings.push({
            path: variant.path,
            bytes: partBytes.length,
            ms: performance.now() - start,
          });
        }
      }
      return timings;
    };

    window.__benchFontErrors = async () => {
      const broken = breakFontParts(presentation);
      const facesBefore = document.fonts.size;

      if (mode === "main") {
        Object.defineProperty(window, "Worker", { value: undefined, configurable: true });
      }

      const errors: EmbeddedFontError[] = [];
      const handle = loadEmbeddedFonts(presentation, { onError: (error) => errors.push(error) });
      await handle.complete;

      return { broken, errors, faces: document.fonts.size - facesBefore };
    };

    window.__benchFonts = async () => {
      // A part that fails to decode is skipped silently, so only the face count
      // shows whether the browser accepted the fonts.
      const facesBefore = document.fonts.size;

      if (mode === "sliced") {
        const stop = startMainThreadMonitor();
        const start = performance.now();
        await decodeSliced(presentation, decodeEmbeddedFont);
        const completeMs = performance.now() - start;
        return {
          mode,
          parts,
          bytes,
          readyMs: completeMs,
          completeMs,
          faces: document.fonts.size - facesBefore,
          ...stop(),
        };
      }

      // The loader picks its path from `typeof Worker`.
      if (mode === "main") {
        Object.defineProperty(window, "Worker", { value: undefined, configurable: true });
      }

      const priorityTypefaces = findPriorityTypefaces(
        presentation,
        presentation.slides.map((slide) => slide.sourceXml),
      );

      const stop = startMainThreadMonitor();
      const start = performance.now();
      const handle = loadEmbeddedFonts(presentation, { priorityTypefaces });
      await handle.ready;
      const readyMs = performance.now() - start;
      await handle.complete;
      const completeMs = performance.now() - start;

      return {
        mode,
        parts,
        bytes,
        readyMs,
        completeMs,
        faces: document.fonts.size - facesBefore,
        ...stop(),
      };
    };

    window.__fontsReady = true;
  } catch (error) {
    window.__renderError = error instanceof Error ? error.message : String(error);
  }
}

void main();
