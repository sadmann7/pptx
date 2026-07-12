/**
 * CPU profiling harness (not a real test). Skipped unless PROFILE=1.
 *
 * Usage:
 *   $env:PROFILE="1"; pnpm -F "@diceui/pptx-core" exec vitest run src/tests/model/profile.test.ts; $env:PROFILE=$null
 *
 * Captures a V8 CPU profile around the hot pipeline stage and prints the top
 * functions by self time, aggregated per function, filtered to package code.
 */
import fs from "node:fs";
import { Session } from "node:inspector";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { buildPresentation, materializeAllSlides } from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { DECK_SPECS, generateDeck } from "../fixtures/bench-decks";

interface ProfileRow {
  name: string;
  file: string;
  selfMs: number;
  totalMs: number;
}

async function captureProfile(workload: () => void | Promise<void>): Promise<ProfileRow[]> {
  const session = new Session();
  session.connect();
  const post = (method: string, params?: object): Promise<unknown> =>
    new Promise((resolve, reject) => {
      session.post(method, params, (err, result) => (err ? reject(err) : resolve(result)));
    });

  await post("Profiler.enable");
  await post("Profiler.setSamplingInterval", { interval: 100 });
  await post("Profiler.start");
  await workload();
  const { profile } = (await post("Profiler.stop")) as {
    profile: {
      nodes: Array<{
        id: number;
        callFrame: { functionName: string; url: string; lineNumber: number };
        hitCount?: number;
        children?: number[];
      }>;
      samples?: number[];
      timeDeltas?: number[];
    };
  };
  session.disconnect();

  // Self time per node from sample hit counts x sampling interval.
  const totalSamples = profile.nodes.reduce((sum, n) => sum + (n.hitCount ?? 0), 0);
  const totalMs = (profile.timeDeltas?.reduce((a, b) => a + b, 0) ?? totalSamples * 100) / 1000;
  const msPerSample = totalSamples > 0 ? totalMs / totalSamples : 0;

  // Aggregate by function identity (name + file + line).
  const byFunction = new Map<string, ProfileRow>();
  for (const node of profile.nodes) {
    const hits = node.hitCount ?? 0;
    if (hits === 0) continue;
    const { functionName, url, lineNumber } = node.callFrame;
    const file = url.replace(/^file:\/\/\/?/, "").replace(/\\/g, "/");
    const key = `${functionName}@${file}:${lineNumber}`;
    const row = byFunction.get(key) ?? {
      name: functionName || "(anonymous)",
      file: `${file.split("/").slice(-2).join("/")}:${lineNumber + 1}`,
      selfMs: 0,
      totalMs: 0,
    };
    row.selfMs += hits * msPerSample;
    byFunction.set(key, row);
  }

  return [...byFunction.values()].sort((a, b) => b.selfMs - a.selfMs);
}

function printTop(label: string, rows: ProfileRow[], limit = 30): void {
  const total = rows.reduce((sum, r) => sum + r.selfMs, 0);
  const lines = [`\n===== ${label}: total self time ${total.toFixed(1)}ms =====`];
  for (const row of rows.slice(0, limit)) {
    const pct = ((row.selfMs / total) * 100).toFixed(1).padStart(5);
    lines.push(`${pct}%  ${row.selfMs.toFixed(1).padStart(8)}ms  ${row.name}  [${row.file}]`);
  }
  const report = lines.join("\n");
  console.log(report);
  // Vitest intercepts console output in workers; persist to a file as well.
  const outPath = path.join(os.tmpdir(), "pptx-profile.txt");
  fs.appendFileSync(outPath, report + "\n");
}

describe.skipIf(!process.env.PROFILE)("cpu profile", () => {
  it("profiles buildPresentation (large deck)", async () => {
    const buffer = await generateDeck(DECK_SPECS.large);
    const files = await readPptx(buffer);
    // Warm up JIT
    buildPresentation(files);

    const rows = await captureProfile(() => {
      for (let i = 0; i < 5; i++) {
        buildPresentation(files);
      }
    });
    printTop("buildPresentation x5 (large)", rows);
    expect(rows.length).toBeGreaterThan(0);
  }, 120000);

  it("profiles renderSlide (medium deck, all slides)", async () => {
    const buffer = await generateDeck(DECK_SPECS.medium);
    const files = await readPptx(buffer);
    const pres = buildPresentation(files);
    materializeAllSlides(pres);
    renderSlide(pres, pres.slides[0]).dispose(); // warm up

    const rows = await captureProfile(() => {
      for (let i = 0; i < 3; i++) {
        for (const slide of pres.slides) {
          renderSlide(pres, slide).dispose();
        }
      }
    });
    printTop("renderSlide x3x20 (medium)", rows);
    expect(rows.length).toBeGreaterThan(0);
  }, 120000);
});
