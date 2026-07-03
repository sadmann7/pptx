/**
 * CPU profile of the internal MTX decompressor against a real deck.
 * Gated: PROFILE_DECK must point at a .pptx with embedded fonts.
 */
import fs from "node:fs";
import { Session } from "node:inspector";
import { describe, expect, it } from "vitest";

import { decompressMtx } from "../fonts/mtx";

const deckPath = process.env.PROFILE_DECK;

function eotPayload(data: Uint8Array): Uint8Array {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const fontDataSize = view.getUint32(4, true);
  const version = view.getUint32(8, true);
  let offset = 80;
  for (let i = 0; i < 4; i++) offset += 4 + view.getUint16(offset + 2, true);
  if (version >= 0x00020001) offset += 4 + view.getUint16(offset + 2, true);
  if (version >= 0x00020002) {
    offset += 8;
    offset += 4 + view.getUint16(offset + 2, true);
    offset += 8 + view.getUint32(offset + 4, true);
  }
  return data.subarray(offset, Math.min(offset + fontDataSize, data.length));
}

describe.skipIf(!deckPath)("mtx cpu profile", () => {
  it("profiles decompressMtx", async () => {
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(fs.readFileSync(deckPath!));
    const payloads: Uint8Array[] = [];
    for (const name of Object.keys(zip.files).filter((n) => n.endsWith(".fntdata"))) {
      payloads.push(eotPayload(new Uint8Array(await zip.files[name].async("uint8array"))));
    }
    for (const p of payloads) decompressMtx(p); // warm up

    const session = new Session();
    session.connect();
    const post = (m: string, p?: object): Promise<unknown> =>
      new Promise((res, rej) => session.post(m, p, (e, r) => (e ? rej(e) : res(r))));
    await post("Profiler.enable");
    await post("Profiler.setSamplingInterval", { interval: 50 });
    await post("Profiler.start");
    const t0 = performance.now();
    for (let i = 0; i < 10; i++) {
      for (const p of payloads) decompressMtx(p);
    }
    const elapsed = performance.now() - t0;
    const { profile } = (await post("Profiler.stop")) as {
      profile: {
        nodes: Array<{
          callFrame: { functionName: string; url: string; lineNumber: number };
          hitCount?: number;
        }>;
      };
    };
    session.disconnect();

    const totalSamples = profile.nodes.reduce((s, n) => s + (n.hitCount ?? 0), 0);
    const byFn = new Map<string, { name: string; ms: number }>();
    for (const node of profile.nodes) {
      const hits = node.hitCount ?? 0;
      if (!hits) continue;
      const { functionName, url, lineNumber } = node.callFrame;
      const key = `${functionName}@${url}:${lineNumber}`;
      const row = byFn.get(key) ?? {
        name: `${functionName || "(anon)"} :${lineNumber + 1}`,
        ms: 0,
      };
      row.ms += (hits * elapsed) / totalSamples;
      byFn.set(key, row);
    }
    const rows = [...byFn.values()].sort((a, b) => b.ms - a.ms).slice(0, 20);
    const lines = [`\n===== internal MTX x10x${payloads.length}: ${elapsed.toFixed(0)}ms =====`];
    for (const r of rows) {
      lines.push(
        `${((r.ms / elapsed) * 100).toFixed(1).padStart(5)}%  ${r.ms.toFixed(0).padStart(6)}ms  ${r.name}`,
      );
    }
    fs.appendFileSync(`${process.env.TEMP ?? "/tmp"}/pptx-profile.txt`, lines.join("\n") + "\n");
    expect(rows.length).toBeGreaterThan(0);
  }, 120000);
});
