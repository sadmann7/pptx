import { decompressMtx as referenceDecompressMtx } from "mtx-decompressor";
import fs from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { decodeEmbeddedFont } from "../fonts/font-decode";
import { decompressMtx } from "../fonts/mtx";

const FIXTURES_DIR = path.join(__dirname, "fixtures");
const FIXTURES = ["InstrumentSansSemiBold-regular.fntdata", "SpaceGroteskSemiBold-bold.fntdata"];

/** Extract the MTX payload from the EOT container (mirrors font-decode.ts). */
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

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(FIXTURES_DIR, name)));
}

describe("internal MTX decompressor", () => {
  it.each(FIXTURES)("produces byte-identical output to mtx-decompressor for %s", (name) => {
    const payload = eotPayload(loadFixture(name));
    const expected = referenceDecompressMtx(payload, { compressed: true, encrypted: false });
    const actual = decompressMtx(payload, { compressed: true, encrypted: false });

    expect(actual.length).toBe(expected.length);
    expect(Buffer.from(actual).equals(Buffer.from(expected))).toBe(true);
  });

  it.each(FIXTURES)("produces a valid TrueType binary for %s", (name) => {
    const payload = eotPayload(loadFixture(name));
    const font = decompressMtx(payload, { compressed: true, encrypted: false });
    // sfnt version 1.0 (TrueType outlines)
    expect([...font.subarray(0, 4)]).toEqual([0x00, 0x01, 0x00, 0x00]);
    const numTables = (font[4] << 8) | font[5];
    expect(numTables).toBeGreaterThan(5);
  });

  it("returns input unchanged when compressed: false", () => {
    const data = new Uint8Array([1, 2, 3, 4]);
    expect([...decompressMtx(data, { compressed: false })]).toEqual([1, 2, 3, 4]);
  });

  it("XOR-decrypts before unpacking when encrypted: true", () => {
    const payload = eotPayload(loadFixture(FIXTURES[0]));
    const encrypted = payload.map((b) => b ^ 80);
    const viaEncrypted = decompressMtx(encrypted, { compressed: true, encrypted: true });
    const direct = decompressMtx(payload, { compressed: true, encrypted: false });
    expect(Buffer.from(viaEncrypted).equals(Buffer.from(direct))).toBe(true);
  });

  it("rejects malformed headers", () => {
    expect(() => decompressMtx(new Uint8Array(4))).toThrow(/too small/);
    const badOffsets = new Uint8Array(16);
    badOffsets[6] = 2; // offset2 = 2 < 10
    expect(() => decompressMtx(badOffsets)).toThrow(/out of bounds/);
  });
});

// Exhaustive oracle + timing against a real deck's embedded fonts.
// Gated: set PROFILE_DECK to a .pptx path with embedded fonts.
describe.skipIf(!process.env.PROFILE_DECK)("real deck oracle", () => {
  it("matches mtx-decompressor byte-for-byte on every embedded part", async () => {
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(fs.readFileSync(process.env.PROFILE_DECK!));
    const names = Object.keys(zip.files).filter((n) => n.endsWith(".fntdata"));
    expect(names.length).toBeGreaterThan(0);

    const payloads: Uint8Array[] = [];
    for (const name of names) {
      const payload = eotPayload(new Uint8Array(await zip.files[name].async("uint8array")));
      payloads.push(payload);
      const expected = referenceDecompressMtx(payload, { compressed: true, encrypted: false });
      const actual = decompressMtx(payload, { compressed: true, encrypted: false });
      expect(Buffer.from(actual).equals(Buffer.from(expected)), name).toBe(true);
    }

    // Warmed timing: 5 rounds each after a warm-up round (already done above).
    const ROUNDS = 5;
    let t = performance.now();
    for (let r = 0; r < ROUNDS; r++) {
      for (const p of payloads) referenceDecompressMtx(p, { compressed: true, encrypted: false });
    }
    const oldMs = (performance.now() - t) / ROUNDS;
    t = performance.now();
    for (let r = 0; r < ROUNDS; r++) {
      for (const p of payloads) decompressMtx(p, { compressed: true, encrypted: false });
    }
    const newMs = (performance.now() - t) / ROUNDS;

    fs.appendFileSync(
      `${process.env.TEMP ?? "/tmp"}/pptx-profile.txt`,
      `\nMTX oracle: ${names.length} parts identical. per round: reference=${oldMs.toFixed(0)}ms internal=${newMs.toFixed(0)}ms (${(oldMs / newMs).toFixed(2)}x)\n`,
    );
  }, 120000);
});

describe("decodeEmbeddedFont end-to-end", () => {
  it.each(FIXTURES)("decodes %s through the full EOT pipeline", (name) => {
    const decoded = decodeEmbeddedFont(loadFixture(name));
    expect(decoded).toBeDefined();
    expect([...decoded!.subarray(0, 4)]).toEqual([0x00, 0x01, 0x00, 0x00]);
  });
});
