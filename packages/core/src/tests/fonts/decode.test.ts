/**
 * End-to-end decode of the real `.fntdata` parts PowerPoint produces.
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

import { decodeEmbeddedFont } from "../../fonts/decode";
import { deobfuscateFont } from "../../fonts/deobfuscate";
import { parseEotMetadata } from "../../fonts/mtx";
import { TtfReader } from "./ttf-reader";

const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "fixtures");
const PARTS = fs.readdirSync(FIXTURES).filter((name) => name.endsWith(".fntdata"));

function read(name: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(FIXTURES, name)));
}

describe("embedded font parts", () => {
  it("finds the fixtures", () => {
    expect(PARTS.length).toBeGreaterThan(0);
  });

  it.each(PARTS)("decodes %s into a usable sfnt", (name) => {
    const raw = read(name);
    const metadata = parseEotMetadata(raw);
    expect(metadata.fontDataSize).toBeGreaterThan(0);
    expect(metadata.fontDataOffset + metadata.fontDataSize).toBe(metadata.totalSize);

    const decoded = decodeEmbeddedFont(raw);
    expect(decoded).toBeDefined();

    const font = new TtfReader(decoded!);
    const tags = font.tableTags;
    expect(tags).toEqual([...tags].sort());
    expect(tags).toContain("head");
    expect(tags).toContain("glyf");
    expect(font.numGlyphs).toBeGreaterThan(0);

    // Exercising every glyph proves loca and glyf agree end to end.
    for (let i = 0; i < font.numGlyphs; i++) expect(() => font.glyph(i)).not.toThrow();
  });

  it.each(PARTS)("reports %s container flags rather than guessing", (name) => {
    // Both fixtures are version-3 EOTs with MTX-compressed payloads, so these
    // cases put real PowerPoint output through LZCOMP and CTF.
    const metadata = parseEotMetadata(read(name));
    expect(metadata.version).toBe(3);
    expect(metadata.compressed).toBe(true);
    expect(metadata.encrypted).toBe(false);
    expect(metadata.badVersion).toBe(false);
  });

  it.each(PARTS)("reconstructs consistent metrics tables for %s", (name) => {
    const font = new TtfReader(decodeEmbeddedFont(read(name))!);
    const hhea = font.tables.get("hhea");
    const hmtx = font.tables.get("hmtx");
    expect(hhea).toBeDefined();
    expect(hmtx).toBeDefined();

    const metrics = new DataView(hhea!.buffer, hhea!.byteOffset, hhea!.byteLength).getUint16(
      34,
      false,
    );
    expect(metrics).toBeGreaterThan(0);
    expect(metrics).toBeLessThanOrEqual(font.numGlyphs);
    // Long metrics, then one side bearing for each remaining glyph.
    expect(hmtx!.length).toBe(metrics * 4 + (font.numGlyphs - metrics) * 2);

    const cvt = font.tables.get("cvt ");
    if (cvt) expect(cvt.length % 2).toBe(0);
  });

  it("returns undefined for data that is not a font", () => {
    expect(decodeEmbeddedFont(new Uint8Array(0))).toBeUndefined();
    expect(decodeEmbeddedFont(Uint8Array.from([1, 2, 3, 4]))).toBeUndefined();
    expect(decodeEmbeddedFont(new Uint8Array(200))).toBeUndefined();
  });

  it("truncating the payload is rejected instead of yielding a broken font", () => {
    const raw = read(PARTS[0]!);
    expect(decodeEmbeddedFont(raw.subarray(0, raw.length - 1))).toBeUndefined();
  });

  it("deobfuscates a GUID-keyed part before parsing the container", () => {
    const raw = read(PARTS[0]!);
    const key = "{9A0A2B1C-3D4E-5F60-7182-93A4B5C6D7E8}";
    const obfuscated = deobfuscateFont(raw, key); // The transform is its own inverse.
    expect(obfuscated).not.toEqual(raw);
    expect(decodeEmbeddedFont(obfuscated, key)).toEqual(decodeEmbeddedFont(raw));
  });
});
