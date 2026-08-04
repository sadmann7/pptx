/**
 * Stage-level benchmarks for embedded-font decoding.
 *
 * Run with: pnpm -F "@diceui/pptx-core" bench
 *
 * The inputs are the real `.fntdata` parts PowerPoint produced, so these
 * numbers cover the MTX-compressed path end to end. Each stage is fed a
 * pre-built input to attribute time to the right phase:
 *   decodeEmbeddedFont : EOT container -> validated sfnt (everything below)
 *   unpackMtx          : LZCOMP decompression of the three CTF streams
 *   parseCtf           : glyf/loca reconstruction, cvt and head rewriting
 *   buildSfnt          : table directory assembly + checksums
 *   validateSfnt       : directory bounds + whole-font checksum
 *
 * LZCOMP dominates, so treat `unpackMtx` as the number to beat.
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { bench, describe } from "vitest";

import { decodeEmbeddedFont } from "../../fonts/decode";
import { parseCtf, parseEotMetadata, unpackMtx } from "../../fonts/mtx";
import { DEFAULT_LIMITS } from "../../fonts/mtx/limits";
import { buildSfnt, validateSfnt } from "../../fonts/mtx/sfnt";

const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "fixtures");

interface Workload {
  /** Fixture name, shortened to the family so bench labels stay readable. */
  label: string;
  raw: Uint8Array;
  /** The MTX payload, with the EOT header already stripped. */
  payload: Uint8Array;
  streams: [Uint8Array, Uint8Array, Uint8Array];
  font: Uint8Array;
}

function prepare(name: string): Workload {
  const raw = new Uint8Array(fs.readFileSync(path.join(FIXTURES, name)));
  const metadata = parseEotMetadata(raw);
  const end = metadata.fontDataOffset + metadata.fontDataSize;
  const payload = raw.subarray(metadata.fontDataOffset, end);
  return {
    label: name.replace(/\.fntdata$/, ""),
    raw,
    payload,
    streams: unpackMtx(payload).streams,
    font: decodeEmbeddedFont(raw)!,
  };
}

const workloads = fs
  .readdirSync(FIXTURES)
  .filter((name) => name.endsWith(".fntdata"))
  .map(prepare);

describe("decodeEmbeddedFont (EOT -> sfnt)", () => {
  for (const { label, raw, font } of workloads) {
    bench(`${label} (${raw.length} B -> ${font.length} B)`, () => {
      decodeEmbeddedFont(raw);
    });
  }
});

describe("unpackMtx (LZCOMP, three streams)", () => {
  for (const { label, payload } of workloads) {
    bench(label, () => {
      unpackMtx(payload);
    });
  }
});

describe("parseCtf (glyf/loca reconstruction)", () => {
  for (const { label, streams } of workloads) {
    bench(label, () => {
      parseCtf(streams, DEFAULT_LIMITS);
    });
  }
});

describe("buildSfnt (directory + checksums)", () => {
  for (const { label, streams } of workloads) {
    // parseCtf leaves most tables as views into the streams, so the container
    // can be reused across rounds; buildSfnt only writes back offsets.
    const container = parseCtf(streams, DEFAULT_LIMITS);
    bench(label, () => {
      buildSfnt(container, DEFAULT_LIMITS.maxFontBytes);
    });
  }
});

describe("validateSfnt (bounds + whole-font checksum)", () => {
  for (const { label, font } of workloads) {
    bench(label, () => {
      validateSfnt(font);
    });
  }
});
