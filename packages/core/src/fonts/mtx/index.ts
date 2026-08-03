import { Reader } from "./binary";
import { parseCtf } from "./ctf";
import {
  canLegallyEdit,
  parseEotMetadata,
  TTEMBED_SUBSET,
  TTEMBED_TTCOMPRESSED,
  TTEMBED_XORENCRYPTDATA,
  type EotMetadata,
  type EotVersion,
} from "./eot";
import { fail, MtxError, type MtxErrorCode } from "./errors";
import { decompressLzcomp } from "./lzcomp";
import { resolveLimits, type DecodeOptions, type DecoderLimits } from "./options";
import { buildSfnt, validateSfnt, type SfntContainer, type SfntTable } from "./sfnt";

export interface UnpackedMtx {
  streams: [Uint8Array, Uint8Array, Uint8Array];
  sizes: [number, number, number];
}

export function unpackMtx(
  data: Uint8Array,
  size = data.length,
  limitOverrides?: Partial<DecoderLimits>,
): UnpackedMtx {
  if (!Number.isSafeInteger(size) || size < 10 || size > data.length)
    fail("INVALID_MTX", "Invalid MTX size");
  const reader = new Reader(data, 0, size);
  const version = reader.u8();
  // Version 3 is MTX 1.0. Version 1 predates the run-length flag bit and is
  // still produced by older embedders; decompressLzcomp branches on it.
  if (version !== 1 && version !== 3) fail("UNSUPPORTED", `Unsupported MTX version ${version}`);
  reader.u24be();
  const offset2 = reader.u24be();
  const offset3 = reader.u24be();
  if (offset2 < 10 || offset3 < offset2 || offset3 > size)
    fail("INVALID_MTX", "Invalid MTX block offsets");
  const limits = resolveLimits(limitOverrides);
  const compressed: [Uint8Array, Uint8Array, Uint8Array] = [
    data.subarray(10, offset2),
    data.subarray(offset2, offset3),
    data.subarray(offset3, size),
  ];
  const streams = compressed.map((block) =>
    decompressLzcomp(block, version, limits),
  ) as UnpackedMtx["streams"];
  return { streams, sizes: [streams[0].length, streams[1].length, streams[2].length] };
}

export function decompressMtx(fontData: Uint8Array, options: DecodeOptions = {}): Uint8Array {
  const encrypted = options.encrypted ?? false;
  const compressed = options.compressed ?? true;
  let data = fontData;
  if (encrypted) {
    data = fontData.slice();
    for (let i = 0; i < data.length; i++) data[i] = data[i]! ^ 0x50;
  }
  if (!compressed) return data;
  const limits = resolveLimits(options.limits);
  const unpacked = unpackMtx(data, data.length, limits);
  const container = parseCtf(unpacked.streams, limits, options.onWarn);
  const font = buildSfnt(container, limits.maxFontBytes);
  validateSfnt(font);
  return font;
}

export function eotToTtf(
  bytes: Uint8Array,
  options: Omit<DecodeOptions, "compressed" | "encrypted"> = {},
): Uint8Array {
  const metadata = parseEotMetadata(bytes);
  const end = metadata.fontDataOffset + metadata.fontDataSize;
  return decompressMtx(bytes.subarray(metadata.fontDataOffset, end), {
    ...options,
    compressed: metadata.compressed,
    encrypted: metadata.encrypted,
  });
}

export {
  canLegallyEdit,
  MtxError,
  parseCtf,
  parseEotMetadata,
  TTEMBED_SUBSET,
  TTEMBED_TTCOMPRESSED,
  TTEMBED_XORENCRYPTDATA,
  validateSfnt,
};
export type {
  DecodeOptions,
  DecoderLimits,
  EotMetadata,
  EotVersion,
  MtxErrorCode,
  SfntContainer,
  SfntTable,
};
