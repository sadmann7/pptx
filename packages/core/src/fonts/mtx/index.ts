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
import { fail, MtxError, type MtxErrorCode } from "./error";
import { resolveLimits, type DecoderLimits } from "./limits";
import { decompressLzcomp } from "./lzcomp";
import { buildSfnt, validateSfnt, type SfntContainer, type SfntTable } from "./sfnt";

export interface DecodeOptions {
  /** Undo the EOT XOR obfuscation before unpacking. */
  encrypted?: boolean;
  /** Whether the payload is MTX-compressed rather than a raw sfnt. */
  compressed?: boolean;
  limits?: Partial<DecoderLimits>;
  /** Called when the decoder drops something recoverable, such as `hdmx`. */
  onWarn?: (message: string) => void;
}

export interface UnpackedMtx {
  streams: [Uint8Array, Uint8Array, Uint8Array];
  sizes: [number, number, number];
}

/** Decompress the three LZCOMP blocks of an MTX payload into CTF streams. */
export function unpackMtx(data: Uint8Array, limits?: Partial<DecoderLimits>): UnpackedMtx {
  const size = data.length;
  if (size < 10) fail("INVALID_MTX", "Invalid MTX size");
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
  const resolved = resolveLimits(limits);
  const unpack = (start: number, end: number): Uint8Array =>
    decompressLzcomp(data.subarray(start, end), version, resolved);
  const streams: UnpackedMtx["streams"] = [
    unpack(10, offset2),
    unpack(offset2, offset3),
    unpack(offset3, size),
  ];
  return { streams, sizes: [streams[0].length, streams[1].length, streams[2].length] };
}

/**
 * Turns the font-data block of an EOT into an sfnt. An uncompressed block is
 * already one, so it is deobfuscated if needed and returned untouched.
 */
export function decodeMtx(data: Uint8Array, options: DecodeOptions = {}): Uint8Array {
  const encrypted = options.encrypted ?? false;
  const compressed = options.compressed ?? true;
  let decrypted = data;
  if (encrypted) {
    decrypted = data.slice();
    for (let i = 0; i < decrypted.length; i++) decrypted[i] = decrypted[i]! ^ 0x50;
  }
  if (!compressed) return decrypted;
  const limits = resolveLimits(options.limits);
  const unpacked = unpackMtx(decrypted, limits);
  const container = parseCtf(unpacked.streams, limits, options.onWarn);
  const font = buildSfnt(container, limits.maxFontBytes);
  validateSfnt(font);
  return font;
}

export function eotToTtf(
  data: Uint8Array,
  options: Omit<DecodeOptions, "compressed" | "encrypted"> = {},
): Uint8Array {
  const metadata = parseEotMetadata(data);
  const end = metadata.fontDataOffset + metadata.fontDataSize;
  return decodeMtx(data.subarray(metadata.fontDataOffset, end), {
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
export type { DecoderLimits, EotMetadata, EotVersion, MtxErrorCode, SfntContainer, SfntTable };
