import { Reader } from "./binary";
import { fail } from "./error";

export const TTEMBED_SUBSET = 0x00000001;
export const TTEMBED_TTCOMPRESSED = 0x00000004;
export const TTEMBED_XORENCRYPTDATA = 0x10000000;

export type EotVersion = 1 | 2 | 3;

export interface EotMetadata {
  version: EotVersion;
  flags: number;
  panose: Uint8Array;
  charset: number;
  italic: boolean;
  weight: number;
  permissions: number;
  unicodeRange: [number, number, number, number];
  codePageRange: [number, number];
  checkSumAdjustment: number;
  familyName: string;
  styleName: string;
  versionName: string;
  fullName: string;
  rootString: string;
  totalSize: number;
  fontDataSize: number;
  fontDataOffset: number;
  compressed: boolean;
  encrypted: boolean;
  badVersion: boolean;
}

const VERSION_MAGIC: Record<number, EotVersion | undefined> = {
  0x00010000: 1,
  0x00020001: 2,
  0x00020002: 3,
};

function decodeUtf16(bytes: Uint8Array): string {
  if (bytes.length & 1) fail("INVALID_EOT", "UTF-16 EOT string has an odd byte length");
  const text = new TextDecoder("utf-16le", { fatal: false }).decode(bytes);
  let end = text.length;
  while (end > 0 && text.charCodeAt(end - 1) === 0) end--;
  return text.slice(0, end);
}

function readString(reader: Reader): string {
  const size = reader.u16le();
  return decodeUtf16(reader.bytes(size));
}

function parseTail(bytes: Uint8Array, start: number, end: number, version: EotVersion): string {
  const reader = new Reader(bytes, start, end);
  let rootString = "";
  if (version >= 2) {
    reader.u16le();
    rootString = readString(reader);
  }
  if (version >= 3) {
    reader.u32le();
    reader.u32le();
    reader.u16le();
    const signatureSize = reader.u16le();
    reader.skip(signatureSize);
    reader.u32le();
    const eudcSize = reader.u32le();
    reader.skip(eudcSize);
  }
  if (reader.pos !== end) fail("INVALID_EOT", `Version ${version} header does not end at FontData`);
  return rootString;
}

export function parseEotMetadata(bytes: Uint8Array): EotMetadata {
  const reader = new Reader(bytes);
  if (bytes.length < 82) fail("INVALID_EOT", "EOT header is truncated");
  const totalSize = reader.u32le();
  const fontDataSize = reader.u32le();
  const rawVersion = reader.u32le();
  const declaredVersion = VERSION_MAGIC[rawVersion];
  if (!declaredVersion)
    fail("INVALID_EOT", `Unsupported EOT version 0x${rawVersion.toString(16).padStart(8, "0")}`);
  const flags = reader.u32le();
  const panose = reader.bytes(10).slice();
  const charset = reader.u8();
  const italic = reader.u8() !== 0;
  const weight = reader.u32le();
  const permissions = reader.u16le();
  if (reader.u16le() !== 0x504c) fail("INVALID_EOT", "Invalid EOT magic number");
  const unicodeRange: [number, number, number, number] = [
    reader.u32le(),
    reader.u32le(),
    reader.u32le(),
    reader.u32le(),
  ];
  const codePageRange: [number, number] = [reader.u32le(), reader.u32le()];
  const checkSumAdjustment = reader.u32le();
  reader.skip(16);
  reader.u16le();
  const familyName = readString(reader);
  reader.u16le();
  const styleName = readString(reader);
  reader.u16le();
  const versionName = readString(reader);
  reader.u16le();
  const fullName = readString(reader);

  if (totalSize > bytes.length || totalSize < fontDataSize || fontDataSize === 0) {
    fail("INVALID_EOT", "Inconsistent EOT or FontData size");
  }
  const fontDataOffset = totalSize - fontDataSize;
  if (fontDataOffset < reader.pos) fail("INVALID_EOT", "EOT strings overlap FontData");

  // Historical embedders sometimes stamped a version that disagrees with the
  // tail they actually wrote. parseTail only accepts a layout that lands
  // exactly on FontData, so trying the others is unambiguous.
  let parsedVersion: EotVersion | undefined;
  let rootString = "";
  for (const candidate of [declaredVersion, 1, 2, 3] as const) {
    if (parsedVersion) break;
    try {
      rootString = parseTail(bytes, reader.pos, fontDataOffset, candidate);
      parsedVersion = candidate;
    } catch {
      // Fall through to the next known layout.
    }
  }
  if (!parsedVersion)
    fail("INVALID_EOT", "No EOT header layout matches the declared FontData offset");

  return {
    version: parsedVersion,
    flags,
    panose,
    charset,
    italic,
    weight,
    permissions,
    unicodeRange,
    codePageRange,
    checkSumAdjustment,
    familyName,
    styleName,
    versionName,
    fullName,
    rootString,
    totalSize,
    fontDataSize,
    fontDataOffset,
    compressed: (flags & TTEMBED_TTCOMPRESSED) !== 0,
    encrypted: (flags & TTEMBED_XORENCRYPTDATA) !== 0,
    badVersion: parsedVersion !== declaredVersion,
  };
}

export function canLegallyEdit(metadata: Pick<EotMetadata, "permissions">): boolean {
  return metadata.permissions === 0 || (metadata.permissions & 0x0008) !== 0;
}
