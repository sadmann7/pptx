// A separate entry point so only consumers that inject embedded fonts load the decoder.

export { decodeEmbeddedFont, decodeEmbeddedFontPart } from "./decode";
export type { FontDecodeResult } from "./decode";
export { deobfuscateFont } from "./deobfuscate";
export { findPriorityTypefaces, loadEmbeddedFonts } from "./loader";
export type {
  EmbeddedFontError,
  EmbeddedFontsHandle,
  LoadEmbeddedFontsOptions,
} from "./loader";
export type { MtxErrorCode } from "./mtx";
