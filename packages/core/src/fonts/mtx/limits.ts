/**
 * Resource ceilings applied while decoding.
 *
 * A malformed container can declare enormous sizes long before it produces
 * any output, so every allocation the decoder makes is checked against these
 * first.
 */
export interface DecoderLimits {
  /** Maximum size of any individual decompressed LZCOMP stream. */
  maxStreamBytes: number;
  /** Maximum size after optional LZCOMP run-length expansion. */
  maxExpandedStreamBytes: number;
  /** Maximum final SFNT size. */
  maxFontBytes: number;
  maxTables: number;
  maxGlyphs: number;
  maxPointsPerGlyph: number;
  maxInstructionsPerGlyph: number;
}

export const DEFAULT_LIMITS: Readonly<DecoderLimits> = Object.freeze({
  maxStreamBytes: 64 * 1024 * 1024,
  maxExpandedStreamBytes: 128 * 1024 * 1024,
  maxFontBytes: 128 * 1024 * 1024,
  maxTables: 256,
  maxGlyphs: 65_535,
  maxPointsPerGlyph: 1_000_000,
  maxInstructionsPerGlyph: 1_000_000,
});

/** Fill in the defaults for any limit a caller left out, rejecting nonsense. */
export function resolveLimits(input?: Partial<DecoderLimits>): DecoderLimits {
  const limits = { ...DEFAULT_LIMITS, ...input };
  for (const [name, value] of Object.entries(limits)) {
    if (!Number.isSafeInteger(value) || value <= 0) {
      throw new TypeError(`${name} must be a positive safe integer`);
    }
  }
  return limits;
}
