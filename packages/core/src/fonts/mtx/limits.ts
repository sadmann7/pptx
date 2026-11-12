// A malformed container can declare enormous sizes before producing any output, so
// every allocation the decoder makes is checked against these first.
export interface DecoderLimits {
  maxStreamBytes: number;
  // After LZCOMP run-length expansion.
  maxExpandedStreamBytes: number;
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

export function resolveLimits(input?: Partial<DecoderLimits>): DecoderLimits {
  const limits = { ...DEFAULT_LIMITS, ...input };
  for (const [name, value] of Object.entries(limits)) {
    if (!Number.isSafeInteger(value) || value <= 0) {
      throw new TypeError(`${name} must be a positive safe integer`);
    }
  }
  return limits;
}
