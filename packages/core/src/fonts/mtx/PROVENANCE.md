# Clean-room provenance for the MTX/EOT decoder

## Implementation boundary

Every file in this directory was written independently for this project. Their
format behavior was derived from:

1. the W3C Member Submission _MicroType Express (MTX) Font Format_, 5 March
   2008: <https://www.w3.org/Submission/2008/SUBM-MTX-20080305/>;
2. the W3C Member Submission _Embedded OpenType (EOT) File Format_, 5 March
   2008: <https://www.w3.org/Submission/2008/SUBM-EOT-20080305/>;
3. the OpenType/SFNT file-format rules needed to serialize and validate the
   reconstructed font.

No implementation source from `ChristopherVR/mtx-decompressor`, libeot, or any
other MPL-licensed decoder was copied, translated, or adapted. Decoder outputs
were compared against a prior MPL-licensed decoder purely as a black-box
conformance check, and only for the uncompressed EOT path.

## Patent and trademark position

W3C's standards-licensing page records Monotype Imaging's no-charge,
royalty-free grant for implementing the MTX specification:
<https://www.w3.org/standards/licensing/>. The MTX submission itself identifies
`MicroType` as a Monotype Imaging trademark. This project is not affiliated with
or endorsed by Monotype Imaging, Microsoft, or W3C.

## License

This directory is covered by the repository's Apache License 2.0. That grant
covers the implementation only; it conveys no rights in third-party fonts a
consumer chooses to decode, nor in their embedding permissions.

## Divergences from the specification text

- **CTF glyph data (spec 5.8).** The table nests `pushCount`/`codeSize` inside
  the branch for simple glyphs _without_ an explicit bounding box. Glyphs with
  `numContours == 0x7FFF` and composite glyphs with `WE_HAVE_INSTRUCTIONS` also
  carry them; the nesting is a typesetting artifact.
- **Contourless glyphs.** A CTF glyph with `numContours == 0` consumes nothing
  further from any stream and is emitted as an empty `loca` range.
- **Reconstructed `glyf` encoding.** The spec fixes the CTF representation, not
  the TrueType one. This decoder emits repeat-compressed point flags and the
  short `PUSHB`/`PUSHW` forms wherever they apply, so the output is smaller than
  (but semantically identical to) a naive reconstruction. Table checksums and
  `head.checksumAdjustment` are recomputed from the bytes actually written.
- **`loca` format.** `head.indexToLocFormat` is rewritten to the long form when
  the reconstructed `glyf` exceeds the 131070-byte reach of short offsets,
  rather than silently truncating offsets.
