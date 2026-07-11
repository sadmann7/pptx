# Test fixtures

## Embedded font parts (`*.fntdata`)

MTX-compressed EOT font parts as PowerPoint embeds them, used as oracle
inputs for the internal MicroType Express decompressor (`src/fonts/mtx/`).

- `InstrumentSansSemiBold-regular.fntdata`: subset of Instrument Sans
- `SpaceGroteskSemiBold-bold.fntdata`: subset of Space Grotesk

Both typefaces are licensed under the [SIL Open Font License 1.1](https://openfontlicense.org)
. The subsets here exist solely for testing
font decoding and are not usable as installable fonts.
