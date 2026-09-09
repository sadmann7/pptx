---
"@diceui/pptx-core": minor
---

Add an `onError` option to `loadEmbeddedFonts` so an embedded font that cannot be used is observable instead of silent. Until now a part that failed to decode, or decoded bytes the `FontFace` API rejected, was skipped without a trace: the deck loaded, `onProgress` counted the part as done, and the affected text rendered with a fallback typeface. The callback reports the part's package path, the typefaces left without it, whether it went missing, failed to decode, or failed to register, and the decoder's message and `MtxErrorCode` where it has one. Loading behavior is unchanged, so one unusable font still never fails a deck.
