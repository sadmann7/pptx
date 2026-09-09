---
"@diceui/pptx-core": minor
---

Add an `onError` option to `loadEmbeddedFonts` so an embedded font that cannot be used is reported instead of silently falling back. Each report names the part, the typefaces left without it, whether it went missing, failed to decode, or was rejected by the `FontFace` API, and the decoder's message.
