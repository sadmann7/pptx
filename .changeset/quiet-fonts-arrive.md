---
"@diceui/pptx-core": patch
---

Load embedded fonts whatever extension their part uses. Only `ppt/fonts/*.fntdata` parts were read, which is what PowerPoint writes, so decks from producers that pick another extension (Walnut Exporter writes `.dat`) had every embedded font reported missing and rendered with a fallback typeface. Any part under `ppt/fonts/` is now read.
