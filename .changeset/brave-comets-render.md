---
"@diceui/pptx-core": patch
---

Support embedded fonts with CFF outlines, and only use the font worker when a part actually needs MTX decompression. The MTX decoder previously required the TrueType `glyf` and `loca` tables, so decks carrying an OpenType/CFF font (`OTTO` sfnt version) failed to decode and fell back to a substitute font; such fonts now pass the `CFF ` table through and skip TrueType glyph reconstruction. Uncompressed EOT parts, which most non-PowerPoint producers emit, are cheap enough to unwrap inline and now avoid the worker and its lazily loaded chunk entirely. When no worker is available, the main-thread fallback yields on a 5ms time budget instead of once per font part, which cuts decode time substantially for decks with many embedded fonts while keeping frames responsive.
