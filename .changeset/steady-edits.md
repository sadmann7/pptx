---
"@diceui/pptx-core": patch
"@diceui/pptx": patch
---

Run edits, undo, and redo one at a time, so overlapping edits can't interleave. Copies keep connector links to shapes copied with them, never reuse an image part, and keep OLE and other wrapped objects whole. Empty paragraphs show the bold, italic, underline, and strikethrough they will type with, `textStyle.fontSize` reports the size the file stores rather than the autofit-shrunk one, and line breaks survive text editing.
