---
"@diceui/pptx-core": minor
"@diceui/pptx": minor
---

Add copy and paste of shapes: `copyNodes()` packages shapes with the images they use, and the `pasteNodes` operation pastes them into any deck, reusing images it already has. `Presentation.Selection` handles Ctrl/Cmd+C, X, and V, within a deck or between decks on the page, and pastes image files as pictures.

Add the `addTextBox` and `setOutline` edit operations, `getPresentationFonts()` for font pickers, and `editText()` and `insertTextBox()` on `useSelection`. Rotated shapes in a multi-selection now resize along their own axes, and the first keystroke typed into an empty text box or placeholder can be undone.
