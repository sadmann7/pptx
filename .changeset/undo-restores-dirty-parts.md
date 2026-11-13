---
"@diceui/pptx-core": patch
---

Undoing an edit puts its parts back to how they were, so a deck undone to its original state saves those parts byte-for-byte instead of re-serializing them.
