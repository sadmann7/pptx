---
"@diceui/pptx": minor
---

Make the editor easier to use with a keyboard and screen reader. Tab and Shift+Tab move through a slide's shapes from `Presentation.Selection` (focus moves on past the last one), and a polite live region announces what's selected or being edited. The overlay is named "Slide editor", the thumbnail list "Slides", and each thumbnail "Slide N"; pass `aria-label` to override.

Add `duplicate()`, `remove()`, and `reorder(order)` to `useSelection`, matching Ctrl/Cmd+D, Delete, and Ctrl/Cmd+] and [.
