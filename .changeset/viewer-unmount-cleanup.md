---
"@diceui/pptx": patch
---

Unmounting a viewer now frees the active slide's picture URLs and removes the deck's embedded fonts from `document.fonts`, including when it unmounts mid-load. `Presentation.Root` resets the store it creates, and `useCreatePresentationStore` resets its store when the calling component unmounts.
