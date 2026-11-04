---
"@diceui/pptx": minor
---

Add draw tools to `useSelection`: `startDrawing({ type: "textBox" })` or `startDrawing({ type: "shape", geometry })` arms one, and a click on the slide places the text box or shape there, or a drag draws it to size. `drawTool` reports the armed tool. `Presentation.Marquee` previews the drawn box with `data-draw`. A text box left empty is removed once editing moves elsewhere, as in PowerPoint and Google Slides.
