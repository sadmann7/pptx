import { readFileSync } from "fs";
import { createRequire } from "module";
const require = createRequire(import.meta.url);

// Use tsx to import our TypeScript parser directly
const { parsePresentation } = require("./packages/parser/src/index.ts");

const buf = readFileSync("c:/Users/sadman/Downloads/diceui-pptx-viewer-test-deck.pptx");

console.log("Parsing...");
const pres = await parsePresentation(buf.buffer);

console.log(`Slides: ${pres.slides.length}`);
console.log(`Theme effectStyles: ${pres.theme.effectStyles?.length ?? 'none'}`);
if (pres.theme.effectStyles) {
  for (let i = 0; i < pres.theme.effectStyles.length; i++) {
    console.log(`  effectStyle[${i}]: ${JSON.stringify(pres.theme.effectStyles[i])}`);
  }
}

// Slide 3 (index 2) — the geometry slide
const slide3 = pres.slides[2];
console.log(`\n=== Slide 3 (${slide3.elements.length} elements) ===`);
for (const el of slide3.elements) {
  const name = el.name ?? el.id;
  const effects = el.effects ? JSON.stringify(el.effects) : 'none';
  const fill = el.fill ? el.fill.type : 'none';
  console.log(`  [${el.type}] id=${el.id} name="${name}" effects=${effects} fill=${fill}`);
  if (el.type === 'shape' && el.shapeType === 'roundRect' && el.effects) {
    console.log(`    >>> FOUND shadow shape! effects:`, JSON.stringify(el.effects, null, 2));
  }
}

// Slide 1 — check for footer/divider
const slide1 = pres.slides[0];
console.log(`\n=== Slide 1 (${slide1.elements.length} elements) ===`);
for (const el of slide1.elements) {
  const name = el.name ?? el.id;
  const ph = el.placeholder ? `${el.placeholder.type}/${el.placeholder.idx}` : 'none';
  console.log(`  [${el.type}] id=${el.id} name="${name}" ph=${ph}`);
  if (el.type === 'connector') {
    console.log(`    connector: shapeType=${el.shapeType} stroke=${JSON.stringify(el.stroke)}`);
  }
  if (el.type === 'text' && el.placeholder) {
    const text = el.paragraphs.map(p => p.runs.map(r => r.text ?? '').join('')).join(' | ');
    console.log(`    text: "${text}"`);
  }
}
