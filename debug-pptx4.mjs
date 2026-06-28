import { readFileSync } from "fs";
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const JSZip = require("./packages/parser/node_modules/jszip");

const buf = readFileSync("c:/Users/sadman/Downloads/diceui-pptx-viewer-test-deck.pptx");
const zip = await JSZip.loadAsync(buf);

// Dump full layout2 XML
const layout2 = await zip.file("ppt/slideLayouts/slideLayout2.xml")?.async("string");
console.log("=== FULL slideLayout2.xml ===");
console.log(layout2);

console.log("\n=== FULL slideMaster1.xml ===");
const master1 = await zip.file("ppt/slideMasters/slideMaster1.xml")?.async("string");
console.log(master1);

// Also check slide 1 for the footer text
console.log("\n=== Searching for 'DiceUI' in all files ===");
for (const [name, file] of Object.entries(zip.files)) {
  if (file.dir) continue;
  if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
  const content = await file.async("string");
  if (content.includes('DiceUI') || content.includes('diceui') || content.includes('PPTX viewer test deck')) {
    console.log(`  FOUND in: ${name}`);
    const idx = content.indexOf('DiceUI');
    const idx2 = content.indexOf('PPTX viewer test deck');
    const start = Math.max(0, Math.min(idx >= 0 ? idx : 999999, idx2 >= 0 ? idx2 : 999999) - 200);
    console.log(`  Context: ...${content.substring(start, start + 500)}...`);
  }
}
