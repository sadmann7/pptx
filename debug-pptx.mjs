import { readFileSync } from "fs";
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const JSZip = require("./packages/parser/node_modules/jszip");

const buf = readFileSync("c:/Users/sadman/Downloads/diceui-pptx-viewer-test-deck.pptx");
const zip = await JSZip.loadAsync(buf);

async function readXml(path) {
  const f = zip.file(path);
  if (!f) { console.log(`  [MISSING] ${path}`); return null; }
  return await f.async("string");
}

// 1. Dump slide3 (the geometry slide with shadow shape)
console.log("=== SLIDE 3 (geometry) ===");
const slide3 = await readXml("ppt/slides/slide3.xml");
// Find the "Drop shadow" shape
const shadowMatch = slide3?.match(/<p:sp>[\s\S]*?Drop shadow[\s\S]*?<\/p:sp>/);
if (shadowMatch) {
  console.log("\n--- Drop shadow shape XML ---");
  console.log(shadowMatch[0]);
} else {
  console.log("  [no 'Drop shadow' shape found by regex]");
  // Try to find any shape with effectLst or outerShdw
  const effectMatch = slide3?.match(/<p:sp>[\s\S]*?(effectLst|outerShdw|effectRef)[\s\S]*?<\/p:sp>/);
  if (effectMatch) {
    console.log("\n--- Shape with effect reference ---");
    console.log(effectMatch[0]);
  }
}

// Also search for effectRef or effectLst anywhere in slide3
console.log("\n--- effectRef occurrences in slide3 ---");
const effectRefMatches = slide3?.match(/effectRef[^>]*>/g);
console.log(effectRefMatches ?? "none");

console.log("\n--- effectLst occurrences in slide3 ---");
const effectLstMatches = slide3?.match(/<a:effectLst[\s\S]*?<\/a:effectLst>/g);
console.log(effectLstMatches ?? "none");

// 2. Theme effect styles
console.log("\n=== THEME ===");
const theme = await readXml("ppt/theme/theme1.xml");
const effectStyleLst = theme?.match(/<a:effectStyleLst[\s\S]*?<\/a:effectStyleLst>/);
if (effectStyleLst) {
  console.log("\n--- effectStyleLst ---");
  console.log(effectStyleLst[0]);
} else {
  console.log("  [no effectStyleLst]");
}

// 3. Slide1 master/layout - look for footer and divider
console.log("\n=== SLIDE 1 RELS ===");
const slide1Rels = await readXml("ppt/slides/_rels/slide1.xml.rels");
console.log(slide1Rels);

// Find the layout
const layoutMatch = slide1Rels?.match(/Target="([^"]*slideLayout[^"]*)"/);
const layoutPath = layoutMatch ? "ppt/slides/" + layoutMatch[1].replace("../", "") : null;
console.log("\nLayout path:", layoutPath);

if (layoutPath) {
  const layoutXml = await readXml(layoutPath.replace("ppt/slides/ppt/", "ppt/"));
  // Look for footer placeholder or connector/line
  const ftrMatch = layoutXml?.match(/<p:ph[^>]*type="ftr"[^>]*>/g);
  console.log("\n--- Footer placeholders in layout ---");
  console.log(ftrMatch ?? "none");

  // Connectors
  const cxnMatch = layoutXml?.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g);
  console.log("\n--- Connectors in layout ---");
  console.log(cxnMatch ?? "none");

  // Layout rels to find master
  const layoutRelsPath = layoutPath.replace("ppt/slides/ppt/", "ppt/").replace(/([^/]+)$/, "_rels/$1.rels");
  const layoutRels = await readXml(layoutRelsPath);
  const masterMatch = layoutRels?.match(/Target="([^"]*slideMaster[^"]*)"/);
  const masterPath = masterMatch ? "ppt/" + masterMatch[1].replace("../", "") : null;
  console.log("\nMaster path:", masterPath);

  if (masterPath) {
    const masterXml = await readXml(masterPath.replace("ppt/ppt/", "ppt/"));
    const masterFtr = masterXml?.match(/<p:ph[^>]*type="ftr"[^>]*>/g);
    console.log("\n--- Footer placeholders in master ---");
    console.log(masterFtr ?? "none");

    const masterCxn = masterXml?.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g);
    console.log("\n--- Connectors in master ---");
    console.log(masterCxn ?? "none");

    // Any line shapes
    const lineShapes = masterXml?.match(/<a:prstGeom prst="line"[\s\S]*?<\/p:sp>/g);
    console.log("\n--- Line shapes in master ---");
    console.log(lineShapes ?? "none");

    // Search for anything that looks like a divider (horizontal line near bottom)
    const allSp = masterXml?.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
    console.log("\n--- All shapes in master (names) ---");
    for (const sp of allSp) {
      const nameM = sp.match(/name="([^"]*)"/);
      const phM = sp.match(/<p:ph[^>]*>/);
      const geomM = sp.match(/prst="([^"]*)"/);
      console.log(`  name="${nameM?.[1]}" ph=${phM?.[0] ?? 'none'} geom=${geomM?.[1] ?? 'none'}`);
    }
  }
}

// 4. Check text cut-off: body properties on slide 2 (text fidelity)
console.log("\n=== SLIDE 2 (text fidelity) - body properties ===");
const slide2 = await readXml("ppt/slides/slide2.xml");
const bodyPrs = slide2?.match(/<a:bodyPr[^>]*\/?>/g);
console.log(bodyPrs);

// 5. Check slide 3 for the roundRect with shadow specifically
console.log("\n=== ALL sp shapes in slide3 with their names ===");
const allSp3 = slide3?.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
for (const sp of allSp3) {
  const nameM = sp.match(/name="([^"]*)"/);
  const hasEffect = sp.includes("effectLst") || sp.includes("effectRef") || sp.includes("outerShdw");
  const geomM = sp.match(/prst="([^"]*)"/);
  const styleRef = sp.match(/<a:effectRef idx="(\d+)"/);
  console.log(`  name="${nameM?.[1]}" geom=${geomM?.[1] ?? 'text'} hasEffect=${hasEffect} effectRefIdx=${styleRef?.[1] ?? 'none'}`);
  if (hasEffect || styleRef) {
    // Print the p:style section
    const styleSection = sp.match(/<p:style>[\s\S]*?<\/p:style>/);
    if (styleSection) console.log("    style:", styleSection[0]);
    const effectSection = sp.match(/<a:effectLst[\s\S]*?<\/a:effectLst>/);
    if (effectSection) console.log("    effectLst:", effectSection[0]);
  }
}
