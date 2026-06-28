import { readFileSync } from "fs";
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const JSZip = require("./packages/parser/node_modules/jszip");

const buf = readFileSync("c:/Users/sadman/Downloads/diceui-pptx-viewer-test-deck.pptx");
const zip = await JSZip.loadAsync(buf);

// List all files in the ZIP
console.log("=== ZIP file listing ===");
const files = Object.keys(zip.files).filter(f => !f.endsWith('/'));
for (const f of files) {
  if (f.includes('slideLayout') || f.includes('slideMaster') || f.includes('theme') || f.includes('_rels/slide')) {
    console.log(`  ${f}`);
  }
}

// Check layout XML
console.log("\n=== slideLayout2.xml ===");
const layout2 = await zip.file("ppt/slideLayouts/slideLayout2.xml")?.async("string");
if (!layout2) {
  console.log("  [MISSING]");
} else {
  // Find footer, date, sldNum, connectors
  const ftrMatch = layout2.match(/<p:ph[^>]*type="ftr"[^>]*>/g);
  console.log("Footer ph:", ftrMatch ?? "none");
  
  const dtMatch = layout2.match(/<p:ph[^>]*type="dt"[^>]*>/g);
  console.log("Date ph:", dtMatch ?? "none");
  
  const sldNumMatch = layout2.match(/<p:ph[^>]*type="sldNum"[^>]*>/g);
  console.log("Slide# ph:", sldNumMatch ?? "none");
  
  const cxnMatch = layout2.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g);
  console.log("Connectors:", cxnMatch ? `${cxnMatch.length} found` : "none");
  
  // All shapes with names
  const allSp = layout2.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
  console.log(`All shapes (${allSp.length}):`);
  for (const sp of allSp) {
    const nameM = sp.match(/name="([^"]*)"/);
    const phM = sp.match(/<p:ph([^>]*)>/);
    console.log(`  name="${nameM?.[1]}" ph=${phM?.[1]?.trim() ?? 'none'}`);
  }
}

// Check master XML  
console.log("\n=== Layout2 rels ===");
const layout2Rels = await zip.file("ppt/slideLayouts/_rels/slideLayout2.xml.rels")?.async("string");
console.log(layout2Rels ?? "[MISSING]");

const masterMatch = layout2Rels?.match(/Target="([^"]*)"/g);
console.log("Master targets:", masterMatch);

console.log("\n=== slideMaster1.xml shapes ===");
const master1 = await zip.file("ppt/slideMasters/slideMaster1.xml")?.async("string");
if (master1) {
  const allSp = master1.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
  console.log(`All shapes (${allSp.length}):`);
  for (const sp of allSp) {
    const nameM = sp.match(/name="([^"]*)"/);
    const phM = sp.match(/<p:ph([^>]*)>/);
    const geomM = sp.match(/prst="([^"]*)"/);
    console.log(`  name="${nameM?.[1]}" ph=${phM?.[1]?.trim() ?? 'none'} geom=${geomM?.[1] ?? 'text'}`);
    // For footer, print text content
    if (sp.includes('type="ftr"')) {
      const textMatch = sp.match(/<a:t>(.*?)<\/a:t>/g);
      console.log(`    text: ${textMatch}`);
    }
  }
  
  const cxnMatch = master1.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g);
  console.log(`\nConnectors: ${cxnMatch ? cxnMatch.length + ' found' : 'none'}`);
  if (cxnMatch) {
    for (const cxn of cxnMatch) {
      console.log(cxn.substring(0, 500));
    }
  }
}

// Trace our parser's parseColor on the shadow XML
console.log("\n=== Shadow color parsing trace ===");
// The shadow XML is:
// <a:outerShdw blurRad="25400" dist="50800" dir="2700000" algn="bl" rotWithShape="0">
//   <a:srgbClr val="CBD5E1"><a:alpha val="30000"/></a:srgbClr>
// </a:outerShdw>
// 
// After fast-xml-parser, the a:outerShdw node would look like:
// {
//   "@_blurRad": "25400",
//   "@_dist": "50800", 
//   "@_dir": "2700000",
//   "@_algn": "bl",
//   "@_rotWithShape": "0",
//   "a:srgbClr": { "@_val": "CBD5E1", "a:alpha": { "@_val": "30000" } }
// }
// parseColor(sn) checks for "a:srgbClr" in sn — this should work!
// But wait — let me check if our fast-xml-parser config strips prefixes...

// Let me check the xml.ts config
console.log("Check packages/parser/src/xml.ts for attributeNamePrefix and tagNameProcessor...");
