/**
 * Writes the decks `/pg?data=worst` loads into apps/docs/public/assets/worst-case/:
 *
 *   pnpm exec tsx scripts/worst-case-decks.ts
 *
 * Keep every value one a real deck could hold.
 *
 * - worst.pptx  A4 portrait, long font names, extreme sizes, RTL, emoji, odd shapes
 * - one.pptx    a single slide
 * - many.pptx   300 slides
 */
import JSZip from "jszip";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { buildCustomPptx } from "../../core/src/tests/fixtures/fixture-extras";

const OUT_DIR = join(
  dirname(fileURLToPath(import.meta.url)),
  "..",
  "..",
  "..",
  "apps",
  "docs",
  "public",
  "assets",
  "worst-case",
);

const PX = 9525;

interface Run {
  text: string;
  font?: string;
  size?: number;
  bold?: boolean;
}

interface ShapeOptions {
  x: number;
  y: number;
  w: number;
  h: number;
  paragraphs?: Run[][];
  geometry?: string;
  fill?: string;
  rotation?: number;
  textBox?: boolean;
  bodyPr?: string;
}

let nextId = 2;

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function shape(options: ShapeOptions): string {
  const id = nextId++;
  const rot = options.rotation ? ` rot="${options.rotation * 60000}"` : "";
  const paragraphs = (options.paragraphs ?? [])
    .map((runs) => {
      const body = runs
        .map((run) => {
          const attrs = [
            'lang="en-US"',
            run.size ? `sz="${Math.round(run.size * 100)}"` : "",
            run.bold ? 'b="1"' : "",
          ]
            .filter(Boolean)
            .join(" ");
          const font = run.font
            ? `<a:latin typeface="${escapeXml(run.font)}"/>`
            : "";
          return `<a:r><a:rPr ${attrs}>${font}</a:rPr><a:t>${escapeXml(run.text)}</a:t></a:r>`;
        })
        .join("");
      return `<a:p>${body}</a:p>`;
    })
    .join("");
  const fill = options.fill
    ? `<a:solidFill><a:srgbClr val="${options.fill}"/></a:solidFill>`
    : "<a:noFill/>";
  const txBody = options.paragraphs
    ? `<p:txBody>${options.bodyPr ?? '<a:bodyPr wrap="square"/>'}<a:lstStyle/>${paragraphs}</p:txBody>`
    : "";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr${options.textBox ? ' txBox="1"' : ""}/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${rot}><a:off x="${Math.round(options.x * PX)}" y="${Math.round(options.y * PX)}"/><a:ext cx="${Math.round(options.w * PX)}" cy="${Math.round(options.h * PX)}"/></a:xfrm><a:prstGeom prst="${options.geometry ?? "rect"}"><a:avLst/></a:prstGeom>${fill}<a:ln w="12700"><a:solidFill><a:srgbClr val="1F2937"/></a:solidFill></a:ln></p:spPr>${txBody}</p:sp>`;
}

// Real family names as decks carry them, long static instances included.
const FONTS = [
  "Bricolage Grotesque 96pt Condensed ExtraBold",
  "Noto Sans Devanagari UI SemiCondensed",
  "IBM Plex Sans Condensed SemiBold",
  "Helvetica Neue LT Pro 57 Condensed",
  "Source Han Sans SC Heavy",
  "游ゴシック Medium",
  "Segoe UI Semibold",
  "Abadi Extra Light",
  "Arial Unicode MS",
  "Avenir Next LT Pro Demi",
  "Roboto Flex",
  "Inter Tight",
  "Playfair Display SC Black",
  "Libre Franklin Thin",
  "Atkinson Hyperlegible Next",
  "JetBrains Mono NL ExtraLight",
  "Times New Roman",
  "Calibri",
];

// A4 portrait, the size a deck exported from Word or a print workflow has.
const PORTRAIT = { cx: 7560310, cy: 10692130 };
const PORTRAIT_PX = { w: PORTRAIT.cx / PX, h: PORTRAIT.cy / PX };

function worstSlides(): string[] {
  const { w, h } = PORTRAIT_PX;
  const text = [
    shape({
      x: 40,
      y: 40,
      w: w - 80,
      h: 160,
      textBox: true,
      paragraphs: [
        [
          {
            text: "Rechtsschutzversicherungsgesellschaften",
            font: FONTS[0],
            size: 54,
            bold: true,
          },
        ],
      ],
    }),
    shape({
      x: 40,
      y: 220,
      w: w - 80,
      h: 60,
      textBox: true,
      paragraphs: [
        [
          {
            text: "https://northwind-industries-holdings.example.com/quarterly-reports/2026/q3/board-presentation-final-v12-approved.pdf",
            font: FONTS[1],
            size: 10.5,
          },
        ],
      ],
    }),
    shape({
      x: 40,
      y: 300,
      w: w - 80,
      h: 120,
      textBox: true,
      paragraphs: [
        [
          {
            text: "Báo cáo tài chính quý ba — Tiếng Việt",
            font: FONTS[2],
            size: 24,
          },
        ],
        [{ text: "تقرير الربع الثالث للمجلس", font: FONTS[4], size: 24 }],
        [{ text: "🚀 Launch day 🎉👩🏽‍💻", font: FONTS[5], size: 24 }],
      ],
    }),
    // Mixed sizes and fonts in one shape: the toolbar has no single value.
    shape({
      x: 40,
      y: 440,
      w: w - 80,
      h: 80,
      textBox: true,
      paragraphs: [
        [
          { text: "Mixed ", font: FONTS[3], size: 12 },
          { text: "sizes ", font: FONTS[6], size: 28, bold: true },
          { text: "and fonts", font: FONTS[7], size: 9 },
        ],
      ],
    }),
    // Autofit shrank a 166pt headline into a short box.
    shape({
      x: 40,
      y: 540,
      w: w - 80,
      h: 140,
      textBox: true,
      bodyPr:
        '<a:bodyPr wrap="square"><a:normAutofit fontScale="25000" lnSpcReduction="20000"/></a:bodyPr>',
      paragraphs: [
        [{ text: "Q3 2026", font: FONTS[8], size: 166, bold: true }],
      ],
    }),
  ];
  const geometry = [
    shape({ x: 60, y: 60, w: 6, h: 6, fill: "DC2626" }),
    shape({ x: 120, y: 200, w: 400, h: 0, geometry: "line" }),
    shape({
      x: -120,
      y: 300,
      w: 300,
      h: 160,
      fill: "2563EB",
      paragraphs: [[{ text: "Off the edge" }]],
    }),
    shape({ x: w - 200, y: -20, w: 160, h: 90, rotation: 45, fill: "16A34A" }),
    shape({ x: 0, y: 0, w, h, fill: "FEF3C7" }),
    shape({
      x: 200,
      y: 700,
      w: 200,
      h: 120,
      fill: "7C3AED",
      paragraphs: [[{ text: "Under nothing" }]],
    }),
  ];
  // The rest of the fonts, one line each, so the font list is as long as a
  // real brand deck's.
  const fontSlide = FONTS.slice(9).map((font, index) =>
    shape({
      x: 40,
      y: 40 + index * 50,
      w: w - 80,
      h: 44,
      textBox: true,
      paragraphs: [[{ text: font, font, size: 20 }]],
    }),
  );
  // Geometry slide: the full-slide shape sits behind the others, as a
  // background card does.
  return [
    text.join(""),
    [geometry[4], ...geometry.filter((_, i) => i !== 4)].join(""),
    fontSlide.join(""),
  ];
}

async function setSlideSize(
  buffer: ArrayBuffer,
  size: { cx: number; cy: number },
) {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("ppt/presentation.xml")!.async("string");
  zip.file(
    "ppt/presentation.xml",
    xml.replace(
      /<p:sldSz [^>]*\/>/,
      `<p:sldSz cx="${size.cx}" cy="${size.cy}"/>`,
    ),
  );
  return zip.generateAsync({ type: "uint8array" });
}

function titleSlide(index: number): string {
  return shape({
    x: 96,
    y: 96,
    w: 1088,
    h: 160,
    textBox: true,
    paragraphs: [[{ text: `Slide ${index + 1}`, size: 44, bold: true }]],
  });
}

mkdirSync(OUT_DIR, { recursive: true });

const worst = await buildCustomPptx({
  slides: worstSlides(),
  majorLatin: FONTS[0],
  minorLatin: FONTS[1],
});
writeFileSync(join(OUT_DIR, "worst.pptx"), await setSlideSize(worst, PORTRAIT));

nextId = 2;
writeFileSync(
  join(OUT_DIR, "one.pptx"),
  new Uint8Array(await buildCustomPptx({ slides: [titleSlide(0)] })),
);

nextId = 2;
writeFileSync(
  join(OUT_DIR, "many.pptx"),
  new Uint8Array(
    await buildCustomPptx({
      slides: Array.from({ length: 300 }, (_, i) => titleSlide(i)),
    }),
  ),
);

console.log(`Wrote worst.pptx, one.pptx, many.pptx to ${OUT_DIR}`);
