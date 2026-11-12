import type { PresentationData } from "./presentation";

export interface PresentationFonts {
  /** The theme's heading and body fonts. */
  theme: { heading?: string; body?: string };
  /** Fonts embedded in the file, which render as the deck was authored. */
  embedded: string[];
  /** Other fonts the text on loaded slides uses, by name. */
  used: string[];
}

function collectTypefaces(element: Element, into: Set<string>): void {
  if (element.localName === "latin") {
    const typeface = element.getAttribute("typeface");
    // `+mj-lt` and `+mn-lt` point at the theme fonts, listed separately.
    if (typeface && !typeface.startsWith("+")) into.add(typeface);
  }
  for (const child of element.children) collectTypefaces(child, into);
}

/**
 * Lists the fonts a font picker should offer for a deck, skipping slides not yet loaded
 * with `materializeSlide`.
 */
export function getPresentationFonts(pres: PresentationData): PresentationFonts {
  const theme = pres.themes.values().next().value;
  const heading = theme?.majorFont.latin || undefined;
  const body = theme?.minorFont.latin || undefined;
  const embedded = [...new Set((pres.embeddedFonts ?? []).map((font) => font.typeface))];

  const typefaces = new Set<string>();
  for (const slide of pres.slides) {
    if (!slide.nodesMaterialized) continue;
    for (const node of slide.nodes) {
      if (node.source.element) collectTypefaces(node.source.element, typefaces);
    }
  }
  const listed = new Set([heading, body, ...embedded]);
  const used = [...typefaces].filter((typeface) => !listed.has(typeface)).sort();

  return { theme: { heading, body }, embedded, used };
}
