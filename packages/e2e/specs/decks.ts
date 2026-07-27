/** The fixture decks the specs run against, and how many slides each has. */
export interface Deck {
  /** File name without the .pptx extension, as served from fixtures/. */
  name: string;
  slides: number;
}

/**
 * Hand-built decks from scripts/generate-fixtures.ts: one narrow feature each,
 * with XML the structural specs can assert against exactly.
 */
export const GENERATED_DECKS: readonly Deck[] = [
  { name: "basic", slides: 3 },
  { name: "bom-rels", slides: 1 },
  { name: "nested-charts", slides: 2 },
  { name: "table-borders", slides: 2 },
  { name: "tables-groups", slides: 2 },
];

/**
 * Decks exported from an authoring tool: full themes and layouts, images,
 * gradients, charts, tables and mixed text. Their XML is nothing like the
 * minimal fixtures, which is where fidelity actually breaks down, so they are
 * only asserted against PowerPoint's own export (oracle.spec.ts) rather than
 * against hand-written expectations.
 */
export const EXPORTED_DECKS: readonly Deck[] = [
  { name: "geometry-of-attention", slides: 8 },
  { name: "internet-with-texture", slides: 8 },
  { name: "make-something-strange", slides: 8 },
  { name: "pocket-machines-sakura-chroma", slides: 8 },
  { name: "the-good-room-soft-editorial", slides: 8 },
  { name: "tiny-adventure-club", slides: 8 },
];

export const ALL_DECKS: readonly Deck[] = [...GENERATED_DECKS, ...EXPORTED_DECKS];
