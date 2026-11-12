export interface Deck {
  name: string;
  slides: number;
}

// One narrow feature each, with XML the structural specs can assert exactly.
export const GENERATED_DECKS: readonly Deck[] = [
  { name: "basic", slides: 3 },
  { name: "bom-rels", slides: 1 },
  { name: "nested-charts", slides: 2 },
  { name: "table-borders", slides: 2 },
  { name: "tables-groups", slides: 2 },
];

// Their XML is where fidelity breaks down, so they're only checked against
// PowerPoint's own export, not hand-written expectations.
export const EXPORTED_DECKS: readonly Deck[] = [
  { name: "adventure-club-pin-and-paper", slides: 8 },
  { name: "geometry-of-attention-cartesian", slides: 8 },
  { name: "internet-with-texture-broadside", slides: 8 },
  { name: "make-something-strange-creative-mode", slides: 8 },
  { name: "pocket-machines-sakura-chroma", slides: 8 },
  { name: "the-good-room-soft-editorial", slides: 8 },
];

export const ALL_DECKS: readonly Deck[] = [
  ...GENERATED_DECKS,
  ...EXPORTED_DECKS,
];
