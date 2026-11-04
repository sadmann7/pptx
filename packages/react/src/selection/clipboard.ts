import type { NodeClipboard, PresentationData } from "@diceui/pptx-core";

/**
 * The clipboard type a copy writes. It carries only a token: the shapes stay
 * in this page, since packaging their images is async and a copy event has to
 * set its data synchronously.
 */
export const CLIPBOARD_TYPE = "application/x-diceui-pptx";

/** The px each repeated paste onto the copied slide shifts by, as in PowerPoint. */
export const PASTE_OFFSET = 10;

interface Stash {
  token: string;
  clipboard: Promise<NodeClipboard>;
  presentation: PresentationData;
  slideId: string;
  pastes: number;
}

let stash: Stash | null = null;

export function stashNodes(
  presentation: PresentationData,
  slideId: string,
  clipboard: Promise<NodeClipboard>,
): string {
  const token = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  stash = { token, clipboard, presentation, slideId, pastes: 0 };
  return token;
}

/**
 * The stashed shapes for a pasted token, and how far to shift them: onto the
 * slide they came from, each paste lands a step further along.
 */
export function takeStash(
  token: string,
  presentation: PresentationData,
  slideId: string,
): { clipboard: Promise<NodeClipboard>; offset: number } | null {
  if (!stash || stash.token !== token) return null;
  const isSourceSlide = stash.presentation === presentation && stash.slideId === slideId;
  if (isSourceSlide) stash.pastes++;
  return { clipboard: stash.clipboard, offset: isSourceSlide ? stash.pastes * PASTE_OFFSET : 0 };
}
