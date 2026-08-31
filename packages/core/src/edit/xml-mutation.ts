/**
 * Low-level XML mutation helpers for edit operations.
 *
 * All structural mutations (inserting/removing child elements) go through
 * these helpers so the xml-parser child-index cache is invalidated for the
 * touched parent. Attribute and text mutations don't affect the index.
 */

import { invalidateChildIndex } from "../ooxml/xml";

export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
export const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

const DEFAULT_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

let sharedSerializer: XMLSerializer | undefined;

/** Insert `el` into `parent` before `ref` (append when `ref` is null). */
export function insertChild(parent: Element, el: Element, ref: Element | null): void {
  parent.insertBefore(el, ref);
  invalidateChildIndex(parent);
}

/** Remove `el` from its parent. Returns the parent for later re-insertion. */
export function removeChild(el: Element): Element | null {
  const parent = el.parentElement;
  el.remove();
  if (parent) invalidateChildIndex(parent);
  return parent;
}

/** Set an attribute, or remove it when `value` is undefined. */
export function setOrRemoveAttr(el: Element, name: string, value: string | undefined): void {
  if (value === undefined) {
    el.removeAttribute(name);
  } else {
    el.setAttribute(name, value);
  }
}

/**
 * Serializes a mutated XML document root back to part text, preserving the
 * original part's `<?xml ...?>` declaration.
 */
export function serializePartText(root: Element, originalText: string): string {
  sharedSerializer ??= new XMLSerializer();
  const declaration = /^<\?xml.*?\?>\r?\n?/.exec(originalText)?.[0] ?? DEFAULT_XML_DECLARATION;
  return declaration + sharedSerializer.serializeToString(root);
}
