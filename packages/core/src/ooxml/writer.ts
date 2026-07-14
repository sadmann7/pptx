/**
 * Writes a presentation (opened with `keepSourcePackage: true`) back to a .pptx
 * OOXML package. Untouched parts round-trip byte-for-byte; parts marked
 * dirty on the retained sourcePackage are re-serialized from their live XML.
 */

import type { PresentationData } from "../model/presentation";
import type { PptxSaveOptions } from "./package";

/**
 * Serialize a presentation back to a .pptx archive.
 *
 * Requires the presentation to have been parsed with `keepSourcePackage: true`
 * (see `PptxReadOptions.keepSourcePackage`), which retains the source package for
 * round-trip fidelity.
 */
export async function writePptx(
  presentation: PresentationData,
  options?: PptxSaveOptions,
): Promise<Uint8Array> {
  if (!presentation.sourcePackage) {
    throw new Error(
      "writePptx: presentation was parsed without package retention. " +
        "Parse the zip with { keepSourcePackage: true } to enable saving.",
    );
  }
  return presentation.sourcePackage.save(options);
}
