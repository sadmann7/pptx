/**
 * .pptx writer — writes a presentation opened with `keepPackage: true` back
 * to an OOXML package. Untouched parts round-trip byte-for-byte; parts marked
 * dirty on the retained package are re-serialized from their live XML.
 */

import type { PresentationData } from "../model/presentation";
import type { PptxSaveOptions } from "./package";

/**
 * Serialize a presentation back to a .pptx archive.
 *
 * Requires the presentation to have been parsed with `keepPackage: true`
 * (see `ZipParseOptions.keepPackage`), which retains the source package for
 * round-trip fidelity.
 */
export async function writePptx(
  presentation: PresentationData,
  options?: PptxSaveOptions,
): Promise<Uint8Array> {
  if (!presentation.pkg) {
    throw new Error(
      "writePptx: presentation was parsed without package retention. " +
        "Parse the zip with { keepPackage: true } to enable saving.",
    );
  }
  return presentation.pkg.save(options);
}
