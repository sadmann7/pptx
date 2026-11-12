import type { PresentationData } from "../model/presentation";
import type { PptxSaveOptions } from "./package";

/**
 * Serializes a presentation parsed with `keepSourcePackage: true` back to a
 * .pptx archive, copying untouched parts byte for byte.
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
