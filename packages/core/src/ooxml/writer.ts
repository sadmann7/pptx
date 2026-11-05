import type { PresentationData } from "../model/presentation";
import type { PptxSaveOptions } from "./package";

/**
 * Serializes a presentation back to a .pptx archive. Untouched parts round-trip
 * byte-for-byte. Requires parsing with `keepSourcePackage: true`.
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
