/**
 * An embedded font that cannot be used has to be reported, not dropped.
 *
 * This runs in a real browser because that is the only place the whole failure
 * surface exists: `FontFace` decides whether decoded bytes are a usable font,
 * and neither happy-dom nor jsdom implements it, so a unit test can reach the
 * decoder but never the registration that follows it.
 *
 * The harness breaks one part per kind of failure and leaves the rest intact,
 * so each run also proves reporting is per part rather than per deck.
 */
import { expect, test } from "@playwright/test";

/** Three embedded parts, doubled, so several can break and several still load. */
const DECK = "the-good-room-soft-editorial.pptx";
const PART_FACTOR = 2;

/**
 * MTX payloads, because a compressed part is what sends the deck to the worker
 * pool in `worker` mode: a failure then has to survive a postMessage to reach
 * the caller, which is a separate path from the fallback's.
 */
const MODES: readonly BenchMode[] = ["worker", "main"];

test.describe("embedded font failures", () => {
  for (const mode of MODES) {
    test(`are reported per part when decoding on the ${mode} path`, async ({ page }) => {
      const query = new URLSearchParams({
        file: DECK,
        mode,
        parts: String(PART_FACTOR),
        mtx: "1",
      });
      await page.goto(`/fonts.html?${query}`);
      await page.waitForFunction(
        () => window.__fontsReady === true || window.__renderError !== undefined,
      );
      const harnessError = await page.evaluate(() => window.__renderError);
      expect(harnessError).toBeUndefined();

      const result = await page.evaluate(() => window.__benchFontErrors?.());
      if (!result) throw new Error("harness did not expose __benchFontErrors");
      const { broken, errors, faces } = result;
      expect(broken.map((part) => part.kind)).toEqual([
        "missing",
        "empty",
        "undecodable",
        "unregisterable",
      ]);

      // One report per broken part, naming every family that lost it, and none
      // for the parts left alone.
      expect(errors).toHaveLength(broken.length);
      for (const part of broken) {
        const reported = errors.filter((error) => error.path === part.path);
        expect(reported).toHaveLength(1);
        expect(reported[0]?.stage).toBe(part.stage);
        expect(reported[0]?.typefaces.sort()).toEqual([...part.typefaces].sort());
        expect(reported[0]?.message).not.toBe("");
      }

      const reportFor = (kind: BrokenFontPart["kind"]) => {
        const part = broken.find((candidate) => candidate.kind === kind);
        return errors.find((error) => error.path === part?.path);
      };

      // A part the deck carries but left empty is the decoder's to reject, so
      // it reads as an unreadable payload and not as one the package lacks.
      expect(reportFor("empty")?.message).toBe("Font part is empty");

      // The decoder identifies a container it cannot read, and that code is
      // what survives the trip out of the worker in `worker` mode.
      expect(reportFor("undecodable")?.code).toBe("INVALID_EOT");

      // Bytes the decoder accepts can still be refused here, which is the
      // failure no amount of decode-level testing would catch. Two families
      // share this part, and one report covers both.
      const unregisterable = reportFor("unregisterable");
      expect(unregisterable?.typefaces.length).toBeGreaterThan(1);

      // The intact parts still registered, so reporting did not come at the
      // cost of loading the deck.
      expect(faces).toBeGreaterThan(0);
    });
  }
});
