/**
 * An embedded font that cannot be used has to be reported, not dropped.
 *
 * This runs in a real browser because that is the only place the whole failure
 * surface exists: `FontFace` decides whether decoded bytes are a usable font,
 * and neither happy-dom nor jsdom implements it, so a unit test can reach the
 * decoder but never the registration that follows it.
 *
 * The harness breaks one part per kind of failure (dropped, undecodable, and
 * decodable but not a font the browser accepts) and leaves the rest intact, so
 * each run also proves reporting is per part rather than per deck.
 */
import { expect, test } from "@playwright/test";

/** Three embedded parts, doubled, so three can break and three still load. */
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
      expect(broken).toHaveLength(3);

      // One report per broken part, and none for the parts left alone.
      expect(errors).toHaveLength(broken.length);
      for (const { path, stage } of broken) {
        const reported = errors.filter((error) => error.path === path);
        expect(reported).toHaveLength(1);
        expect(reported[0]?.stage).toBe(stage);
        // Something to act on: which typefaces lost their font, and why.
        expect(reported[0]?.typefaces.length).toBeGreaterThan(0);
        expect(reported[0]?.message).not.toBe("");
      }

      // The decoder identifies a container it cannot read, and that code is
      // what survives the trip out of the worker in `worker` mode.
      const undecodable = errors.find((error) => error.stage === "decode");
      expect(undecodable?.code).toBe("INVALID_EOT");

      // A part the decoder accepts can still be refused here, which is the
      // failure no amount of decode-level testing would catch.
      const unregisterable = errors.find((error) => error.stage === "register");
      expect(unregisterable?.typefaces).toHaveLength(1);

      // The intact parts still registered, so reporting did not come at the
      // cost of loading the deck.
      expect(faces).toBeGreaterThan(0);
    });
  }
});
