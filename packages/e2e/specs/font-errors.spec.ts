/**
 * An unusable embedded font has to be reported per part, not dropped. Only a
 * real browser has `FontFace`, which decides whether decoded bytes register.
 */
import { expect, test } from "@playwright/test";

// Enough embedded parts that several can break and several still load.
const DECK = "the-good-room-soft-editorial.pptx";
const PART_FACTOR = 2;

// MTX payloads send the deck to the worker pool in `worker` mode, where a failure
// has to survive a postMessage to reach the caller.
const MODES: readonly BenchMode[] = ["worker", "main"];

test.describe("embedded font failures", () => {
  for (const mode of MODES) {
    test(`are reported per part when decoding on the ${mode} path`, async ({
      page,
    }) => {
      const query = new URLSearchParams({
        file: DECK,
        mode,
        parts: String(PART_FACTOR),
        mtx: "1",
      });
      await page.goto(`/fonts.html?${query}`);
      await page.waitForFunction(
        () =>
          window.__fontsReady === true || window.__renderError !== undefined,
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

      expect(errors).toHaveLength(broken.length);
      for (const part of broken) {
        const reported = errors.filter((error) => error.path === part.path);
        expect(reported).toHaveLength(1);
        expect(reported[0]?.stage).toBe(part.stage);
        expect(reported[0]?.typefaces.sort()).toEqual(
          [...part.typefaces].sort(),
        );
        expect(reported[0]?.message).not.toBe("");
      }

      const reportFor = (kind: BrokenFontPart["kind"]) => {
        const part = broken.find((candidate) => candidate.kind === kind);
        return errors.find((error) => error.path === part?.path);
      };

      // An empty part the deck does carry is the decoder's to reject.
      expect(reportFor("empty")?.message).toBe("Font part is empty");

      // The decoder identifies a container it cannot read, and that code is
      // what survives the trip out of the worker in `worker` mode.
      expect(reportFor("undecodable")?.code).toBe("INVALID_EOT");

      // Bytes the decoder accepts can still be refused by `FontFace`. Two
      // families share this part, and one report covers both.
      const unregisterable = reportFor("unregisterable");
      expect(unregisterable?.typefaces.length).toBeGreaterThan(1);

      expect(faces).toBeGreaterThan(0);
    });
  }
});
