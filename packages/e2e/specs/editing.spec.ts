/**
 * A formatting toolbar outside the slide, pressed with a real mouse.
 *
 * Unit tests fire synthetic events at happy-dom, which has no focus model to
 * speak of; what these check only shows up in a browser: a toolbar press moves
 * focus and the document selection out of the text, and editing still has to
 * survive it and format the text that was selected.
 */
import { expect, test, type Page } from "@playwright/test";

// macOS moves to the line's edges with Cmd+arrows; Home and End do nothing there.
const LINE_START = process.platform === "darwin" ? "Meta+ArrowLeft" : "Home";
const LINE_END = process.platform === "darwin" ? "Meta+ArrowRight" : "End";

/** `[text, rPr attributes]` for each run of the first shape's first paragraph. */
function readRuns(page: Page) {
  return page.evaluate(() => {
    const slide = window.__editStore!.getActiveSlide()!;
    const shape = slide.nodes[0] as unknown as {
      textBody: { paragraphs: { runs: { text: string; properties?: { element?: Element } }[] }[] };
    };
    return shape.textBody.paragraphs[0].runs.map((run) => {
      const element = run.properties?.element;
      return {
        text: run.text,
        i: element?.getAttribute("i") ?? null,
        u: element?.getAttribute("u") ?? null,
        color: element?.getElementsByTagNameNS("*", "srgbClr")[0]?.getAttribute("val") ?? null,
      };
    });
  });
}

async function open(page: Page) {
  await page.goto("/editing.html");
  await page.waitForFunction(() => window.__editingReady === true);
  const overlay = page.locator("[data-pptx-selection]");
  await expect(overlay).toBeVisible();

  // The selection overlay covers the shapes and hit-tests what is under the
  // pointer, so press where the shape is rather than on its element.
  const box = (await page.locator('[data-pptx-node-id="2"]').boundingBox())!;
  const shape = {
    click: () => page.mouse.click(box.x + box.width / 2, box.y + box.height / 2),
  };
  return { overlay, shape };
}

test.describe("editing toolbar", () => {
  test("keeps the selection through toolbar and portalled popover presses", async ({ page }) => {
    const { overlay, shape } = await open(page);

    await shape.click();
    await expect(overlay).toHaveAttribute("data-mode", "selected");

    await page.getByTestId("underline").click();
    await expect(overlay).toHaveAttribute("data-mode", "selected");
    await expect
      .poll(() => readRuns(page))
      .toEqual([{ text: "Slide one", i: null, u: "sng", color: "FFFFFF" }]);

    await page.getByTestId("swatch").click();
    await expect(overlay).toHaveAttribute("data-mode", "selected");
    await expect.poll(async () => (await readRuns(page))[0].color).toBe("00FF00");

    // Undo from the toolbar's focus applies to the deck.
    await page.keyboard.press("ControlOrMeta+z");
    await expect.poll(async () => (await readRuns(page))[0].color).toBe("FFFFFF");

    await page.getByTestId("outside").click();
    await expect(overlay).toHaveAttribute("data-mode", "idle");
  });

  test("formats the selected text and keeps editing it", async ({ page }) => {
    const { overlay, shape } = await open(page);

    await shape.click();
    await page.keyboard.press("Enter");
    await expect(overlay).toHaveAttribute("data-mode", "text");

    // Select "one".
    await page.keyboard.press(LINE_END);
    for (let i = 0; i < 3; i++) await page.keyboard.press("Shift+ArrowLeft");

    await page.getByTestId("italic").click();
    await expect
      .poll(() => readRuns(page))
      .toEqual([
        { text: "Slide ", i: null, u: null, color: "FFFFFF" },
        { text: "one", i: "1", u: null, color: "FFFFFF" },
      ]);
    await expect(overlay).toHaveAttribute("data-mode", "text");

    // Focus and the selection come back to the text once the slide has
    // re-rendered, a frame or two after the edit lands; typing goes on from it.
    await expect
      .poll(() =>
        page.evaluate(() => ({
          editing: (document.activeElement as HTMLElement | null)?.isContentEditable ?? false,
          selected: window.getSelection()?.toString(),
        })),
      )
      .toEqual({ editing: true, selected: "one" });
    await page.keyboard.press("ArrowRight");
    await page.keyboard.type("!");
    await page.keyboard.press("Escape");
    await expect(overlay).toHaveAttribute("data-mode", "selected");
    await expect
      .poll(async () => (await readRuns(page)).map((run) => run.text).join(""))
      .toBe("Slide one!");
  });

  test("formats the word at the caret", async ({ page }) => {
    const { overlay, shape } = await open(page);

    await shape.click();
    await page.keyboard.press("Enter");
    await expect(overlay).toHaveAttribute("data-mode", "text");
    await page.keyboard.press(LINE_START);
    await page.keyboard.press("ArrowRight");

    await page.getByTestId("underline").click();
    await expect
      .poll(() => readRuns(page))
      .toEqual([
        { text: "Slide", i: null, u: "sng", color: "FFFFFF" },
        { text: " one", i: null, u: null, color: "FFFFFF" },
      ]);
    await expect(overlay).toHaveAttribute("data-mode", "text");
  });

  test("undoes the keystroke that started text editing", async ({ page }) => {
    const { overlay, shape } = await open(page);
    const text = page.locator('[data-pptx-node-id="2"] [data-pptx-p]').first();

    // Typing on a selected shape starts editing with that keystroke.
    await shape.click();
    await expect(overlay).toHaveAttribute("data-mode", "selected");
    await page.keyboard.type("X");
    await expect(overlay).toHaveAttribute("data-mode", "text");
    await expect(text).toHaveText("Slide oneX");

    await page.keyboard.press("ControlOrMeta+z");
    await expect(text).toHaveText("Slide one");
    await expect(overlay).toHaveAttribute("data-mode", "text");
  });
});
