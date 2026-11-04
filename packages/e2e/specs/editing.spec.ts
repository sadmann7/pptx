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

test.describe("snapping", () => {
  /** The shape's committed x, and where its center would sit on the slide's center. */
  function readGeometry(page: Page) {
    return page.evaluate(() => {
      const store = window.__editStore!;
      const node = store.getActiveSlide()!.nodes[0];
      const width = store.getState().presentation!.width;
      return { x: node.position.x, centeredX: width / 2 - node.size.w / 2 };
    });
  }

  /** Drags shape 2 so its center ends `offset` px right of the slide's center. */
  async function dragNearCenter(page: Page, offset: number, modifiers: "Alt"[] = []) {
    const { x, centeredX } = await readGeometry(page);
    const box = (await page.locator('[data-pptx-node-id="2"]').boundingBox())!;
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    // The harness renders at zoom 1, so slide pixels are screen pixels.
    const travel = centeredX - x + offset;

    for (const key of modifiers) await page.keyboard.down(key);
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    await page.mouse.move(startX + travel / 2, startY, { steps: 5 });
    await page.mouse.move(startX + travel, startY, { steps: 5 });
    return {
      release: async () => {
        await page.mouse.up();
        for (const key of modifiers) await page.keyboard.up(key);
      },
    };
  }

  test("snaps a dragged shape onto the slide's center with a guide", async ({ page }) => {
    await open(page);
    const drag = await dragNearCenter(page, 3);

    await expect(page.locator('[data-snap-guide="x"]')).toBeVisible();
    await drag.release();
    await expect(page.locator("[data-snap-guide]")).toHaveCount(0);

    await expect
      .poll(async () => {
        const { x, centeredX } = await readGeometry(page);
        return Math.abs(x - centeredX);
      })
      .toBeLessThan(0.01);
  });

  test("places freely with Alt held", async ({ page }) => {
    await open(page);
    const { x: before } = await readGeometry(page);
    const drag = await dragNearCenter(page, 3, ["Alt"]);

    await expect(page.locator("[data-snap-guide]")).toHaveCount(0);
    await drag.release();

    await expect.poll(async () => (await readGeometry(page)).x).not.toBe(before);
    const { x, centeredX } = await readGeometry(page);
    expect(Math.abs(x - centeredX - 3)).toBeLessThan(1);
  });
});

test.describe("rotating", () => {
  test("rotates a shape by dragging its rotate handle", async ({ page }) => {
    const { overlay, shape } = await open(page);
    await shape.click();
    await expect(overlay).toHaveAttribute("data-mode", "selected");

    const box = (await page.locator('[data-pptx-node-id="2"]').boundingBox())!;
    const knob = (await page.locator("[data-rotate-handle]").boundingBox())!;
    const centerX = box.x + box.width / 2;
    const centerY = box.y + box.height / 2;

    // A quarter turn clockwise: from above the center to its right.
    await page.mouse.move(knob.x + knob.width / 2, knob.y + knob.height / 2);
    await page.mouse.down();
    await page.mouse.move(centerX + 60, centerY - 60, { steps: 4 });
    await page.mouse.move(centerX + 150, centerY, { steps: 4 });
    await expect(overlay).toHaveAttribute("data-mode", "rotate");
    await page.mouse.up();

    await expect
      .poll(() => page.evaluate(() => window.__editStore!.getActiveSlide()!.nodes[0].rotation))
      .toBeCloseTo(90, 0);
    await expect(overlay).toHaveAttribute("data-mode", "selected");
  });

  test("resizes a rotated shape along its own axes", async ({ page }) => {
    const { overlay, shape } = await open(page);
    await page.evaluate(async () => {
      const store = window.__editStore!;
      const slideId = store.getActiveSlide()!.id;
      await store.edit({ type: "setNodeTransform", slideId, nodeId: "2", rotation: 90 });
    });
    await shape.click();
    await expect(overlay).toHaveAttribute("data-mode", "selected");

    const readGeometry = () =>
      page.evaluate(() => {
        const node = window.__editStore!.getActiveSlide()!.nodes[0];
        const centerX = node.position.x + node.size.w / 2;
        // Turned 90°, the shape's height runs across the slide.
        return { w: node.size.w, h: node.size.h, left: centerX - node.size.h / 2 };
      });
    const before = await readGeometry();

    // The north handle now sits on the right side: drag it further right.
    const handle = (await page.locator('[data-resize-grip="n"]').boundingBox())!;
    const x = handle.x + handle.width / 2;
    const y = handle.y + handle.height / 2;
    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x + 30, y, { steps: 5 });
    await page.mouse.up();

    await expect.poll(async () => (await readGeometry()).h).toBeCloseTo(before.h + 30, 0);
    const after = await readGeometry();
    expect(after.w).toBeCloseTo(before.w, 1);
    expect(after.left).toBeCloseTo(before.left, 1);
  });
});

test.describe("images", () => {
  test("renders an inserted image straight away", async ({ page }) => {
    await open(page);
    const nodeId = await page.evaluate(async () => {
      const canvas = document.createElement("canvas");
      canvas.width = 120;
      canvas.height = 60;
      const context = canvas.getContext("2d")!;
      context.fillStyle = "#16a34a";
      context.fillRect(0, 0, 120, 60);
      const blob = await new Promise<Blob>((resolve) =>
        canvas.toBlob((b) => resolve(b!), "image/png"),
      );
      const store = window.__editStore!;
      const { createdNodeIds } = await store.edit({
        type: "addImage",
        slideId: store.getActiveSlide()!.id,
        data: new Uint8Array(await blob.arrayBuffer()),
        contentType: "image/png",
      });
      return createdNodeIds![0];
    });

    const image = page.locator(`[data-pptx-node-id="${nodeId}"] img`);
    await expect(image).toBeVisible();
    await expect.poll(() => image.evaluate((img: HTMLImageElement) => img.naturalWidth)).toBe(120);
    const box = (await image.boundingBox())!;
    expect(box.width).toBeCloseTo(120, 0);
  });
});
