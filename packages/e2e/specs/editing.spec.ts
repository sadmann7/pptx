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

test.describe("scrolling", () => {
  async function openCutOff(page: Page) {
    const opened = await open(page);
    const shape = (await page.locator('[data-pptx-node-id="2"]').boundingBox())!;
    const overlay = (await opened.overlay.boundingBox())!;
    const height = Math.ceil(shape.y + shape.height + 20);
    expect(height).toBeLessThan(overlay.y + overlay.height);
    await page.setViewportSize({ width: page.viewportSize()!.width, height });
    return { ...opened, shape };
  }

  test("keeps the page still when dragging a shape", async ({ page }) => {
    const { overlay, shape } = await openCutOff(page);
    const x = shape.x + shape.width / 2;
    const y = shape.y + shape.height / 2;

    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x + 40, y, { steps: 5 });
    await page.mouse.up();

    await expect(overlay).toHaveAttribute("data-mode", "selected");
    await expect(overlay).toBeFocused();
    expect(await page.evaluate(() => window.scrollY)).toBe(0);
  });

  test("keeps the page still when starting a drawing", async ({ page }) => {
    const { overlay, shape } = await openCutOff(page);
    await page.getByTestId("draw-ellipse").click();
    await page.getByTestId("italic").focus();
    await expect(overlay).toHaveCSS("cursor", "crosshair");

    await page.mouse.move(shape.x + 10, shape.y + 10);
    await page.mouse.down();
    await expect(overlay).toBeFocused();
    expect(await page.evaluate(() => window.scrollY)).toBe(0);
    await page.mouse.up();
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
    const handle = (await page.locator('[data-resize-handle="n"]').boundingBox())!;
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

test.describe("text boxes", () => {
  /** The text of the most recently added node, as the overlay currently shows it. */
  function lastNodeText(page: Page) {
    return page.evaluate(() => {
      const nodes = window.__editStore!.getActiveSlide()!.nodes;
      const id = nodes[nodes.length - 1].id;
      const element = document.querySelector(`[data-pptx-node-id="${id}"]`);
      return (element?.textContent ?? "").replace(/\u200B/g, "");
    });
  }

  test("inserts a text box ready to type in, with native undo", async ({ page }) => {
    const { overlay } = await open(page);
    await page.getByTestId("text-box").click();
    await expect(overlay).toHaveAttribute("data-mode", "text");
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement)?.isContentEditable))
      .toBe(true);

    await page.keyboard.type("Hi");
    await expect.poll(() => lastNodeText(page)).toBe("Hi");
    // Each keystroke undoes in turn, the first one included.
    await page.keyboard.press("ControlOrMeta+z");
    await expect.poll(() => lastNodeText(page)).toBe("H");
    await page.keyboard.press("ControlOrMeta+z");
    await expect.poll(() => lastNodeText(page)).toBe("");

    await page.keyboard.type("Hello");
    await page.keyboard.press("Escape");
    await expect(overlay).toHaveAttribute("data-mode", "selected");
    const text = await page.evaluate(() => {
      const nodes = window.__editStore!.getActiveSlide()!.nodes as unknown as {
        textBody: { paragraphs: { runs: { text: string }[] }[] };
      }[];
      return nodes[nodes.length - 1].textBody.paragraphs[0].runs.map((run) => run.text).join("");
    });
    expect(text).toContain("Hello");
  });

  test("undoes the first keystroke typed into an empty selected box", async ({ page }) => {
    const { overlay } = await open(page);
    await page.getByTestId("text-box").click();
    await expect(overlay).toHaveAttribute("data-mode", "text");
    await page.keyboard.press("Escape");
    await expect(overlay).toHaveAttribute("data-mode", "selected");

    await page.keyboard.type("X");
    await expect(overlay).toHaveAttribute("data-mode", "text");
    await expect.poll(() => lastNodeText(page)).toBe("X");
    await page.keyboard.press("ControlOrMeta+z");
    await expect.poll(() => lastNodeText(page)).toBe("");
  });
});

test.describe("drawing", () => {
  const lastNode = (page: Page) =>
    page.evaluate(() => {
      const nodes = window.__editStore!.getActiveSlide()!.nodes;
      const node = nodes[nodes.length - 1] as unknown as {
        presetGeometry?: string;
        position: { x: number; y: number };
        size: { w: number; h: number };
      };
      return { count: nodes.length, ...node };
    });

  test("places a text box where the slide is clicked, ready to type in", async ({ page }) => {
    const { overlay } = await open(page);
    await page.getByTestId("draw-text-box").click();
    await expect(overlay).toHaveCSS("cursor", "crosshair");

    const slide = (await page
      .locator("[data-pptx-selection]")
      .locator("xpath=../..")
      .boundingBox())!;
    const before = await lastNode(page);
    await page.mouse.click(slide.x + 40, slide.y + 400);
    await expect(overlay).toHaveAttribute("data-mode", "text");
    // Resize handles stay while typing, as in Google Slides.
    await expect(page.locator("[data-resize-handle]")).toHaveCount(8);
    const node = await lastNode(page);
    expect(node.count).toBe(before.count + 1);
    expect(node.position.x * (slide.width / 1280)).toBeCloseTo(40, -1);
    await expect(overlay).not.toHaveCSS("cursor", "crosshair");

    // Left empty, it goes away once editing moves elsewhere.
    await page.mouse.click(slide.x + slide.width - 40, slide.y + 400);
    await expect.poll(async () => (await lastNode(page)).count).toBe(before.count);
  });

  test("draws a shape the size of the drag and selects it", async ({ page }) => {
    const { overlay } = await open(page);
    await page.getByTestId("draw-ellipse").click();
    const slide = (await page
      .locator("[data-pptx-selection]")
      .locator("xpath=../..")
      .boundingBox())!;
    const scale = slide.width / 1280;
    await page.mouse.move(slide.x + 600 * scale, slide.y + 300 * scale);
    await page.mouse.down();
    await page.mouse.move(slide.x + 800 * scale, slide.y + 400 * scale, { steps: 5 });
    await expect(page.locator("[data-marquee][data-draw]")).toBeVisible();
    await page.mouse.up();

    await expect(overlay).toHaveAttribute("data-mode", "selected");
    const node = await lastNode(page);
    expect(node.presetGeometry).toBe("ellipse");
    expect(node.size.w).toBeCloseTo(200, -1);
    expect(node.size.h).toBeCloseTo(100, -1);
  });

  test("selects a text box on click, edits it on a second, and selects it from its frame", async ({
    page,
  }) => {
    const { overlay } = await open(page);
    const slide = (await page
      .locator("[data-pptx-selection]")
      .locator("xpath=../..")
      .boundingBox())!;
    await page.getByTestId("draw-text-box").click();
    await page.mouse.click(slide.x + 40, slide.y + 400);
    await expect(overlay).toHaveAttribute("data-mode", "text");
    await page.keyboard.type("Hello");
    await page.mouse.click(slide.x + slide.width - 40, slide.y + 400);
    await expect(overlay).toHaveAttribute("data-mode", "idle");

    const box = (await page.locator("[data-pptx-node-id]").last().boundingBox())!;
    const center = { x: box.x + box.width / 2, y: box.y + box.height / 2 };
    await page.mouse.click(center.x, center.y);
    await expect(overlay).toHaveAttribute("data-mode", "selected");
    await page.mouse.click(center.x, center.y);
    await expect(overlay).toHaveAttribute("data-mode", "text");

    const frame = (await page.locator("[data-selection-box]").first().boundingBox())!;
    await page.mouse.click(frame.x + frame.width / 2, frame.y + 1);
    await expect(overlay).toHaveAttribute("data-mode", "selected");

    const before = await lastNode(page);
    await page.keyboard.press("Delete");
    await expect.poll(async () => (await lastNode(page)).count).toBe(before.count - 1);
  });

  test("disarms with Escape", async ({ page }) => {
    const { overlay } = await open(page);
    await page.getByTestId("draw-ellipse").click();
    await expect(overlay).toHaveCSS("cursor", "crosshair");
    await page.keyboard.press("Escape");
    await expect(overlay).not.toHaveCSS("cursor", "crosshair");
  });
});

test.describe("outlines", () => {
  test("draws a shape's outline", async ({ page }) => {
    await open(page);
    await page.evaluate(async () => {
      const store = window.__editStore!;
      await store.edit({
        type: "setOutline",
        slideId: store.getActiveSlide()!.id,
        nodeId: "2",
        color: "#00FF00",
        width: 4,
      });
    });

    const stroke = page.locator('[data-pptx-node-id="2"] [stroke]').first();
    await expect(stroke).toHaveAttribute("stroke", /#00ff00|rgb\(0, ?255, ?0\)/i);
  });
});

test.describe("clipboard", () => {
  test("copies and pastes a shape with the keyboard", async ({ page }) => {
    const { overlay, shape } = await open(page);
    await shape.click();
    await expect(overlay).toHaveAttribute("data-mode", "selected");

    await page.keyboard.press("ControlOrMeta+c");
    await page.keyboard.press("ControlOrMeta+v");

    const count = () => page.evaluate(() => window.__editStore!.getActiveSlide()!.nodes.length);
    await expect.poll(count).toBe(2);
    const positions = await page.evaluate(() =>
      window.__editStore!.getActiveSlide()!.nodes.map((node) => node.position),
    );
    expect(positions[1].x).toBeCloseTo(positions[0].x + 10);

    await page.keyboard.press("ControlOrMeta+z");
    await expect.poll(count).toBe(1);
  });
});

test.describe("tables", () => {
  interface TableModel {
    columns: number[];
    position: { x: number; y: number };
    size: { w: number; h: number };
    rows: {
      cells: {
        gridSpan: number;
        rowSpan: number;
        textBody?: { paragraphs: { runs: { text: string }[] }[] };
        properties?: { element?: Element };
      }[];
    }[];
  }

  function readTable(page: Page, nodeId: string) {
    return page.evaluate((id) => {
      const node = window
        .__editStore!.getActiveSlide()!
        .nodes.find((n) => n.id === id) as unknown as TableModel;
      return {
        columns: node.columns,
        position: node.position,
        size: node.size,
        cells: node.rows.map((row) =>
          row.cells.map((cell) => ({
            text: (cell.textBody?.paragraphs ?? [])
              .map((p) => p.runs.map((run) => run.text).join(""))
              .join("\n"),
            span: [cell.rowSpan, cell.gridSpan],
            fill:
              cell.properties?.element
                ?.getElementsByTagNameNS("*", "srgbClr")[0]
                ?.getAttribute("val") ?? null,
          })),
        ),
      };
    }, nodeId);
  }

  const editingCell = (page: Page) =>
    page.evaluate(() => document.activeElement?.getAttribute("data-pptx-cell") ?? null);

  const readState = async (page: Page) =>
    JSON.parse((await page.getByTestId("state").textContent()) ?? "{}");

  /** Inserts a 2×3 table, which starts editing its first cell. Returns its id. */
  async function insertTable(page: Page) {
    const { overlay } = await open(page);
    await page.getByTestId("insert-table").click();
    await expect(overlay).toHaveAttribute("data-mode", "text");
    await expect.poll(() => editingCell(page)).toBe("0,0");
    const nodeId = await page.evaluate(() => {
      const nodes = window.__editStore!.getActiveSlide()!.nodes;
      return nodes[nodes.length - 1].id;
    });
    const cellCenter = async (row: number, column: number) => {
      const box = (await page
        .locator(`[data-pptx-node-id="${nodeId}"] [data-pptx-cell="${row},${column}"]`)
        .boundingBox())!;
      return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
    };
    return { overlay, nodeId, cellCenter };
  }

  test("types into cells, moves with Tab and arrow keys, and adds a row past the last cell", async ({
    page,
  }) => {
    const { nodeId } = await insertTable(page);

    await page.keyboard.type("A");
    await page.keyboard.press("Tab");
    await expect.poll(() => editingCell(page)).toBe("0,1");
    await page.keyboard.type("B");
    await expect.poll(async () => (await readTable(page, nodeId)).cells[0][0].text).toBe("A");

    for (const [key, cell] of [
      ["ArrowRight", "0,2"],
      ["ArrowDown", "1,2"],
      ["ArrowLeft", "1,1"],
      ["ArrowUp", "0,1"],
    ] as const) {
      // The caret has to sit at the text's edge before an arrow leaves the cell.
      await page.keyboard.press(key === "ArrowLeft" || key === "ArrowUp" ? LINE_START : LINE_END);
      await page.keyboard.press(key);
      await expect.poll(() => editingCell(page)).toBe(cell);
    }
    await expect.poll(async () => (await readTable(page, nodeId)).cells[0][1].text).toBe("B");

    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("ArrowRight");
    await expect.poll(() => editingCell(page)).toBe("1,2");
    await page.keyboard.press("Tab");
    await expect.poll(async () => (await readTable(page, nodeId)).cells.length).toBe(3);
    await expect.poll(() => editingCell(page)).toBe("2,0");
  });

  test("selects cells by dragging from an unselected table, then fills, merges, and undoes", async ({
    page,
  }) => {
    const { overlay, nodeId, cellCenter } = await insertTable(page);
    await page.getByTestId("outside").click();
    await expect(overlay).toHaveAttribute("data-mode", "idle");

    const from = await cellCenter(0, 0);
    const to = await cellCenter(1, 1);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move(to.x, to.y, { steps: 5 });
    await page.mouse.up();
    await expect(overlay).toHaveAttribute("data-mode", "text");
    await expect
      .poll(async () => (await readState(page)).cellRange)
      .toEqual({ start: { row: 0, column: 0 }, end: { row: 1, column: 1 } });

    await page.getByTestId("cell-fill").click();
    await expect
      .poll(async () => (await readTable(page, nodeId)).cells.map((row) => row.map((c) => c.fill)))
      .toEqual([
        ["FF0000", "FF0000", null],
        ["FF0000", "FF0000", null],
      ]);

    await page.getByTestId("merge-cells").click();
    await expect.poll(async () => (await readTable(page, nodeId)).cells[0][0].span).toEqual([2, 2]);
    // Undo and redo land without leaving the cell first.
    await expect.poll(() => editingCell(page)).toBe("0,0");
    await page.keyboard.press("ControlOrMeta+z");
    await expect.poll(async () => (await readTable(page, nodeId)).cells[0][0].span).toEqual([1, 1]);
    await page.keyboard.press("ControlOrMeta+Shift+z");
    await expect.poll(async () => (await readTable(page, nodeId)).cells[0][0].span).toEqual([2, 2]);
  });

  test("selects the whole table with Cmd/Ctrl+A only once cells are selected", async ({ page }) => {
    const { nodeId, cellCenter } = await insertTable(page);
    await page.keyboard.type("Hi");

    // In a cell, it selects the cell's text, however often it's pressed.
    await page.keyboard.press("ControlOrMeta+a");
    await page.keyboard.press("ControlOrMeta+a");
    await expect
      .poll(() =>
        page.evaluate(() =>
          window
            .getSelection()
            ?.toString()
            .replace(/\u200B/g, ""),
        ),
      )
      .toBe("Hi");
    expect((await readState(page)).cellRange).toBeNull();

    const from = await cellCenter(0, 1);
    const to = await cellCenter(1, 1);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move(to.x, to.y, { steps: 5 });
    await page.mouse.up();
    await page.keyboard.press("ControlOrMeta+a");
    await expect
      .poll(async () => (await readState(page)).cellRange)
      .toEqual({ start: { row: 0, column: 0 }, end: { row: 1, column: 2 } });

    await page.getByTestId("cell-fill").click();
    await expect
      .poll(async () =>
        (await readTable(page, nodeId)).cells.flat().every((cell) => cell.fill === "FF0000"),
      )
      .toBe(true);
    expect((await readTable(page, nodeId)).cells[0][0].text).toBe("Hi");
  });

  test("types over selected cells at the end of the first one", async ({ page }) => {
    const { nodeId, cellCenter } = await insertTable(page);
    await page.keyboard.type("Old");
    await page.keyboard.press("Tab");
    await page.keyboard.type("Keep");

    const from = await cellCenter(0, 0);
    const to = await cellCenter(1, 1);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move(to.x, to.y, { steps: 5 });
    await page.mouse.up();
    await expect.poll(async () => (await readState(page)).cellRange).not.toBeNull();

    await page.keyboard.type("ab");
    await expect.poll(async () => (await readState(page)).cellRange).toBeNull();
    await expect.poll(() => editingCell(page)).toBe("0,0");
    await page.keyboard.press("Escape");
    await expect
      .poll(async () => (await readTable(page, nodeId)).cells[0].map((cell) => cell.text))
      .toEqual(["Oldab", "Keep", ""]);
  });

  test("pastes a spreadsheet column, copies the range, and clears it", async ({
    page,
    browserName,
  }) => {
    test.skip(
      browserName === "firefox",
      "Firefox drops clipboardData on scripted clipboard events",
    );
    const { nodeId } = await insertTable(page);

    await page.evaluate(() => {
      const data = new DataTransfer();
      data.setData("text/plain", "North\nSouth");
      data.setData("text/html", "<table><tr><td>North</td></tr><tr><td>South</td></tr></table>");
      document.activeElement!.dispatchEvent(
        new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true }),
      );
    });
    await expect
      .poll(async () => (await readTable(page, nodeId)).cells.map((row) => row[0].text))
      .toEqual(["North", "South"]);
    await expect
      .poll(async () => (await readState(page)).cellRange)
      .toEqual({ start: { row: 0, column: 0 }, end: { row: 1, column: 0 } });

    const copied = await page.evaluate(() => {
      const data = new DataTransfer();
      document.activeElement!.dispatchEvent(
        new ClipboardEvent("copy", { clipboardData: data, bubbles: true, cancelable: true }),
      );
      return { text: data.getData("text/plain"), html: data.getData("text/html") };
    });
    expect(copied.text).toBe("North\nSouth");
    expect(copied.html).toContain("<td>North</td>");

    // The paste re-renders the slide; the caret is back in the new cell a frame later.
    await expect.poll(() => page.evaluate(() => document.activeElement?.textContent)).toBe("North");
    await page.keyboard.press("Delete");
    await expect
      .poll(async () => (await readTable(page, nodeId)).cells.map((row) => row[0].text))
      .toEqual(["", ""]);
  });

  test("resizes a column of an unselected table by its border, and moves it by its frame", async ({
    page,
  }) => {
    const { overlay, nodeId } = await insertTable(page);
    // Borders resize a table that isn't selected yet.
    await page.getByTestId("outside").click();
    await expect(overlay).toHaveAttribute("data-mode", "idle");
    const before = await readTable(page, nodeId);

    const border = (await page.locator('[data-pptx-table-border="column"]').nth(1).boundingBox())!;
    await page.mouse.move(border.x + border.width / 2, border.y + 10);
    await page.mouse.down();
    await page.mouse.move(border.x + border.width / 2 + 40, border.y + 10, { steps: 5 });
    await page.mouse.up();
    await expect
      .poll(async () => (await readTable(page, nodeId)).columns[0])
      .toBeGreaterThan(before.columns[0] + 20);
    await expect(overlay).toHaveAttribute("data-mode", "selected");
    const resized = await readTable(page, nodeId);
    expect(resized.columns[0] + resized.columns[1]).toBeCloseTo(
      before.columns[0] + before.columns[1],
      0,
    );

    // The left edge resizes the first column and keeps the right edge in place.
    const edge = (await page.locator('[data-pptx-table-border="column"]').first().boundingBox())!;
    const right = resized.position.x + resized.size.w;
    await page.mouse.move(edge.x + edge.width / 2, edge.y + 10);
    await page.mouse.down();
    await page.mouse.move(edge.x + edge.width / 2 + 30, edge.y + 10, { steps: 5 });
    await page.mouse.up();
    await expect
      .poll(async () => (await readTable(page, nodeId)).columns[0])
      .toBeLessThan(resized.columns[0] - 20);
    const trimmed = await readTable(page, nodeId);
    expect(trimmed.position.x + trimmed.size.w).toBeCloseTo(right, 0);

    const table = (await page.locator(`[data-pptx-node-id="${nodeId}"]`).boundingBox())!;
    const grab = { x: table.x + table.width + 5, y: table.y + 10 };
    await page.mouse.move(grab.x, grab.y);
    await page.mouse.down();
    await page.mouse.move(grab.x - 60, grab.y + 40, { steps: 5 });
    await page.mouse.up();
    await expect
      .poll(async () => (await readTable(page, nodeId)).position.y)
      .toBeGreaterThan(trimmed.position.y + 20);
    expect((await readTable(page, nodeId)).size.w).toBeCloseTo(trimmed.size.w, 0);
  });
});
