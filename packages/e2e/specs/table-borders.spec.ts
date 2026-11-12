/**
 * A table with no a:tableStyleId gets PowerPoint's "No Style, Table Grid"
 * hairlines, each shared edge is painted once by the cell above or left of it,
 * and an outline every cell agrees on moves to the table. Only a real
 * rasterizer shows whether CSS zoom keeps hairlines on whole device pixels.
 */
import { expect, type Page, test } from "@playwright/test";

import { colorDistance, formatPixel, type Pixel, readPixels } from "./pixels";
import { openSlide, slideContainer } from "./utils";

const DECK = "table-borders.pptx";
const GRID_SLIDE = 0;
const EXPLICIT_SLIDE = 1;

const BORDER = "#000000";
const BLANK_CELL_FILL = "#ffffff";

// Lands the grid lines between device pixels.
const FRACTIONAL_SCALE = 0.86;

interface CellBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

async function cellBoxes(page: Page): Promise<CellBox[]> {
  return page.evaluate(() =>
    [...document.querySelectorAll("table td")].map((td) => {
      const rect = td.getBoundingClientRect();
      return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
    }),
  );
}

async function borderStyles(page: Page, selector: string, index = 0) {
  return page.evaluate(
    ({ sel, at }) => {
      const style = getComputedStyle(document.querySelectorAll(sel)[at]);
      return {
        top: [style.borderTopWidth, style.borderTopColor],
        right: [style.borderRightWidth, style.borderRightColor],
        bottom: [style.borderBottomWidth, style.borderBottomColor],
        left: [style.borderLeftWidth, style.borderLeftColor],
      };
    },
    { sel: selector, at: index },
  );
}

function darkestNear(
  samples: Pixel[],
  from: number,
): { pixel: Pixel; at: number; distance: number } {
  let best = {
    pixel: samples[0],
    at: from,
    distance: colorDistance(samples[0], BORDER),
  };
  samples.forEach((pixel, index) => {
    const distance = colorDistance(pixel, BORDER);
    if (distance < best.distance) best = { pixel, at: from + index, distance };
  });
  return best;
}

test.describe("table with no style id", () => {
  test("paints the built-in grid and hoists the outline onto the table", async ({
    page,
  }) => {
    await openSlide(page, DECK, GRID_SLIDE);

    // tx1 hairlines from the "No Style, Table Grid" fallback: 1pt black.
    const table = await borderStyles(page, "table");
    for (const [width, color] of Object.values(table)) {
      expect(width).toBe("1px");
      expect(color).toBe("rgb(0, 0, 0)");
    }

    const topLeft = await borderStyles(page, "table td", 0);
    expect(topLeft.top[0]).toBe("0px");
    expect(topLeft.left[0]).toBe("0px");
    expect(topLeft.right[0]).toBe("1px");
    expect(topLeft.bottom[0]).toBe("1px");

    const topMiddle = await borderStyles(page, "table td", 1);
    expect(topMiddle.left[0]).toBe("0px");
    const middleLeft = await borderStyles(page, "table td", 3);
    expect(middleLeft.top[0]).toBe("0px");
  });

  test("keeps grid lines on whole pixels when the slide is scaled", async ({
    page,
  }) => {
    await openSlide(page, DECK, GRID_SLIDE, { scale: FRACTIONAL_SCALE });

    const cells = await cellBoxes(page);
    // Bottom row, which has no text: a scan across it crosses grid lines and
    // cell fills only.
    const bottomLeft = cells[6];
    const y = Math.round(bottomLeft.y + bottomLeft.height / 2);
    const edge = Math.round(bottomLeft.x + bottomLeft.width);
    const [from, to] = [edge - 3, edge + 3];

    const pixels = await readPixels(await slideContainer(page).screenshot());
    const line = darkestNear(pixels.row(y, from, to), from);

    expect(
      line.distance,
      `grid line at x=${edge} rendered as ${formatPixel(line.pixel)}; a hairline spread over two device pixels never reaches the border colour`,
    ).toBeLessThanOrEqual(24);

    // Both sides of the line are the untouched cell fill, so the line is one
    // pixel wide rather than a wide smear that happens to contain a dark core.
    for (const offset of [-3, 3]) {
      expect(
        colorDistance(pixels.at(line.at + offset, y), BLANK_CELL_FILL),
      ).toBeLessThanOrEqual(4);
    }
  });

  test("would smear the same grid line under transform scaling", async ({
    page,
  }) => {
    // Proves the crisp result above comes from zoom. If this fails, transform
    // scaling became viable and applySlideScale can be revisited.
    await openSlide(page, DECK, GRID_SLIDE, {
      scale: FRACTIONAL_SCALE,
      mode: "transform",
    });

    const cells = await cellBoxes(page);
    const bottomLeft = cells[6];
    const y = Math.round(bottomLeft.y + bottomLeft.height / 2);
    const edge = Math.round(bottomLeft.x + bottomLeft.width);

    const pixels = await readPixels(await slideContainer(page).screenshot());
    const line = darkestNear(pixels.row(y, edge - 3, edge + 3), edge - 3);

    expect(
      line.distance,
      `transform-scaled hairline rendered as ${formatPixel(line.pixel)}`,
    ).toBeGreaterThan(24);
  });

  test("renders the grid", async ({ page }) => {
    await openSlide(page, DECK, GRID_SLIDE);
    await expect(slideContainer(page)).toHaveScreenshot(
      "table-borders-grid.png",
    );
  });
});

test.describe("table with per-cell borders", () => {
  test("hoists only the outline edges its cells agree on", async ({ page }) => {
    await openSlide(page, DECK, EXPLICIT_SLIDE);

    // Left, right and bottom are a uniform blue, so they move to the table;
    // the top edge is red over the first column and green over the rest, which
    // one border on the table could not express.
    const table = await borderStyles(page, "table");
    expect(table.top[0]).toBe("0px");
    for (const side of [table.left, table.right, table.bottom]) {
      expect(side[0]).toBe("1px");
      expect(side[1]).toBe("rgb(0, 112, 192)");
    }

    const [first, second] = await Promise.all([
      borderStyles(page, "table td", 0),
      borderStyles(page, "table td", 1),
    ]);
    expect(first.top).toEqual(["1px", "rgb(192, 0, 0)"]);
    expect(second.top).toEqual(["1px", "rgb(0, 176, 80)"]);
    expect(first.right[0]).toBe("1px");
    expect(second.left[0]).toBe("0px");
  });

  test("renders the per-cell borders", async ({ page }) => {
    await openSlide(page, DECK, EXPLICIT_SLIDE);
    await expect(slideContainer(page)).toHaveScreenshot(
      "table-borders-explicit.png",
    );
  });
});
