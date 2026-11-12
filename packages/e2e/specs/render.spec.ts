import { expect, test } from "@playwright/test";

import { openSlide, slideContainer } from "./utils";

test.describe("basic deck", () => {
  test("renders the first slide", async ({ page }) => {
    await openSlide(page, "basic.pptx", 0);

    await expect(slideContainer(page)).toContainText("Slide one");
    await expect(slideContainer(page)).toHaveScreenshot("basic-slide-1.png");
  });
});

test.describe("BOM-prefixed relationship parts (Open XML SDK output)", () => {
  // Chromium's DOMParser rejects XML behind a UTF-8 BOM, which Firefox tolerates.
  test("renders instead of producing an empty deck", async ({ page }) => {
    await openSlide(page, "bom-rels.pptx", 0);

    await expect(slideContainer(page)).toContainText("BOM deck renders");
    await expect(slideContainer(page)).toHaveScreenshot("bom-rels-slide-1.png");
  });
});

test.describe("tables and groups deck", () => {
  test("renders the table with unequal grid column widths", async ({
    page,
  }) => {
    await openSlide(page, "tables-groups.pptx", 0);

    await expect(slideContainer(page)).toContainText("Matches");
    await expect(slideContainer(page)).toHaveScreenshot(
      "tables-groups-table.png",
    );
  });

  test("renders group children scaled by chExt", async ({ page }) => {
    await openSlide(page, "tables-groups.pptx", 1);

    await expect(slideContainer(page)).toHaveScreenshot(
      "tables-groups-group.png",
    );
  });
});

test.describe("charts nested under ppt/slides/charts/ with literal data", () => {
  // The deck holds chart parts at a non-standard depth, c:strLit/c:numLit literal
  // data, and noFill axis lines that must not fall back to ECharts' gray.
  test("renders the doughnut chart", async ({ page }) => {
    await openSlide(page, "nested-charts.pptx", 0);

    await expect(slideContainer(page).locator("canvas").first()).toBeVisible();
    await expect(slideContainer(page)).toHaveScreenshot(
      "nested-charts-doughnut.png",
    );
  });

  test("renders the bar chart without default axis lines or grid border", async ({
    page,
  }) => {
    await openSlide(page, "nested-charts.pptx", 1);

    await expect(slideContainer(page).locator("canvas").first()).toBeVisible();
    await expect(slideContainer(page)).toHaveScreenshot(
      "nested-charts-bar.png",
    );
  });
});
