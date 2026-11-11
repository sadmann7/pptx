import { expect, test } from "@playwright/test";
/**
 * Scores every fixture slide (SSIM) against PowerPoint's PNG export in
 * fixtures/powerpoint/, catching differences from PowerPoint, not just from yesterday.
 * Whole-image SSIM, worst-tile SSIM and region coverage are checked against
 * per-platform baselines in specs/oracle-baselines/, and platforms without any skip.
 *
 *   ORACLE_UPDATE=1 pnpm test   (or: pnpm test:oracle-update)
 */
import { existsSync } from "node:fs";
import { join } from "node:path";

import { ALL_DECKS } from "./decks";
import {
  INK_TOLERANCE,
  isUpdateMode,
  POWERPOINT_DIR,
  readScoreBaseline,
  SCORE_FLOOR,
  SCORE_TOLERANCE,
  scoreAgainstPowerPoint,
  TILE_TOLERANCE,
  writeScoreBaseline,
} from "./oracle";
import { openSlide, slideContainer } from "./utils";

for (const { name: deck, slides } of ALL_DECKS) {
  test.describe(`oracle: ${deck}`, () => {
    for (let slide = 0; slide < slides; slide++) {
      test(`slide ${slide + 1} matches PowerPoint output`, async ({ page }, testInfo) => {
        const project = testInfo.project.name;
        const baseline = readScoreBaseline(deck, slide, project);

        // Baselines are per platform (fonts and antialiasing move SSIM by 0.01-0.03), so
        // without recorded ones this skips before rendering, which is the spec's whole
        // cost; record with "pnpm test:oracle-update". Blank or broken slides are still
        // caught by exported-decks.spec.ts and the screenshot specs.
        test.skip(
          !isUpdateMode && baseline === null,
          `No oracle baseline for ${deck} slide ${slide} (${project}, ${process.platform})`,
        );

        const powerPointPath = join(POWERPOINT_DIR, deck, `slide-${slide}.png`);
        if (!existsSync(powerPointPath)) {
          throw new Error(
            `Missing PowerPoint export ${powerPointPath}. Run "pnpm oracle:export" on a machine with PowerPoint.`,
          );
        }

        await openSlide(page, `${deck}.pptx`, slide);
        const screenshot = await slideContainer(page).screenshot();
        const score = await scoreAgainstPowerPoint(screenshot, powerPointPath);

        if (isUpdateMode) {
          writeScoreBaseline(deck, slide, project, score);
          return;
        }

        // Absolute floor, in case a baseline was recorded from something
        // already broken.
        expect(
          score.overall,
          `SSIM vs PowerPoint fell below the absolute floor ${SCORE_FLOOR}: the slide is likely blank or badly broken`,
        ).toBeGreaterThanOrEqual(SCORE_FLOOR);

        // Narrowing only. A null baseline either skipped above or was just
        // written and returned, so this cannot be reached.
        if (baseline === null) return;

        expect(
          score.overall,
          `SSIM vs PowerPoint dropped below baseline ${baseline.overall} - ${SCORE_TOLERANCE}`,
        ).toBeGreaterThanOrEqual(baseline.overall - SCORE_TOLERANCE);

        expect(
          score.worstTile,
          `The worst tile dropped below baseline ${baseline.worstTile} - ${TILE_TOLERANCE}: ` +
            `some region of the slide regressed even if the slide as a whole still scores well`,
        ).toBeGreaterThanOrEqual(baseline.worstTile - TILE_TOLERANCE);

        // Per region, not worst-of: a region that permanently sits worst would
        // otherwise mask a different region losing its content entirely.
        const worst = score.ink.reduce(
          (found, drift, region) =>
            drift - baseline.ink[region] > found.excess
              ? { region, drift, excess: drift - baseline.ink[region] }
              : found,
          { region: -1, drift: 0, excess: -Infinity },
        );
        expect(
          worst.drift,
          `Region ${worst.region} of the 4x3 grid paints ${worst.drift.toFixed(3)} away from ` +
            `PowerPoint's coverage against a baseline of ${baseline.ink[worst.region]}: content ` +
            `there is missing, collapsed, or newly overflowing`,
        ).toBeLessThanOrEqual(baseline.ink[worst.region] + INK_TOLERANCE);
      });
    }
  });
}
