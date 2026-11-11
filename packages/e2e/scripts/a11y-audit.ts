/**
 * Runs axe (WCAG 2.1 AA) over the docs editor's states in light and dark:
 *
 *   pnpm a11y [origin]   (default http://localhost:3000)
 */
import AxeBuilder from "@axe-core/playwright";
import { chromium, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const origin = process.argv[2] ?? "http://localhost:3000";
const deck = join(dirname(fileURLToPath(import.meta.url)), "..", "fixtures", "basic.pptx");

interface Finding {
  id: string;
  impact: string;
  help: string;
  targets: Set<string>;
  states: Set<string>;
}

const findings = new Map<string, Finding>();

async function audit(page: Page, state: string): Promise<void> {
  // Popups finish opening first: axe reads colors mid-fade as low contrast.
  await page.waitForTimeout(300);
  const results = await new AxeBuilder({ page })
    .withTags(["wcag2a", "wcag2aa", "wcag21a", "wcag21aa"])
    .include("[data-slot=presentation-editor]")
    .include("[data-presentation-keep-selection]")
    // The slide is the deck's own content, with the author's colors.
    .exclude("[data-slot=presentation-slide] [data-pptx-node-id]")
    // Non-modal popups add focus guards: aria-hidden sentinels that pass Tab
    // focus straight on into or out of the popup, never keeping it.
    .exclude("[data-base-ui-focus-guard]")
    // Thumbnails are named "Slide N" and their miniature is aria-hidden, but
    // axe still counts the miniature's text as the option's visible label.
    .disableRules(["label-content-name-mismatch"])
    .analyze();
  for (const violation of results.violations) {
    const finding = findings.get(violation.id) ?? {
      id: violation.id,
      impact: violation.impact ?? "unknown",
      help: violation.help,
      targets: new Set<string>(),
      states: new Set<string>(),
    };
    for (const node of violation.nodes) {
      finding.targets.add(`${node.target.join(" ")}  ${node.html.slice(0, 140)}`);
    }
    finding.states.add(state);
    findings.set(violation.id, finding);
  }
}

async function slideCenter(page: Page) {
  const box = (await page
    .locator("[data-slot=presentation-slide] [data-pptx-node-id]")
    .first()
    .boundingBox())!;
  return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
}

async function walk(page: Page, label: string): Promise<void> {
  await page.waitForSelector("[data-presentation-selection]");
  await audit(page, `${label}: loaded`);

  const { x, y } = await slideCenter(page);
  await page.mouse.click(x, y);
  await audit(page, `${label}: shape selected`);

  await page.getByRole("button", { name: "Fill" }).click();
  await page.waitForSelector("[role=menu]");
  await audit(page, `${label}: fill menu`);
  await page.keyboard.press("Escape");

  await page.mouse.dblclick(x, y);
  await audit(page, `${label}: editing text`);
  await page.getByRole("combobox", { name: "Font" }).click();
  await page.waitForSelector("[role=listbox]");
  await audit(page, `${label}: font list`);
  await page.keyboard.press("Escape");
}

const browser = await chromium.launch();
for (const theme of ["light", "dark"] as const) {
  const context = await browser.newContext({ colorScheme: theme });
  const page = await context.newPage();

  await page.goto(`${origin}/pg`);
  await audit(page, `${theme} /pg: empty`);
  await page.locator('input[type=file][accept=".pptx"]').setInputFiles(deck);
  await walk(page, `${theme} /pg`);

  await page.goto(`${origin}/docs/editing`);
  await walk(page, `${theme} /docs/editing`);
  await context.close();
}
await browser.close();

const order = ["critical", "serious", "moderate", "minor", "unknown"];
const sorted = [...findings.values()].sort(
  (a, b) => order.indexOf(a.impact) - order.indexOf(b.impact),
);
if (sorted.length === 0) console.log("No violations.");
for (const finding of sorted) {
  console.log(`\n[${finding.impact}] ${finding.id}: ${finding.help}`);
  console.log(`  states: ${[...finding.states].join("; ")}`);
  for (const target of [...finding.targets].slice(0, 60)) console.log(`  - ${target}`);
  if (finding.targets.size > 60) console.log(`  … ${finding.targets.size - 60} more`);
}
process.exitCode = sorted.length > 0 ? 1 : 0;
