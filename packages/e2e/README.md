# @diceui/pptx-e2e

Browser tests for the renderer, and the tooling for debugging what they catch.
Everything runs against a Vite harness that renders one slide of one deck in
isolation, with no viewer, React, or docs app in the way.

## Running the tests

```bash
pnpm test                # both browsers
pnpm test:headed
pnpm test:update         # re-record screenshot baselines
pnpm report
```

The Playwright config starts the harness itself, so no dev server is needed.

## Layers

| Spec             | Asserts on                                                           |
| ---------------- | -------------------------------------------------------------------- |
| `structural`     | the parsed model: nodes, transforms, text, with no pixels involved   |
| `render`         | screenshot baselines per browser                                     |
| `table-borders`  | border ownership and hairline rasterization at fractional scale      |
| `navigation`     | switching slides in a loaded deck                                    |
| `exported-decks` | every slide of an exported deck renders with content and no failures |
| `oracle`         | SSIM and painted coverage against PNGs exported from real PowerPoint |

Structural specs are the cheapest place to pin a parser fix; reach for pixels
only when the bug is in how something rasterizes.

## Fixtures

Decks are listed in `specs/decks.ts`, which the oracle and exported-deck specs
both iterate; add a deck there once its ground truth exists.

The generated decks come from `scripts/generate-fixtures.ts` (`pnpm fixtures`)
and are committed so runs are deterministic. Each is minimal and targets
specific regressions, so read the header comment in the generator before adding
to one.

The exported decks come from an authoring tool, committed as-is. They carry
the full theme/layout/master chain, images, gradients, charts and mixed text
that the minimal fixtures deliberately leave out, so they are what catches
fidelity drift in the parts of the renderer no hand-written fixture reaches.
Hand-written assertions against them would be guesswork, so they are checked
against PowerPoint's own export instead.

Ground-truth PNGs under `fixtures/ground-truth/` come from PowerPoint itself via
`pnpm oracle:export` (Windows, PowerPoint installed), exported at each deck's
native slide size.

Each slide's scores against that ground truth are recorded per platform in
`specs/oracle-baselines/`, since Linux scores lower for want of the decks' fonts
(~0.015, and 0.06 on the worst slide). Re-record locally with
`pnpm test:oracle-update`, and for Linux with the `Record PowerPoint oracle
baselines` workflow, which uploads them as an artifact to review before
committing.

Three numbers are recorded, because SSIM alone scores a _blank_ region higher
than a _misplaced_ one and so rates some broken renders above correct ones.
`specs/oracle.ts` explains what each covers. `pnpm faults` is how that claim
stays honest: it breaks a rendered slide in known ways (hides a text block,
shrinks one, hides a graphic, shifts everything, recolours a fill) and prints
which of the three notice, so a tolerance can be sized against real damage
instead of against the bug that happened to motivate it. Re-run it when
changing those tolerances; it needs `pnpm harness` running.

## Debugging a rendering bug

Start the harness and open a slide directly:

```bash
pnpm harness
# http://localhost:5000/?file=table-borders.pptx&slide=0&scale=0.86
```

| Param   | Meaning                                                                 |
| ------- | ----------------------------------------------------------------------- |
| `file`  | fixture name, or `decks/<name>.pptx` for a local deck (gitignored)      |
| `slide` | 0-based index                                                           |
| `scale` | display scale, for reproducing anything that only breaks when scaled    |
| `mode`  | `zoom` (what the viewer ships) or `transform`, to compare rasterization |

Then screenshot and inspect it at the pixel level:

```bash
pnpm shoot --file decks/customer.pptx --slide 6 --scale 0.86 --select "table td"
pnpm probe out/decks-customer.pptx-6.png --row 292 --x 350:365
pnpm probe out/decks-customer.pptx-6.png --crop 340,270,40,30 --zoom 9
```

`--select` prints client rects plus the styles that decide where an edge lands
on the device pixel grid. `probe` prints channel values along a row or column,
which is the only way to tell a washed-out hairline (partial coverage spread
over two pixels) from a genuinely wrong colour.

Drop decks that cannot be committed into `decks/`; screenshots land in `out/`.
Both are gitignored.
