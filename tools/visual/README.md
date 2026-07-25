# Visual debugging harness

Renders a single slide outside the viewer and the docs app, screenshots it in a
real browser, and reads the result back pixel by pixel. Built for rendering bugs
that only appear at a particular zoom — hairlines, seams between cell fills,
antialiasing — where "looks wrong" needs to become a number before you can fix
it.

Nothing here is part of the published packages or CI.

## Setup

```bash
pnpm install
pnpm exec playwright install chromium
```

Drop the deck you are debugging in `decks/` (gitignored, so real presentations
stay out of the repo):

```bash
cp ~/Downloads/some-deck.pptx tools/visual/decks/deck.pptx
```

## Render a slide

```bash
pnpm --filter @pptx/visual dev
```

Then open `http://localhost:5399/?file=deck.pptx&slide=7&scale=0.86`. Parameters:

| param   | meaning                                                       |
| ------- | ------------------------------------------------------------- |
| `file`  | deck under `decks/` (default `deck.pptx`)                     |
| `slide` | 1-based slide number                                          |
| `scale` | display scale — fractional values are where artefacts show up |
| `mode`  | `zoom` (what the viewer uses) or `transform` (raster scaling) |

## Screenshot and measure

```bash
cd tools/visual
node shoot.mjs --slide 7 --scale 0.86 --out out/slide7.png
node shoot.mjs --slide 7 --select table --select "tr:nth-child(2) td"
```

`--select` prints each match's client rect plus its zoom/transform, fills and
borders. Use those coordinates as the input to `probe.mjs`. Other flags:
`--file`, `--mode`, `--dpr`, `--width`, `--height`, `--port`, and `--url` to
point at something else entirely (for example the docs playground).

## Read the pixels

```bash
node probe.mjs out/slide7.png --col 740 --y 470:490
node probe.mjs out/slide7.png --row 480 --x 330:1400
node probe.mjs out/slide7.png --crop 700,470,90,40 --zoom 9 --out out/corner.png
```

A crisp 1pt rule reads as a single row of the line colour. A rule spread over
two rows at partial coverage (say `130, 209` where the fills are `248` and
`245`) is the signature of a hairline landing between device pixels: the ink is
split and each half blends with the fill behind it, so the same line looks like
a different colour over each cell.

## Comparing two candidate fixes

Render both, screenshot at the same geometry, and diff the numbers rather than
the images — the difference is usually a few levels on one row of pixels:

```bash
node shoot.mjs --slide 5 --scale 0.86 --mode transform --out out/before.png
node shoot.mjs --slide 5 --scale 0.86 --mode zoom      --out out/after.png
node probe.mjs out/before.png --col 737 --y 575:585
node probe.mjs out/after.png  --col 737 --y 575:585
```

Screenshots land in `out/`, which is gitignored. Delete it whenever.
