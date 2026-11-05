# Editor performance notes

Goal: 120 fps scrolling the thumbnail list and editing on long decks, measured the way tablecn's data grid is (`data-grid-scroll-perf-notes.md` there).

Frame budget at 120 Hz is ~8.3 ms.

## Setup

```sh
pnpm -F @diceui/pptx-e2e bench-deck            # recordings/perf/decks/bench-300.pptx
PROFILE=1 pnpm -F docs build && pnpm -F docs start -p 3003
pnpm -F @diceui/pptx-e2e profile:editor <label> [scenario]
pnpm -F @diceui/pptx-e2e perf:trace recordings/perf/<label>/<scenario>.json
```

- `bench-deck` grows `fixtures/adventure-club-pin-and-paper.pptx` (pictures, charts, tables, embedded fonts) to 300 slides.
- `profile:editor` loads it into `/pg` on a production build, headed, and runs five scenarios: `thumbs-slow-wheel`, `thumbs-fast-fling`, `thumbs-scrollbar-drag`, `slide-step` (holding ArrowDown on the thumbnails), and `shape-drag`. Frame stats come from an untraced run, then a Chrome trace per scenario lands in `recordings/perf/<label>/`. `blankFramesPct` is the share of frames with an unrendered thumbnail on screen.
- `PROFILE=1` keeps source maps, so `perf:trace` names functions (`packages/core/src/renderer/shape.ts:2524`, not `l`).
- `EXTRA_CSS`, `STRIP_CSS`, and `CPU_THROTTLE` A/B CSS and slower machines without a rebuild.
- The 70–140 ms `EventDispatch` task at the start of every trace is the CPU profiler starting (`V8.StackGuard`), not the editor.

Only compare runs made back to back.

## Baseline (2026-11-05, production build, 120 Hz, 300 slides)

| scenario              | fps         | p95 ms  | max ms | dropped | blank % | main thread ms |
| --------------------- | ----------- | ------- | ------ | ------- | ------- | -------------- |
| thumbs-slow-wheel     | 105.9–107.5 | 16.7    | 18–33  | 80–92   | 34      | 2,651–2,766    |
| thumbs-fast-fling     | 93.8–94.8   | 16.7    | 25–33  | 94–98   | 52      | 2,015–2,022    |
| thumbs-scrollbar-drag | 119.1       | 8.6–9.9 | 17     | 3       | 74–75   | 2,377–2,489    |
| slide-step            | 117.1–117.2 | 9.3–10  | 18–25  | 5       | —       | 1,260–1,266    |
| shape-drag            | 112.9–116.2 | 8.7–9.2 | 65–150 | 17–33   | —       | 1,340–1,399    |

Wheel scrolling spent ~1 s of the run in `Layerize` (~1.4 ms every frame), and shape drag had 42–52 ms tasks in `collectTypefaces`.

## Results

| change                                                                   | effect, back to back                                                                                                                                 | verdict                                                                                          |
| ------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| Toolbar lists the deck's fonts only while the font list is open          | shape drag: 3–4 long frames (42–52 ms in `getPresentationFonts`, rescanning 300 slides on every commit) → 0; max 65–150 → 17 ms; 113–116 → 119.8 fps | kept                                                                                             |
| Render queue budget 8 → 4 ms                                             | slow wheel 107 → 104 fps, 81 → 104 dropped                                                                                                           | rejected: the queue is not what drops frames                                                     |
| No pulse on pending thumbnails (`EXTRA_CSS`, 2 rounds)                   | slow wheel 107–108 → 120 fps, 78–86 → 0–1 dropped, main thread −55%; fast fling 86–90 → 109–110 fps, 114–128 → 30–33 dropped                         | kept: dropped `animate-pulse` from the styled preview                                            |
| `content-visibility: auto` on previews (before this setup, 2,000 slides) | fling p50 17 → 25 ms                                                                                                                                 | rejected: per-frame relevance checks on every item, and off-screen previews are already detached |
| Cap the preview cache at 50 (2,000 slides)                               | heap 157 → 148 MB after reading the whole deck, flat with deck size                                                                                  | kept                                                                                             |

The pulse was the big one. Every pending preview ran an infinite opacity animation, so 48 compositing layers sat in `ActiveOpacityAnimation` and the page re-layerized on every frame, scrolling or not.

## After (same build setup, all kept changes)

| scenario              | fps   | p95 ms | max ms | dropped | blank % | main thread ms |
| --------------------- | ----- | ------ | ------ | ------- | ------- | -------------- |
| thumbs-slow-wheel     | 120   | 9.2    | 9.4    | 0       | 37      | 1,284          |
| thumbs-fast-fling     | 114.8 | 9.3    | 17.5   | 16      | 54      | 1,287          |
| thumbs-scrollbar-drag | 120   | 9      | 9.4    | 0       | 73      | 1,614          |
| slide-step            | 118.3 | 8.8    | 25     | 3       | —       | 830            |
| shape-drag            | 119.5 | 8.9    | 24.9   | 2       | —       | 692            |

## Remaining

1. **Blank thumbnails while scrolling (37–73% of frames).** Rendering a preview costs more than a frame at fling speed. The top self time is `applyDynamicAutofit` (`packages/core/src/renderer/shape.ts`), which forces a layout per autofit text shape to measure it: ~250 ms of a fling's ~1.4 s scripting.
   - Thumbnails stay on `renderSlide`, so they match the slide exactly; `renderThumbnail` and skipping autofit are ruled out.
   - Next lever: measure all of a slide's text shapes in one pass instead of one forced layout each. Same output, unmeasured win.
2. **Fast fling still drops 16 frames.** Frames that render a preview overshoot the 8 ms queue budget (12–13 ms), because one slide's render can't be split.
