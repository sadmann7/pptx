# Review

pptx opens, renders, edits, and saves PowerPoint decks. Failure means a deck that loses data, renders differently from PowerPoint, or saves a file PowerPoint refuses to open.

## Checks

- `pnpm lint`, `pnpm format:check`, `pnpm typecheck`, `pnpm test`.
- `pnpm -F @diceui/pptx-e2e test` for interaction, rendering, and save changes.
- `pnpm ui:check` when `packages/ui` changes.

## What to look for

- **Round trip.** Every edit in `packages/core` survives apply, save, reopen, and undo. Parts, relationships, and attributes the editor doesn't understand are kept as they were, not dropped.
- **Valid OOXML.** Element order, required attributes, relationship ids, and content types follow the schema. PowerPoint is stricter than our parser.
- **Rendering.** Changes to `packages/core/src/renderer` match PowerPoint's output. Oracle baseline updates need a reason and the SSIM scores.
- **Interaction.** Pointer, focus, keyboard, clipboard, and menus are tested in a real browser under `packages/e2e/specs`, not in happy-dom with mocked layout APIs.
- **Tests.** A bug fix comes with a test that fails without it. No sleeps; tests wait on `data-mode`, the store, or the focused element.
- **Published packages.** Changes to behavior or public types need a changeset.
- **Fixes.** A bug fix names the cause and fixes it, not the symptom.

## Skip

- Formatting and lint-enforced style.
- `packages/ui` files synced from Dice UI; the fix belongs there.
- `pnpm-lock.yaml`, `CHANGELOG.md`, fixtures, and recorded baselines unless the change is about them.

## Format

- Start with a count: `N important, N nit, N pre-existing`. Say plainly when nothing needs fixing.
- Every finding gives `file:line` and the input or state that breaks.
- At most three nits.
