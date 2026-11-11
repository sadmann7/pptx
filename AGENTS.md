# Agents

## Verifying changes

- Every bug fix comes with a test that fails without it.
- Interaction behavior (pointer, focus, keyboard, clipboard, menus) is tested in a real browser in `packages/e2e/specs`, not in happy-dom. Don't mock browser APIs like `elementsFromPoint` or `getClientRects` to make a unit test pass.
- Unit tests in `packages/react` cover state and pure logic. Core edits are tested on real decks: apply, save, reopen, undo.
- No sleeps in tests. Wait on app state: `data-mode`, the store, or the focused element.
- Before opening a PR, check the change in the docs playground (`/pg`) and note how it was verified in the PR description.

## Naming DOM attributes and CSS variables

- Core writes `data-pptx-*` and `--pptx-*`.
- The React primitives write `data-presentation-*` and `--presentation-*`, matching the `Presentation` namespace.
- State on a part's own element has no prefix: `data-mode`, `data-active`, `data-dragging`.
