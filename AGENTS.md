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

## Comments

Default to no comment. Write one only for a constraint, an external quirk, or an invariant the code can't show, in one line where you can and four at most. Don't restate names, narrate steps, record history, or explain the fix you just made; the PR carries that.

- Full sentences ending with a full stop, including notes after code on the same line.
- No required verb form, and no "we" or "our".
- JSDoc only on the public API (props, options, hooks and functions users import): one sentence, no types, no empty `@param` or `@returns`. Components and parts get none, and neither do internal helpers, even ones another file imports.
- A `TODO:` links the issue that tracks it. No bare `#123`.
- No section dividers, commented-out code, or test step labels.

`pptx/comment-style` in `lint/comment-style.ts` enforces the mechanical parts.
