# Contributing Guide

Thank you for investing your time in contributing to pptx.

If you have any questions, feel free to reach out to [@sadmann17](https://x.com/sadmann17).

## Project Structure

- `packages/core`: `@diceui/pptx-core`, the engine that parses, renders, edits and saves decks. No React.
- `packages/react`: `@diceui/pptx`, the composable `Presentation` primitives built on core.
- `packages/e2e`: browser end-to-end tests with Playwright.
- `packages/ui`: the styled Presentation components, synced from the diceui registry with `pnpm ui:sync`.
- `apps/docs`: the documentation site at [pptx.diceui.com](https://pptx.diceui.com), with the playground at `/pg`.
- `apps/video`: the launch video, made with Remotion.

## Getting started

### Clone the project

```shell
git clone https://github.com/sadmann7/pptx
cd pptx
```

### Create a new branch

```shell
git checkout -b sadman/add-kickflip
```

### Install dependencies

```shell
pnpm install
```

### Start the dev server

```shell
pnpm dev
```

### Check your changes

```shell
pnpm lint
pnpm format:check
pnpm typecheck
pnpm test
```

Run `pnpm -F @diceui/pptx-e2e test` for changes to interaction, rendering or saving, and `pnpm ui:check` when `packages/ui` changes. A bug fix comes with a test that fails without it.

### Add a changeset

A change to `@diceui/pptx` or `@diceui/pptx-core` that users will notice needs a changeset:

```shell
pnpm changeset
```

### Commit your changes

When committing your changes, use the following format:

```text
<type>(<scope>): <description>
```

**Types:**

- `feat` : a new feature
- `fix` : a bug fix
- `refactor` : code change that is neither a feature nor a bug fix
- `perf` : performance improvement
- `style` : formatting, missing semicolons, etc. (no logic change)
- `test` : adding or updating tests
- `docs` : documentation changes only
- `chore` : build process, dependency updates, tooling
- `ci` : continuous integration changes

**Scope:**
Scope your commit to the part of the codebase that is affected, for example `core`, `react`, `selection`, `ui`, `docs`, `ci`, or `deps`.

Examples:

```shell
git commit -m "fix(core): keep a shape's fill when its layout changes"
git commit -m "feat(selection): nudge shapes with the arrow keys"
```

### Create a pull request

When you're finished pushing your changes, open a pull request. Its title uses the same format, since pull requests are squash-merged. A change that breaks the public API gets a `!` after the scope, as in `feat(react)!: rename the slide prop`.
