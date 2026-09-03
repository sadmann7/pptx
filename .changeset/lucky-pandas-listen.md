---
"@diceui/pptx-core": patch
---

Fix a build failure in consuming apps (`Module not found: Can't resolve './worker.ts'`) when importing the package from npm. The embedded-font worker was spawned from a relative source path that only resolved inside the repo, so the published build pointed at a file it never shipped. The worker is now bundled into the build and spawned from a blob URL, the same way the PDF renderer already worked.
