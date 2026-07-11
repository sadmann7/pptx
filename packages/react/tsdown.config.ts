import { defineConfig } from "tsdown";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  dts: true,
  clean: true,
  deps: { neverBundle: ["react", "react-dom", "react/jsx-runtime", "@diceui/pptx-core"] },
  publint: false,
  banner: { js: '"use client";' },
});
