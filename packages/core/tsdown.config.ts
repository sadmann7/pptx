import { defineConfig } from "tsdown";

export default defineConfig({
  entry: ["src/index.ts", "src/fonts/index.ts"],
  format: ["esm"],
  dts: true,
  clean: true,
  deps: { neverBundle: ["jszip", "echarts"] },
  publint: false,
});
