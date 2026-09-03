import path from "node:path";
import { rolldown } from "rolldown";
import { defineConfig } from "tsdown";

const workerHost = path.join(import.meta.dirname, "src", "fonts", "worker-host.ts");
const workerEntry = path.join(import.meta.dirname, "src", "fonts", "worker.ts");

async function bundleWorker(): Promise<string> {
  const bundle = await rolldown({
    input: workerEntry,
    platform: "browser",
    logLevel: "silent",
  });
  try {
    const { output } = await bundle.generate({ format: "iife", minify: true });
    const chunk = output[0];
    if (!chunk) throw new Error("The font worker bundle produced no output.");
    return chunk.code;
  } finally {
    await bundle.close();
  }
}

export default defineConfig({
  entry: ["src/index.ts", "src/fonts/index.ts"],
  format: ["esm"],
  dts: true,
  clean: true,
  deps: { neverBundle: ["jszip", "echarts"] },
  publint: false,
  plugins: [
    {
      name: "inline-font-worker",
      async load(id) {
        if (path.normalize(id) !== workerHost) return null;
        return `const source = ${JSON.stringify(await bundleWorker())};

export function createFontWorker() {
  const url = URL.createObjectURL(new Blob([source], { type: "text/javascript" }));
  try {
    return new Worker(url);
  } finally {
    URL.revokeObjectURL(url);
  }
}
`;
      },
    },
  ],
});
