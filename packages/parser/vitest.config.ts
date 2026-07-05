import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // SafeXmlNode relies on DOMParser; happy-dom provides it in Node.
    environment: "happy-dom",
    include: ["src/**/*.test.ts"],
  },
});
