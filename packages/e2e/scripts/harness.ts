import { spawn } from "node:child_process";

import { HARNESS_PORT } from "../lib/constants";

spawn("vite", ["--port", String(HARNESS_PORT), "--strictPort"], { stdio: "inherit" }).on(
  "exit",
  (code) => process.exit(code ?? 1),
);
