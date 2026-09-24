import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import net from "node:net";
import path from "node:path";

function getIsFreePort(port: number) {
  return new Promise<boolean>((resolve) => {
    const server = net.createServer();
    server.once("error", () => resolve(false));
    server.listen(port, () => server.close(() => resolve(true)));
  });
}

let port = 3001;
while (!(await getIsFreePort(port))) port++;

// `remotion-cli.js` isn't in the package's exports, so resolve it from the
// package root. Running it through node skips the `.cmd` shim on Windows.
const require = createRequire(import.meta.url);
const cli = path.join(
  path.dirname(require.resolve("@remotion/cli/package.json")),
  "remotion-cli.js",
);

spawn(process.execPath, [cli, "studio", "--no-open", "--port", String(port)], {
  stdio: "inherit",
}).on("exit", (code, signal) => {
  if (signal) process.kill(process.pid, signal);
  else process.exit(code ?? 1);
});
