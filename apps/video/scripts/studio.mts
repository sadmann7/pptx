import { spawn } from "node:child_process";
import net from "node:net";

function getIsFreePort(port: number) {
  return new Promise<boolean>((resolve) => {
    const server = net.createServer();
    server.once("error", () => resolve(false));
    server.listen(port, () => server.close(() => resolve(true)));
  });
}

let port = 3001;
while (!(await getIsFreePort(port))) port++;

spawn("remotion", ["studio", "--no-open", "--port", String(port)], { stdio: "inherit" }).on(
  "exit",
  (code) => process.exit(code ?? 0),
);
