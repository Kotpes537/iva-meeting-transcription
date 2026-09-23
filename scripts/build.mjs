import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const command = process.platform === "win32" ? "eve.cmd" : "eve";
const result = spawnSync(command, ["extension", "build"], {
  cwd: resolve(root, "plugin", "sh.iva"),
  stdio: "inherit",
  shell: process.platform === "win32",
});
process.exit(result.status ?? 1);
