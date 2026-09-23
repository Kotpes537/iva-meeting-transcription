import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const commands = [
  ["node_modules/typescript/bin/tsc", "--noEmit", "-p", "plugin/sh.iva/tsconfig.json"],
  ["--test", "test/pipeline.test.mjs"],
  ["scripts/build.mjs"],
];
for (const args of commands) {
  const result = spawnSync(process.execPath, args, { cwd: root, stdio: "inherit" });
  if (result.status !== 0) process.exit(result.status ?? 1);
}
