import { readFile, writeFile } from "node:fs/promises";
import { summarize } from "../plugin/sh.iva/extension/lib/pipeline.ts";

const [source, target] = process.argv.slice(2);
if (!source || !target) throw new Error("usage: node probe-summary.mjs DEEPGRAM_JSON OUTPUT_JSON");
const data = JSON.parse(await readFile(source, "utf8"));
const summary = await summarize(data.results.utterances);
await writeFile(target, JSON.stringify(summary, null, 2), { mode: 0o600 });
console.log({ bullets: summary.bullets.length, tasks: summary.tasks.length, decisions: summary.decisions.length });
