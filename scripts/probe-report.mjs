import { readFile } from "node:fs/promises";
import { createWord } from "../plugin/sh.iva/extension/lib/pipeline.ts";

const [transcriptPath, summaryPath, directory] = process.argv.slice(2);
if (!transcriptPath || !summaryPath || !directory) throw new Error("usage: node probe-report.mjs DEEPGRAM_JSON SUMMARY_JSON OUTPUT_DIR");
const transcript = JSON.parse(await readFile(transcriptPath, "utf8"));
const summary = JSON.parse(await readFile(summaryPath, "utf8"));
for (const mode of ["bullets", "transcript", "both"]) {
  const job = { id: `probe-${mode}`, mode, fileName: "meeting.m4a" };
  const result = await createWord(job, summary, transcript.results.utterances, {
    duration: 1576.8,
    dateLabel: "дата не установлена",
  }, directory);
  console.log(result);
}
