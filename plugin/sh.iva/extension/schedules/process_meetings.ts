import { readdir, readFile, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { defineSchedule } from "eve/schedules";
import { dataRoot, loadJob, saveJob, tryLock, type MeetingJob } from "../lib/jobs.ts";
import {
  audioMetadata, createWord, fetchLargeAudio, sendBullets, sendWord,
  summarize, telegramBullets, transcribe, type Summary, type Utterance,
} from "../lib/pipeline.ts";

async function readOrCompute<T>(path: string, compute: () => Promise<T>): Promise<T> {
  try { return JSON.parse(await readFile(path, "utf8")) as T; }
  catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
  }
  const result = await compute();
  await writeFile(path, JSON.stringify(result), { mode: 0o600 });
  return result;
}

export async function processOne(job: MeetingJob, directory: string): Promise<void> {
  if (job.state === "done" || job.state === "failed") return;
  job.state = "working";
  job.attempts += 1;
  await saveJob(job);
  try {
    const audio = job.sourcePath ?? await fetchLargeAudio(job, directory);
    job.sourcePath = audio;
    await saveJob(job);
    const metadata = await readOrCompute(join(directory, "metadata.json"), () => audioMetadata(audio, job.fileName));
    const utterances = await readOrCompute<Utterance[]>(join(directory, "transcript.json"), () => transcribe(audio));
    const summary = await readOrCompute<Summary>(join(directory, "summary.json"), () => summarize(utterances));
    if (!job.bulletsMessageId) {
      job.bulletsMessageId = await sendBullets(job, telegramBullets(summary, metadata));
      job.state = "bullets_sent";
      await saveJob(job);
    }
    if (!job.documentMessageId) {
      const report = await createWord(job, summary, utterances, metadata, directory);
      job.documentMessageId = await sendWord(job, report);
      await saveJob(job);
    }
    job.state = "done";
    await saveJob(job);
  } catch (error) {
    job.error = error instanceof Error ? error.message.slice(0, 180) : "processing failed";
    job.state = job.attempts >= 3 ? "failed" : "queued";
    await saveJob(job);
    if (job.state === "failed") {
      await sendBullets(job, "Не удалось завершить транскрибацию встречи. Запись сохранена; причина указана в журнале задачи. Повторная отправка не требуется до проверки ошибки.").catch(() => undefined);
    }
    throw error;
  }
}

export async function processPending(): Promise<void> {
  const jobsRoot = join(dataRoot(), "jobs");
  const directories = await readdir(jobsRoot, { withFileTypes: true }).catch((error: NodeJS.ErrnoException) => {
    if (error.code === "ENOENT") return [];
    throw error;
  });
  for (const entry of directories.filter((x) => x.isDirectory()).sort((a, b) => a.name.localeCompare(b.name))) {
    const directory = join(jobsRoot, entry.name);
    const job = await loadJob(join(directory, "job.json")).catch(() => null);
    if (!job || !["queued", "working", "bullets_sent"].includes(job.state)) continue;
    const unlock = await tryLock(directory);
    if (!unlock) continue;
    try {
      await processOne(job, directory);
    } catch (error) {
      console.error("[meeting-transcription] job failed:", job.id, error instanceof Error ? error.message : "unknown");
    } finally {
      await unlock();
    }
    break; // one long recording per run; next minute handles the next job
  }
}

export default defineSchedule({
  cron: "* * * * *",
  run({ waitUntil }) { waitUntil(processPending()); },
});
