import { readdir, readFile, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { defineSchedule } from "eve/schedules";
import { dataRoot, loadJob, saveJob, tryLock, type MeetingJob } from "../lib/jobs.ts";
import {
  audioMetadata, createWord, fetchLargeAudio, optimizeAudio, sendBullets, sendWord,
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
  let transcriptForFallback: Utterance[] | null = null;
  let metadataForFallback: { duration: number; dateLabel: string } | null = null;
  let summaryReady = false;
  try {
    const sourceAudio = job.sourcePath ?? await fetchLargeAudio(job, directory);
    job.sourcePath = sourceAudio;
    await saveJob(job);
    const metadata = await readOrCompute(join(directory, "metadata.json"), () => audioMetadata(sourceAudio, job.fileName));
    metadataForFallback = metadata;
    const audio = await optimizeAudio(sourceAudio, directory);
    const utterances = await readOrCompute<Utterance[]>(join(directory, "transcript.json"), () => transcribe(audio));
    transcriptForFallback = utterances;
    const summary = await readOrCompute<Summary>(join(directory, "summary.json"), () => summarize(utterances));
    summaryReady = true;
    if (!job.documentMessageId || job.documentIsFallback) {
      const report = await createWord(job, summary, utterances, metadata, directory);
      job.documentMessageId = await sendWord(job, report);
      job.documentIsFallback = false;
      await saveJob(job);
    }
    if (!job.bulletsMessageId) {
      job.bulletsMessageId = await sendBullets(job, telegramBullets(summary, metadata));
      job.state = "bullets_sent";
      await saveJob(job);
    }
    job.state = "done";
    await saveJob(job);
  } catch (error) {
    job.error = error instanceof Error ? error.message.slice(0, 180) : "processing failed";
    job.state = job.attempts >= 3 ? "failed" : "queued";
    await saveJob(job);
    if (!summaryReady && transcriptForFallback && metadataForFallback && !job.documentMessageId) {
      try {
        const fallbackJob = { ...job, mode: "transcript" as const };
        const emptySummary: Summary = { overview: "", participants: [], agenda: [], bullets: [], topics: [], decisions: [], tasks: [], open_questions: [] };
        const report = await createWord(fallbackJob, emptySummary, transcriptForFallback, metadataForFallback, directory);
        job.documentMessageId = await sendWord(job, report);
        job.documentIsFallback = true;
        await saveJob(job);
      } catch (fallbackError) {
        console.error("[meeting-transcription] transcript Word fallback failed:", job.id, fallbackError instanceof Error ? fallbackError.message : "unknown");
      }
    }
    if (job.state === "failed") {
      const notice = job.documentIsFallback
        ? "Полный транскрипт отправлен в Word. Не удалось подготовить итоги встречи; причина указана в журнале задачи."
        : "Не удалось завершить транскрибацию встречи. Запись сохранена; причина указана в журнале задачи. Повторная отправка не требуется до проверки ошибки.";
      await sendBullets(job, notice).catch(() => undefined);
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
