import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";

import { cleanName, privateTarget } from "../plugin/sh.iva/extension/lib/jobs.ts";
import { createWord, dateFromFileName, summarize, timecode, transcriptLines } from "../plugin/sh.iva/extension/lib/pipeline.ts";
import { processOne } from "../plugin/sh.iva/extension/schedules/process_meetings.ts";

test("the recipient comes only from an authenticated private owner turn", () => {
  const auth = { principalType: "user", attributes: { chat_id: "123", user_id: "123", chat_type: "private", message_id: "456" } };
  assert.deepEqual(privateTarget(auth), { chatId: "123", messageId: 456 });
  assert.throws(() => privateTarget({ ...auth, attributes: { ...auth.attributes, chat_type: "group" } }));
  assert.throws(() => privateTarget({ ...auth, attributes: { ...auth.attributes, user_id: "789" } }));
});

test("audio filenames are constrained and timing preserves both ends", () => {
  assert.equal(cleanName("Встреча.m4a"), "Встреча.m4a");
  assert.equal(cleanName("meeting.MP3"), "meeting.MP3");
  for (const ext of ["wav", "flac", "ogg", "oga", "opus", "aac", "webm", "mp4"]) assert.equal(cleanName(`meeting.${ext}`), `meeting.${ext}`);
  assert.throws(() => cleanName("../secret.mp3"));
  assert.throws(() => cleanName("meeting.exe"));
  assert.equal(timecode(1576.8), "00:26:16");
  assert.deepEqual(transcriptLines([{ start: 1.2, end: 3.9, speaker: 1, transcript: "Привет." }]), ["[00:00:01–00:00:03] Спикер 2: Привет."]);
});

test("all three Word modes produce nonempty DOCX files", async () => {
  const directory = await mkdtemp(join(tmpdir(), "iva-meeting-test-"));
  const summary = {
    bullets: ["Обсудили вопрос [00:00:01].", "Назвали следующий шаг [00:00:02]."],
    participants: [{ name: "Анна", role: "руководитель", evidence: "[00:00:01]" }],
    topics: ["Тема"], decisions: [],
    tasks: [{ task: "Проверить", owner: "не назван", due: "не назван", evidence: "[00:00:01]" }],
    open_questions: [],
  };
  const utterances = [{ start: 1, end: 3, speaker: 0, transcript: "Проверим." }];
  try {
    for (const mode of ["bullets", "transcript", "both"]) {
      const job = { id: `test-${mode}`, mode, fileName: "meeting.mp3" };
      const path = await createWord(job, summary, utterances, { duration: 3, dateLabel: "дата не установлена" }, directory);
      assert.ok((await stat(path)).size > 5_000);
      assert.equal((await readFile(path)).subarray(0, 2).toString(), "PK");
    }
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("the selected Iva model can summarize without a Gemini key", async () => {
  let called = false;
  const utterances = [
    { start: 1, end: 2, transcript: "Первый вопрос." },
    { start: 3, end: 4, transcript: "Проверим второй вопрос." },
  ];
  const summary = await summarize(utterances, async (system, prompt) => {
    called = true;
    assert.match(system, /не придумывай имена/iu);
    assert.match(prompt, /\[00:00:01\] Спикер \?: Первый вопрос/u);
    return JSON.stringify({
      overview: "Обсудили два вопроса.", agenda: ["Первый вопрос", "Второй вопрос"],
      participants: [{ name: "Анна", role: "руководитель", evidence: "[00:00:01]" }],
      bullets: [
        { text: "Обсудили первый вопрос.", evidence: "[00:00:01]" },
        "Проверят второй вопрос [00:00:03].",
      ],
      topics: ["Первый вопрос", "Второй вопрос"], decisions: [], tasks: [], open_questions: [],
    });
  });
  assert.equal(called, true);
  assert.equal(summary.bullets.length, 2);
  assert.equal(summary.bullets[0], "Обсудили первый вопрос. [00:00:01]");
  assert.equal(summary.participants[0].name, "Анна");
});

test("summary rejects mentioned third parties and missing issue counts", async () => {
  const utterances = [{ start: 1, end: 2, transcript: "Посоветоваться с Димой. Решили проверить отчёт." }];
  const base = {
    overview: "Проверка отчёта.", agenda: ["Отчёт"],
    participants: [{ name: "Дима", role: "Участник, упомянутый в разговоре", evidence: "[00:00:01]" }],
    bullets: ["Обсудили отчёт [00:00:01]", "Решили проверить [00:00:01]"],
    decisions: [{ decision: "Проверить отчёт", evidence: "[00:00:01]" }], tasks: [], open_questions: [],
  };
  let calls = 0;
  const recovered = await summarize(utterances, async () => {
    calls++;
    return JSON.stringify({ ...base, topics: calls === 1 ? [] : ["Проверка отчёта"] });
  });
  assert.equal(calls, 2);
  assert.deepEqual(recovered.topics, ["Проверка отчёта"]);
  const summary = await summarize(utterances, async () => JSON.stringify({ ...base, topics: [{ issue: "Проверка отчёта", evidence: "[00:00:01]" }] }));
  assert.deepEqual(summary.participants, []);
  assert.deepEqual(summary.topics, ["Проверка отчёта"]);
});

test("recording date is accepted only when the filename encodes a valid time", () => {
  assert.match(dateFromFileName("Голос 260723_160511.wav"), /2026-07-23 16:05:11/u);
  assert.match(dateFromFileName("Голос 260231_160511.wav"), /не установлены/u);
});

test("summary failure still sends the full transcript as a Word file", async () => {
  const root = await mkdtemp(join(tmpdir(), "iva-meeting-fallback-"));
  const id = "test-summary-failure";
  const directory = join(root, "meeting-transcription", "jobs", id);
  await mkdir(directory, { recursive: true });
  const sourcePath = join(directory, "source.mp3");
  await writeFile(sourcePath, "test audio placeholder");
  await writeFile(join(directory, "metadata.json"), JSON.stringify({ duration: 3, dateLabel: "дата не установлена" }));
  await writeFile(join(directory, "transcript.json"), JSON.stringify([{ start: 1, end: 3, speaker: 0, transcript: "Полная реплика встречи." }]));
  await writeFile(join(directory, "summary.json"), "{");
  const job = { id, state: "queued", mode: "both", chatId: "123", messageId: 456, fileName: "meeting.mp3", fileId: null, sourcePath, createdAt: new Date().toISOString(), attempts: 0 };
  const originalFetch = globalThis.fetch;
  const originalDataDir = process.env.ASSISTANT_DATA_DIR;
  const originalToken = process.env.TELEGRAM_BOT_TOKEN;
  const originalAllowed = process.env.TELEGRAM_ALLOWED_USER_IDS;
  const replies = [];
  globalThis.fetch = async (_url, options) => {
    replies.push(options.body.get("reply_to_message_id"));
    return replies.length === 1
      ? Response.json({ ok: false }, { status: 400 })
      : Response.json({ ok: true, result: { message_id: 99 } });
  };
  process.env.ASSISTANT_DATA_DIR = root;
  process.env.TELEGRAM_BOT_TOKEN = "test-token";
  process.env.TELEGRAM_ALLOWED_USER_IDS = "123";
  try {
    await assert.rejects(() => processOne(job, directory), SyntaxError);
    const saved = JSON.parse(await readFile(join(directory, "job.json"), "utf8"));
    assert.equal(saved.documentMessageId, 99);
    assert.equal(saved.documentIsFallback, true);
    assert.deepEqual(replies, ["456", null]);
    assert.ok((await stat(join(directory, `meeting-${id}-transcript.docx`))).size > 5_000);
  } finally {
    globalThis.fetch = originalFetch;
    for (const [name, value] of [["ASSISTANT_DATA_DIR", originalDataDir], ["TELEGRAM_BOT_TOKEN", originalToken], ["TELEGRAM_ALLOWED_USER_IDS", originalAllowed]]) {
      if (value === undefined) delete process.env[name];
      else process.env[name] = value;
    }
    await rm(root, { recursive: true, force: true });
  }
});
