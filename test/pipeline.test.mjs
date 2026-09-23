import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";

import { cleanName, privateTarget } from "../plugin/sh.iva/extension/lib/jobs.ts";
import { createWord, summarize, timecode, transcriptLines } from "../plugin/sh.iva/extension/lib/pipeline.ts";

test("the recipient comes only from an authenticated private owner turn", () => {
  const auth = { principalType: "user", attributes: { chat_id: "123", user_id: "123", chat_type: "private", message_id: "456" } };
  assert.deepEqual(privateTarget(auth), { chatId: "123", messageId: 456 });
  assert.throws(() => privateTarget({ ...auth, attributes: { ...auth.attributes, chat_type: "group" } }));
  assert.throws(() => privateTarget({ ...auth, attributes: { ...auth.attributes, user_id: "789" } }));
});

test("audio filenames are constrained and timing preserves both ends", () => {
  assert.equal(cleanName("Встреча.m4a"), "Встреча.m4a");
  assert.equal(cleanName("meeting.MP3"), "meeting.MP3");
  assert.throws(() => cleanName("../secret.mp3"));
  assert.throws(() => cleanName("meeting.exe"));
  assert.equal(timecode(1576.8), "00:26:16");
  assert.deepEqual(transcriptLines([{ start: 1.2, end: 3.9, speaker: 1, transcript: "Привет." }]), ["[00:00:01–00:00:03] Спикер 2: Привет."]);
});

test("all three Word modes produce nonempty DOCX files", async () => {
  const directory = await mkdtemp(join(tmpdir(), "iva-meeting-test-"));
  const summary = {
    bullets: ["Обсудили вопрос [00:00:01].", "Назвали следующий шаг [00:00:02]."],
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
    assert.match(prompt, /\[00:00:01\] Первый вопрос/u);
    return JSON.stringify({
      overview: "Обсудили два вопроса.", agenda: ["Первый вопрос", "Второй вопрос"],
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
});
