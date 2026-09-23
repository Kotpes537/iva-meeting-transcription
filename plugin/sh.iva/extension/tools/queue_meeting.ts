import { defineTool } from "eve/tools";
import { z } from "zod";
import { createJob, privateTarget } from "../lib/jobs.ts";

export default defineTool({
  description: "Queue a supported audio meeting recording from the owner's private Telegram chat. Background processing sends summary bullets and a Word report containing the full transcript to the same chat.",
  inputSchema: z.object({
    path: z.string().optional().describe("Source path relative to vault/attachments for a file already saved by Iva"),
    file_id: z.string().optional().describe("Telegram file_id provided by the inbound channel for a large file"),
    file_name: z.string().describe("Original audio filename (.m4a, .mp3, .wav, .flac, .ogg, .oga, .opus, .aac, .webm or .mp4)"),
    mode: z.enum(["bullets", "transcript", "both"]).default("both"),
  }),
  async execute(input, ctx) {
    try {
      const target = privateTarget(ctx.session.auth.current);
      const job = await createJob({
        mode: input.mode,
        chatId: target.chatId,
        messageId: target.messageId,
        fileName: input.file_name,
        path: input.path,
        fileId: input.file_id,
      });
      return { ok: true as const, job_id: job.id, mode: job.mode, status: "queued" as const };
    } catch (error) {
      const message = error instanceof Error ? error.message : "unknown error";
      const safe = new Set([
        "meeting transcription requires the owner's private Telegram chat",
        "invalid audio filename", "unsupported audio format",
        "supply exactly one audio path or Telegram file ID",
        "invalid Telegram file ID", "audio path must be relative to vault/attachments",
        "audio path is outside vault/attachments", "audio file has an unsupported size",
        "not enough free disk space for audio", "ASSISTANT_DATA_DIR must be absolute",
        "ASSISTANT_VAULT_DIR must be absolute",
      ]);
      return { ok: false as const, error: safe.has(message) ? message : "meeting job could not be queued" };
    }
  },
});
