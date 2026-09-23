import { randomUUID } from "node:crypto";
import { constants } from "node:fs";
import { copyFile, mkdir, open, readFile, realpath, stat, statfs, writeFile } from "node:fs/promises";
import { basename, extname, isAbsolute, join, relative, resolve, sep } from "node:path";

export type Mode = "bullets" | "transcript" | "both";
export type MeetingJob = {
  id: string;
  state: "queued" | "working" | "bullets_sent" | "done" | "failed";
  mode: Mode;
  chatId: string;
  messageId: number | null;
  fileName: string;
  fileId: string | null;
  sourcePath: string | null;
  createdAt: string;
  attempts: number;
  bulletsMessageId?: number;
  documentMessageId?: number;
  error?: string;
};

const MAX_INPUT_BYTES = 2_000_000_000;

export function dataRoot(env = process.env): string {
  const data = env.ASSISTANT_DATA_DIR?.trim();
  if (!data || !isAbsolute(data)) throw new Error("ASSISTANT_DATA_DIR must be absolute");
  return resolve(data, "meeting-transcription");
}

export function cleanName(value: string): string {
  const name = basename(value.normalize("NFC"));
  if (name !== value.normalize("NFC") || name.startsWith(".") || name.length > 180 || /[\x00-\x1f\x7f]/u.test(name))
    throw new Error("invalid audio filename");
  if (![".m4a", ".mp3"].includes(extname(name).toLowerCase()))
    throw new Error("only .m4a and .mp3 are supported");
  return name;
}

export function privateTarget(auth: {
  principalType?: string;
  attributes?: Readonly<Record<string, string | readonly string[] | undefined>>;
} | null): { chatId: string; messageId: number | null } {
  const attrs = auth?.attributes;
  const chatId = typeof attrs?.chat_id === "string" ? attrs.chat_id : "";
  const userId = typeof attrs?.user_id === "string" ? attrs.user_id : "";
  if (auth?.principalType !== "user" || attrs?.chat_type !== "private" || !chatId || chatId !== userId)
    throw new Error("meeting transcription requires the owner's private Telegram chat");
  const message = typeof attrs.message_id === "string" ? Number(attrs.message_id) : NaN;
  return { chatId, messageId: Number.isSafeInteger(message) && message > 0 ? message : null };
}

function inside(root: string, target: string): boolean {
  return target === root || target.startsWith(root + sep);
}

export async function stageAudio(pathFromTool: string, target: string, env = process.env): Promise<string> {
  if (isAbsolute(pathFromTool) || pathFromTool.split(/[\\/]/u).some((part) => part === ".." || part.startsWith(".")))
    throw new Error("audio path must be relative to vault/attachments");
  const vault = env.ASSISTANT_VAULT_DIR?.trim();
  if (!vault || !isAbsolute(vault)) throw new Error("ASSISTANT_VAULT_DIR must be absolute");
  const root = await realpath(resolve(vault, "attachments"));
  const source = await realpath(resolve(root, pathFromTool));
  if (!inside(root, source) || relative(root, source).startsWith(".."))
    throw new Error("audio path is outside vault/attachments");
  const metadata = await stat(source);
  if (!metadata.isFile() || metadata.size <= 0 || metadata.size > MAX_INPUT_BYTES)
    throw new Error("audio file has an unsupported size");
  const disk = await statfs(resolve(target, ".."));
  if (metadata.size > disk.bavail * disk.bsize * 0.8)
    throw new Error("not enough free disk space for audio");
  await copyFile(source, target, constants.COPYFILE_EXCL);
  return target;
}

export async function createJob(args: {
  mode: Mode;
  chatId: string;
  messageId: number | null;
  fileName: string;
  path?: string;
  fileId?: string;
}, env = process.env): Promise<MeetingJob> {
  const name = cleanName(args.fileName);
  if (Boolean(args.path) === Boolean(args.fileId))
    throw new Error("supply exactly one audio path or Telegram file ID");
  if (args.fileId && !/^[A-Za-z0-9_-]{10,512}$/u.test(args.fileId))
    throw new Error("invalid Telegram file ID");
  const id = randomUUID();
  const directory = join(dataRoot(env), "jobs", id);
  await mkdir(directory, { recursive: true, mode: 0o700 });
  const sourcePath = args.path
    ? await stageAudio(args.path, join(directory, "source" + extname(name).toLowerCase()), env)
    : null;
  const job: MeetingJob = {
    id, state: "queued", mode: args.mode, chatId: args.chatId,
    messageId: args.messageId, fileName: name, fileId: args.fileId ?? null,
    sourcePath, createdAt: new Date().toISOString(), attempts: 0,
  };
  await saveJob(job, env);
  return job;
}

export async function saveJob(job: MeetingJob, env = process.env): Promise<void> {
  const directory = join(dataRoot(env), "jobs", job.id);
  const temporary = join(directory, "job.json.tmp");
  await writeFile(temporary, JSON.stringify(job), { mode: 0o600 });
  const { rename } = await import("node:fs/promises");
  await rename(temporary, join(directory, "job.json"));
}

export async function loadJob(path: string): Promise<MeetingJob> {
  return JSON.parse(await readFile(path, "utf8")) as MeetingJob;
}

export async function tryLock(jobDirectory: string): Promise<(() => Promise<void>) | null> {
  const lockPath = join(jobDirectory, ".lock");
  try {
    const handle = await open(lockPath, "wx", 0o600);
    return async () => {
      await handle.close();
      const { unlink } = await import("node:fs/promises");
      await unlink(lockPath).catch(() => undefined);
    };
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    const metadata = await stat(lockPath).catch(() => null);
    if (metadata && Date.now() - metadata.mtimeMs > 60 * 60 * 1000) {
      const { unlink } = await import("node:fs/promises");
      await unlink(lockPath).catch(() => undefined);
    }
    return null;
  }
}
