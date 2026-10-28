import { createHash } from "node:crypto";
import { readdir, readFile, rm, mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { cacheDir, writeFileAtomic } from "./config.ts";

export const SENT_WINDOW_MS = 10 * 60_000;
export const UNCERTAIN_WINDOW_MS = 24 * 60 * 60_000;
/** A reservation younger than this may still be in flight, so the message says so. */
export const IN_FLIGHT_WINDOW_MS = 30_000;

export type LedgerState = "reserved" | "sent" | "uncertain";
export type LedgerRecord = { state: LedgerState; at: string; id?: string; label?: string; conversationId?: string };

/** `images` and `files` hold the SHA-256 digest of each attachment. `sendAt` is set only for a scheduled send. Each joins the key only when present, so other keys do not change. */
export function ledgerKey(parts: { conversationId: string; content: string; mentionMris: string[]; contentType: string; images?: string[]; files?: string[]; sendAt?: number }): string {
  return createHash("sha256").update(JSON.stringify({
    conversationId: parts.conversationId,
    content: parts.content,
    mentionMris: parts.mentionMris,
    contentType: parts.contentType,
    ...(parts.sendAt === undefined ? {} : { sendAt: parts.sendAt }),
    ...(parts.images?.length ? { images: parts.images } : {}),
    ...(parts.files?.length ? { files: parts.files } : {}),
  })).digest("hex");
}

function dirFor(tenantId: string): string {
  return join(cacheDir(), tenantId, "ledger");
}

function pathFor(tenantId: string, key: string): string {
  if (!/^[0-9a-f]{64}$/.test(key)) throw new Error("Unknown ledger key. Run: teams-cli ledger list");
  return join(dirFor(tenantId), `${key}.json`);
}

/** Reads one record. A file that cannot be parsed is treated as a reservation made now. */
async function readRecord(path: string, now: number): Promise<LedgerRecord> {
  try {
    const parsed = JSON.parse(await readFile(path, "utf8"));
    if (parsed && typeof parsed.state === "string" && typeof parsed.at === "string") return parsed;
  } catch { /* A racing writer may still be filling the file. */ }
  return { state: "reserved", at: new Date(now).toISOString() };
}

/** Thrown when a send is refused. The message says what to do next. */
export class LedgerRefusal extends Error {
  constructor(message: string) {
    super(message);
    this.name = "LedgerRefusal";
  }
}

/**
 * Reserves a send before the POST. A file that exists refuses the send unless it is too old.
 * The exclusive create (`wx`) makes two racing sends safe: exactly one wins.
 */
export async function reserve(
  tenantId: string,
  key: string,
  context: { label: string; conversationId: string; input: string },
  now = Date.now(),
): Promise<void> {
  const path = pathFor(tenantId, key);
  await mkdir(dirFor(tenantId), { recursive: true, mode: 0o700 });
  const record: LedgerRecord = { state: "reserved", at: new Date(now).toISOString(), label: context.label, conversationId: context.conversationId };
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      await writeFile(path, `${JSON.stringify(record, null, 2)}\n`, { flag: "wx", mode: 0o600 });
      return;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    }
    const existing = await readRecord(path, now);
    const age = now - Date.parse(existing.at);
    if (existing.state === "sent" && age < SENT_WINDOW_MS) {
      throw new LedgerRefusal(`Identical message already sent to ${context.label} at ${new Date(existing.at).toLocaleString()} as ${existing.id}. Not sending again.`);
    }
    if (existing.state === "reserved" && age < IN_FLIGHT_WINDOW_MS) {
      throw new LedgerRefusal("Another identical send is in progress. Not sending again.");
    }
    if ((existing.state === "reserved" || existing.state === "uncertain") && age < UNCERTAIN_WINDOW_MS) {
      throw new LedgerRefusal(`A previous identical send may have posted (outcome unknown). Check with: teams-cli history ${context.input} --top=5. Then run: teams-cli ledger clear ${key}`);
    }
    await rm(path, { force: true });
  }
  throw new Error("Could not reserve this send. Run the command again.");
}

export async function markSent(tenantId: string, key: string, id: string, context: { label: string; conversationId: string }, now = Date.now()): Promise<void> {
  await writeFileAtomic(pathFor(tenantId, key), `${JSON.stringify({ state: "sent", at: new Date(now).toISOString(), id, ...context }, null, 2)}\n`, 0o600);
}

export async function markUncertain(tenantId: string, key: string, context: { label: string; conversationId: string }, now = Date.now()): Promise<void> {
  await writeFileAtomic(pathFor(tenantId, key), `${JSON.stringify({ state: "uncertain", at: new Date(now).toISOString(), ...context }, null, 2)}\n`, 0o600);
}

/** Removes a record. Used after a definite failure, where the message did not post. */
export async function removeRecord(tenantId: string, key: string): Promise<void> {
  await rm(pathFor(tenantId, key), { force: true });
}

export async function listRecords(tenantId: string): Promise<Array<LedgerRecord & { key: string }>> {
  let names: string[] = [];
  try {
    names = await readdir(dirFor(tenantId));
  } catch { /* No ledger yet. */ }
  const records = await Promise.all(names.filter(name => name.endsWith(".json")).map(async name => ({
    key: name.slice(0, -".json".length),
    ...(await readRecord(join(dirFor(tenantId), name), Date.now())),
  })));
  return records.sort((a, b) => a.at.localeCompare(b.at));
}

export async function clearRecord(tenantId: string, key: string): Promise<boolean> {
  const path = pathFor(tenantId, key);
  const existed = await readFile(path, "utf8").then(() => true, () => false);
  await rm(path, { force: true });
  return existed;
}
