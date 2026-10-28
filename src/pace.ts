import { mkdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { cacheDir, writeFileAtomic } from "./config.ts";

/** The service allows 20 posts a minute. 15 keeps a margin under that. */
export const WRITE_LIMIT_PER_MINUTE = 15;
/** The service also refuses more than 3 posts in 5 seconds (`3Per05Secs`). */
export const BURST_LIMIT = 3;
export const BURST_WINDOW_MS = 5_000;
const WINDOW_MS = 60_000;
const LOCK_WAIT_MS = 5_000;
const LOCK_STALE_MS = 10_000;
const NOTICE_EVERY_MS = 5_000;

let clock = () => Date.now();
let sleep = (ms: number) => Bun.sleep(ms);

/** Thrown by `--no-wait` when the limit is reached. The seconds say when to try again. */
export class RateLimitError extends Error {
  constructor(message: string, readonly retry_after_seconds: number) {
    super(message);
    this.name = "RateLimitError";
  }
}

/** Replaces the clock and the wait for tests. */
export function setPaceForTests(options: { now?: () => number; sleep?: (ms: number) => Promise<unknown> }): void {
  clock = options.now ?? (() => Date.now());
  sleep = (options.sleep ?? ((ms: number) => Bun.sleep(ms))) as (ms: number) => Promise<void>;
}

function pathsFor(tenantId: string) {
  const dir = join(cacheDir(), tenantId);
  return { dir, file: join(dir, "send-pace.json"), lock: join(dir, "send-pace.lock") };
}

async function readStamps(file: string): Promise<number[]> {
  try {
    const parsed = JSON.parse(await readFile(file, "utf8"));
    return Array.isArray(parsed?.sends) ? parsed.sends.filter((value: unknown): value is number => typeof value === "number") : [];
  } catch {
    return [];
  }
}

/** Takes the cross-process lock. A lock older than 10 seconds is stale and is removed. */
async function acquire(lock: string): Promise<void> {
  const deadline = Date.now() + LOCK_WAIT_MS;
  for (;;) {
    try {
      await writeFile(lock, String(process.pid), { flag: "wx", mode: 0o600 });
      return;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    }
    try {
      const info = await stat(lock);
      if (Date.now() - info.mtimeMs > LOCK_STALE_MS) {
        await rm(lock, { force: true });
        continue;
      }
    } catch {
      continue;
    }
    if (Date.now() >= deadline) throw new Error("Could not get the send lock. Try again.");
    await new Promise(resolve => setTimeout(resolve, 25));
  }
}

/** True when a write at `time` keeps both limits. Only the stamps before it count. */
function fitsAt(stamps: number[], time: number): boolean {
  const minute = stamps.filter(stamp => time - stamp < WINDOW_MS && stamp <= time).length;
  const burst = stamps.filter(stamp => time - stamp < BURST_WINDOW_MS && stamp <= time).length;
  return minute < WRITE_LIMIT_PER_MINUTE && burst < BURST_LIMIT;
}

/** The earliest time after `now` when a write fits. Each limit only changes when an old stamp expires. */
function nextFit(stamps: number[], now: number): number {
  const candidates = stamps.flatMap(stamp => [stamp + WINDOW_MS, stamp + BURST_WINDOW_MS]).filter(time => time > now).sort((a, b) => a - b);
  return candidates.find(time => fitsAt(stamps, time)) ?? now;
}

/**
 * Waits until one more send or delete fits both limits, then records it. The record is a file in
 * the tenant cache, so separate processes share it. `noWait` throws at once instead of waiting.
 */
export async function paceWrite(tenantId: string, options: { noWait?: boolean; send?: boolean } = {}): Promise<void> {
  const noun = options.send === false ? "request" : "send";
  const { dir, file, lock } = pathsFor(tenantId);
  await mkdir(dir, { recursive: true, mode: 0o700 });
  let lastNotice = Number.NEGATIVE_INFINITY;
  for (;;) {
    await acquire(lock);
    let now = 0;
    let stamps: number[] = [];
    let fit: number | undefined;
    try {
      now = clock();
      stamps = (await readStamps(file)).filter(stamp => now - stamp < WINDOW_MS);
      if (fitsAt(stamps, now)) {
        await writeFileAtomic(file, `${JSON.stringify({ sends: [...stamps, now] })}\n`, 0o600);
        return;
      }
      fit = nextFit(stamps, now);
    } finally {
      await rm(lock, { force: true });
    }
    const seconds = Math.max(1, Math.ceil((fit - now) / 1000));
    if (options.noWait) {
      throw new RateLimitError(`Rate limit reached: ${stamps.length} ${noun}s in the last minute. Try again in ${seconds}s.`, seconds);
    }
    if (now - lastNotice >= NOTICE_EVERY_MS) {
      console.error(`Rate limit: waiting ${seconds}s before the next ${noun}.`);
      lastNotice = now;
    }
    await sleep(Math.min(fit - now, 1000));
  }
}
