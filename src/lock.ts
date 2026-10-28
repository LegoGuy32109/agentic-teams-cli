import { mkdir, open, stat, unlink } from "node:fs/promises";
import { dirname } from "node:path";
import { credentialsPath } from "./config.ts";

/**
 * The renewal lock serializes token rotation across CLI processes. A browser renewal can run for
 * about two minutes, so a waiting process outlasts it and a stale lock is one older than three.
 */
export const LOCK_WAIT_MS = 150_000;
export const LOCK_STALE_MS = 180_000;

export async function locked<T>(work: () => Promise<T>): Promise<T> {
  const path = `${credentialsPath()}.refresh.lock`;
  const deadline = Date.now() + LOCK_WAIT_MS;
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  for (;;) {
    try {
      const file = await open(path, "wx", 0o600);
      try { return await work(); } finally { await file.close(); await unlink(path).catch(() => {}); }
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
      const info = await stat(path).catch(() => undefined);
      if (info && Date.now() - info.mtimeMs > LOCK_STALE_MS) { await unlink(path).catch(() => {}); continue; }
      if (Date.now() >= deadline) throw new Error("Another token renewal is still running. Try again shortly.");
      await Bun.sleep(100);
    }
  }
}
