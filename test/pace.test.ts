import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { mkdir, readFile, utimes, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { cacheDir } from "../src/config.ts";
import { paceWrite, RateLimitError, setPaceForTests, WRITE_LIMIT_PER_MINUTE } from "../src/pace.ts";
import { send } from "../src/commands/send.ts";
import { request, setSleepForTests, TeamsError } from "../src/transport.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Reply } from "./fake-teams.ts";
import { TID } from "./jwt.ts";

const TENANT_DIR = join(cacheDir(), TID);
const PACE_FILE = join(TENANT_DIR, "send-pace.json");
const LOCK = join(TENANT_DIR, "send-pace.lock");
const SEND = "POST /api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages";
const CHAT_PATH = "/api/chatsvc/amer/v1/users/ME/properties";
let fake: ReturnType<typeof installFakeTeams> | undefined;
let errors: string[] = [];
const originalError = console.error;

/** A fake clock. `sleep` moves it forward, so a wait costs no real time. */
function fakeClock(start = 1_000_000) {
  let now = start;
  const waits: number[] = [];
  setPaceForTests({
    now: () => now,
    sleep: async (ms: number) => { waits.push(ms); now += ms; },
  });
  return { waits, advance: (ms: number) => { now += ms; }, get now() { return now; } };
}

async function stamps(): Promise<number[]> {
  return JSON.parse(await readFile(PACE_FILE, "utf8")).sends;
}

beforeEach(async () => {
  await resetLocalState();
  await seedProfile();
  errors = [];
  console.error = (...args: unknown[]) => { errors.push(args.map(String).join(" ")); };
  setSleepForTests(async () => {});
});

afterEach(() => {
  console.error = originalError;
  fake?.uninstall();
  fake = undefined;
  setPaceForTests({});
  setSleepForTests(ms => Bun.sleep(ms));
});

describe("paceWrite", () => {
  test("20 sends over one minute never exceed 15 in any minute or 3 in any 5 seconds", async () => {
    const clock = fakeClock();
    const sent: number[] = [];
    for (let index = 0; index < 20; index++) {
      await paceWrite(TID);
      sent.push(clock.now);
    }
    expect(sent.length).toBe(20);
    for (const time of sent) {
      expect(sent.filter(other => other <= time && time - other < 60_000).length).toBeLessThanOrEqual(WRITE_LIMIT_PER_MINUTE);
      expect(sent.filter(other => other <= time && time - other < 5_000).length).toBeLessThanOrEqual(3);
    }
    // The 16th send waits for the first send to leave the minute.
    expect(sent[15]! - sent[0]!).toBeGreaterThanOrEqual(60_000);
  });

  test("the pace file holds the timestamps and is owner-only", async () => {
    fakeClock(5_000_000);
    await paceWrite(TID);
    expect(await stamps()).toEqual([5_000_000]);
    const mode = (await import("node:fs")).statSync(PACE_FILE).mode & 0o777;
    expect(mode).toBe(0o600);
  });

  test("--no-wait throws at once with the seconds to wait", async () => {
    const start = 1_000_000;
    fakeClock(start);
    await mkdir(TENANT_DIR, { recursive: true });
    // Fifteen sends in the last minute, four seconds apart, so the burst limit is not the cause.
    await writeFile(PACE_FILE, JSON.stringify({ sends: Array.from({ length: 15 }, (_, index) => start - 56_000 + index * 4_000) }));
    const refused = await paceWrite(TID, { noWait: true }).then(() => undefined, (error: RateLimitError) => error);
    expect(refused).toBeInstanceOf(RateLimitError);
    expect(refused?.message).toBe("Rate limit reached: 15 sends in the last minute. Try again in 4s.");
    expect(refused?.retry_after_seconds).toBe(4);
  });

  test("the wait prints one stderr line, updated at most once per 5 seconds", async () => {
    fakeClock();
    for (let index = 0; index < 3; index++) await paceWrite(TID);
    // Burst is full. The fourth send waits 5 seconds and prints once.
    await paceWrite(TID);
    expect(errors).toEqual(["Rate limit: waiting 5s before the next send."]);
  });

  test("two processes share one pace file", async () => {
    const env = { ...process.env, XDG_CACHE_HOME: process.env.XDG_CACHE_HOME! };
    const worker = join(import.meta.dir, "pace-worker.ts");
    const run = async (count: number) => {
      const proc = Bun.spawn([process.execPath, worker, TID, String(count)], { env, stdout: "pipe", stderr: "pipe" });
      const out = await new Response(proc.stdout).text();
      await proc.exited;
      return JSON.parse(out) as { accepted: number; refused: number };
    };
    // Three back to back in one process fill the 5 second burst.
    expect(await run(3)).toEqual({ accepted: 3, refused: 0 });
    // A second process sees the first one's three sends and is refused.
    expect(await run(1)).toEqual({ accepted: 0, refused: 1 });
    expect((await stamps()).length).toBe(3);
  });

  test("a lock older than 10 seconds is stale and is taken over", async () => {
    fakeClock(Date.now());
    await mkdir(TENANT_DIR, { recursive: true });
    await writeFile(LOCK, "999999");
    const old = new Date(Date.now() - 20_000);
    await utimes(LOCK, old, old);
    await paceWrite(TID);
    expect((await stamps()).length).toBe(1);
  });
});

describe("send with the pace file", () => {
  const plain = { json: false, verify: false, allowWrite: true } as const;
  const ok = { status: 201, json: { OriginalArrivalTime: 1791990424088 } };

  test("--no-wait refuses a send over the limit before any request", async () => {
    fakeClock(Date.now());
    await mkdir(TENANT_DIR, { recursive: true });
    const now = Date.now();
    await writeFile(PACE_FILE, JSON.stringify({ sends: Array.from({ length: 15 }, (_, index) => now - 30_000 + index * 1_000) }));
    fake = installFakeTeams({ [SEND]: ok });
    const error = await send("me", ["hello"], { ...plain, noWait: true }).then(() => undefined, (e: RateLimitError) => e);
    expect(error).toBeInstanceOf(RateLimitError);
    expect(error?.retry_after_seconds).toBeGreaterThan(0);
    expect(fake.requests.filter(item => item.method === "POST")).toHaveLength(0);
  });

  test("a 429 with Retry-After prints the wait, retries, and posts", async () => {
    fakeClock();
    fake = installFakeTeams({
      [SEND]: [{ status: 429, headers: { "retry-after": "7" }, json: { errorCode: 429, message: "API calls quota exceeded! 20Per01Min" } }, ok],
    });
    await send("me", ["hello"], plain);
    expect(fake.requests.filter(item => item.method === "POST")).toHaveLength(2);
    expect(errors).toEqual(["Rate limit: waiting 7s before the next send."]);
  });

  test("four 429s refuse the send, without the raw body, and leave no ledger record", async () => {
    fakeClock();
    const limited: Reply = { status: 429, headers: { "retry-after": "57" }, json: { errorCode: 429, message: "API calls quota exceeded! 20Per01Min" } };
    fake = installFakeTeams({ [SEND]: [limited] });
    const error = await send("me", ["hello"], plain).then(() => undefined, (e: TeamsError) => e);
    expect(error?.message).toBe("Teams refused the send: 20 per minute quota. Try again in 57s.");
    expect(error?.message).not.toContain("errorCode");
    expect(error?.uncertain).toBe(false);
    expect(fake.requests.filter(item => item.method === "POST")).toHaveLength(4);
  });
});
