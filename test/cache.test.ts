import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { cacheUsers } from "../src/commands/cache.ts";
import { clearUserCache } from "../src/users.ts";
import { installFakeTeams, resetLocalState, seedProfile } from "./fake-teams.ts";
import { TID } from "./jwt.ts";

let logs: string[] = [];
const originalLog = console.log;
let fake: ReturnType<typeof installFakeTeams> | undefined;

beforeEach(async () => {
  await resetLocalState();
  await seedProfile();
  logs = [];
  console.log = (...args: unknown[]) => { logs.push(args.map(String).join(" ")); };
});

afterEach(() => {
  console.log = originalLog;
  fake?.uninstall();
  fake = undefined;
});

describe("cache users --refresh", () => {
  test("reports zero and succeeds when no user cache exists", async () => {
    expect(await clearUserCache(TID)).toBe(0);
    await cacheUsers();
    expect(logs.at(-1)).toBe("Cleared 0 cached users. Names refill on demand.");
  });

  test("sends no network request", async () => {
    fake = installFakeTeams({});
    await cacheUsers();
    expect(fake.requests.length).toBe(0);
  });
});
