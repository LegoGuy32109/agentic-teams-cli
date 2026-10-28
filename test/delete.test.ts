import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { deleteMessages, parseDeleteRefs } from "../src/commands/delete.ts";
import { setPaceForTests } from "../src/pace.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";

const BASE = "/api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages";
const A = "1791990424001";
const B = "1791990424002";
const C = "1791990424003";
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
const originalLog = console.log;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}

const note = (id: string, text: string, extra: Record<string, unknown> = {}) => ({
  status: 200,
  json: { id, content: `<p>${text}</p>`, imdisplayname: "Test User", messagetype: "RichText/Html", properties: {}, ...extra },
});
const gone = (id: string) => ({ status: 200, json: { id, content: "", imdisplayname: "Test User", messagetype: "RichText/Html", properties: { deletetime: "1791990500000" } } });
const deleted = { status: 200, text: "null" };
const deletes = () => fake!.requests.filter(item => item.method === "DELETE");

beforeEach(async () => {
  await resetLocalState();
  await seedProfile();
  logs = [];
  console.log = (...args: unknown[]) => { logs.push(args.map(String).join(" ")); };
  setPaceForTests({});
});

afterEach(() => {
  console.log = originalLog;
  fake?.uninstall();
  fake = undefined;
  process.exitCode = 0;
  setPaceForTests({});
});

describe("parseDeleteRefs", () => {
  test("accepts bare ids and refs to the same conversation, and drops repeats", () => {
    expect(parseDeleteRefs("48:notes", [`48:notes#${A}`, B, A])).toEqual([A, B]);
  });

  test("refuses a ref that names another conversation", () => {
    expect(() => parseDeleteRefs("48:notes", [`19:dev@thread.tacv2#${A}`])).toThrow("names another conversation");
  });

  test("refuses a short id and a ref without a message id", () => {
    expect(() => parseDeleteRefs("48:notes", ["12345"])).toThrow("is not a message id");
    expect(() => parseDeleteRefs("48:notes", ["48:notes#abc"])).toThrow("no valid message id");
  });
});

describe("preview", () => {
  test("lists each id with its text, sender and time, and sends no DELETE", async () => {
    serve({ [`GET ${BASE}/{id}`]: note(A, "first test line"), [`DELETE ${BASE}/{id}`]: deleted });
    await deleteMessages("me", [A], { json: false, allowWrite: false });
    expect(deletes()).toHaveLength(0);
    expect(logs[0]).toBe("Dry run (effect: message): would delete 1 message(s) in Notes to self:");
    expect(logs[1]).toContain(`${A}  "first test line" by Test User at`);
  });

  test("an already deleted message is listed but not counted", async () => {
    serve({ [`GET ${BASE}/${A}`]: note(A, "still here"), [`GET ${BASE}/${B}`]: gone(B) });
    await deleteMessages("me", [A, B], { json: false, allowWrite: false });
    expect(logs[0]).toBe("Dry run (effect: message): would delete 1 message(s) in Notes to self:");
    expect(logs).toContain(`  ${B}  already deleted`);
  });

  test("the JSON preview counts only the messages it would delete", async () => {
    serve({ [`GET ${BASE}/${A}`]: gone(A), [`GET ${BASE}/${B}`]: note(B, "live") });
    await deleteMessages("me", [A, B], { json: true, allowWrite: false });
    const out = JSON.parse(logs.join("\n"));
    expect(out.count).toBe(1);
    expect(out.messages.map((item: { status: string }) => item.status)).toEqual(["already deleted", "delete"]);
  });

  test("the text preview is the first 40 characters", async () => {
    const long = "x".repeat(60);
    serve({ [`GET ${BASE}/{id}`]: note(A, long) });
    await deleteMessages("me", [A], { json: true, allowWrite: false });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.messages[0].text).toBe("x".repeat(40));
    expect(preview.method).toBe("DELETE");
    expect(preview.query).toEqual({ behavior: "softDelete" });
  });

  test("refuses a ref from another conversation before any request", async () => {
    serve({ [`GET ${BASE}/{id}`]: note(A, "x") });
    await expect(deleteMessages("me", [`19:dev@thread.tacv2#${A}`], { json: false, allowWrite: false })).rejects.toThrow("names another conversation");
    expect(fake!.requests).toHaveLength(0);
  });
});

describe("write", () => {
  test("deletes each id in order with softDelete and prints one line per id", async () => {
    serve({
      [`GET ${BASE}/${A}`]: note(A, "one"), [`GET ${BASE}/${B}`]: note(B, "two"),
      [`DELETE ${BASE}/${A}`]: deleted, [`DELETE ${BASE}/${B}`]: deleted,
    });
    await deleteMessages("me", [`48:notes#${A}`, B], { json: false, allowWrite: true });
    expect(deletes().map(item => item.path)).toEqual([`${BASE}/${A}`, `${BASE}/${B}`]);
    expect(deletes()[0]!.query).toEqual({ behavior: "softDelete" });
    expect(deletes()[0]!.body).toBeUndefined();
    expect(logs).toEqual([`deleted ${A}`, `deleted ${B}`, "Deleted 2, skipped 0, failed 0."]);
    expect(process.exitCode).toBe(0);
  });

  test("a message that is already deleted is skipped with no DELETE", async () => {
    serve({ [`GET ${BASE}/${A}`]: gone(A), [`DELETE ${BASE}/${A}`]: deleted });
    await deleteMessages("me", [A], { json: true, allowWrite: true });
    expect(deletes()).toHaveLength(0);
    expect(JSON.parse(logs.join("\n"))).toEqual({ ok: true, deleted: [], skipped: [A], failed: [] });
  });

  test("a partial failure continues, reports it, and exits 1", async () => {
    serve({
      [`GET ${BASE}/${A}`]: note(A, "one"), [`GET ${BASE}/${B}`]: note(B, "two"), [`GET ${BASE}/${C}`]: note(C, "three"),
      [`DELETE ${BASE}/${A}`]: deleted,
      [`DELETE ${BASE}/${B}`]: { status: 403, json: { message: "Forbidden" } },
      [`DELETE ${BASE}/${C}`]: deleted,
    });
    await deleteMessages("me", [A, B, C], { json: true, allowWrite: true });
    expect(deletes()).toHaveLength(3);
    const result = JSON.parse(logs.join("\n"));
    expect(result.ok).toBe(false);
    expect(result.deleted).toEqual([A, C]);
    expect(result.failed).toHaveLength(1);
    expect(result.failed[0].id).toBe(B);
    expect(result.failed[0].error).toContain("403");
    expect(process.exitCode).toBe(1);
  });

  test("an unreadable message is failed and never deleted", async () => {
    serve({ [`GET ${BASE}/${A}`]: { status: 404, json: { message: "not found" } }, [`DELETE ${BASE}/${A}`]: deleted });
    await deleteMessages("me", [A], { json: false, allowWrite: true });
    expect(deletes()).toHaveLength(0);
    expect(logs[0]).toMatch(new RegExp(`^failed ${A}: `));
    expect(process.exitCode).toBe(1);
  });

  test("the pacing wait says request, not send, for a delete", async () => {
    const lines: string[] = [];
    const originalError = console.error;
    console.error = (...args: unknown[]) => { lines.push(args.map(String).join(" ")); };
    try {
      let now = 2_000_000;
      setPaceForTests({ now: () => now, sleep: async ms => { now += ms; } });
      const ids = Array.from({ length: 16 }, (_, index) => String(1791990440000 + index));
      const routes: Record<string, Route> = {};
      for (const id of ids) {
        routes[`GET ${BASE}/${id}`] = note(id, "x");
        routes[`DELETE ${BASE}/${id}`] = deleted;
      }
      serve(routes);
      await deleteMessages("me", ids, { json: true, allowWrite: true });
    } finally {
      console.error = originalError;
    }
    expect(lines.length).toBeGreaterThan(0);
    for (const line of lines) expect(line).toMatch(/before the next request\.$/);
  });

  test("paces deletes so the 16th waits for the first minute to pass", async () => {
    let now = 1_000_000;
    const waits: number[] = [];
    setPaceForTests({ now: () => now, sleep: async ms => { waits.push(ms); now += ms; } });
    const ids = Array.from({ length: 16 }, (_, index) => String(1791990430000 + index));
    const routes: Record<string, Route> = {};
    for (const id of ids) {
      routes[`GET ${BASE}/${id}`] = note(id, "x");
      routes[`DELETE ${BASE}/${id}`] = deleted;
    }
    serve(routes);
    await deleteMessages("me", ids, { json: false, allowWrite: true });
    expect(deletes()).toHaveLength(16);
    // Fifteen a minute and three in five seconds: the waits add up to the minute the first delete needs.
    expect(waits.reduce((total, ms) => total + ms, 0)).toBe(60_000);
  });
});
