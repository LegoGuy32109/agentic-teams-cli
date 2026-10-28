import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { reactToMessage } from "../src/commands/react.ts";
import { resolveReactionKey } from "../src/reactions.ts";
import { setPaceForTests } from "../src/pace.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";
import { OID } from "./jwt.ts";

const BASE = "/api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages";
const A = "1791990424001";
const MINE = `8:orgid:${OID}`;
const OTHER = "8:orgid:00000000-0000-0000-0000-000000000009";
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
const originalLog = console.log;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}

/** A message with reactions. `users` maps a key to the user ids under it. */
const message = (users: Record<string, string[]> = {}, extra: Record<string, unknown> = {}) => ({
  status: 200,
  json: {
    id: A,
    content: "<p>note</p>",
    properties: { emotions: JSON.stringify(Object.entries(users).map(([key, mris]) => ({ key, users: mris.map(mri => ({ mri, time: 1, value: "x" })) }))), ...extra },
  },
});
const writes = () => fake!.requests.filter(item => item.method !== "GET");
const ok = { status: 200, text: "" };

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

describe("resolveReactionKey", () => {
  test("maps the common emoji, with or without the variation selector", () => {
    expect(["👍", "❤️", "❤", "😆", "😮", "😢", "😠", "✅"].map(resolveReactionKey)).toEqual(
      ["like", "heart", "heart", "laugh", "surprised", "sad", "angry", "2705_whiteheavycheckmark"]);
  });

  test("passes named, codepoint and custom keys unchanged", () => {
    const custom = "party-blob;0-cus-d6-a89232cada85823f8811d01327be0bcf";
    expect(resolveReactionKey("laugh")).toBe("laugh");
    expect(resolveReactionKey("2705_whiteheavycheckmark")).toBe("2705_whiteheavycheckmark");
    expect(resolveReactionKey(custom)).toBe(custom);
  });

  test("maps any catalog emoji, key, or shortcut, with skin tones", () => {
    expect(resolveReactionKey("🚀")).toBe("launch");
    expect(resolveReactionKey("👍🏽")).toBe("yes-tone3");
    expect(resolveReactionKey(":launch:")).toBe("launch");
    expect(resolveReactionKey("(launch)")).toBe("launch");
    expect(resolveReactionKey("🦄")).toMatch(/^[\w-]+$/);
  });

  test("refuses input that is neither an emoji nor a key", () => {
    expect(() => resolveReactionKey("not a key!")).toThrow("is not in the Teams emoji catalog");
  });
});

describe("react", () => {
  test("the preview shows the PUT body and sends no write", async () => {
    serve({ [`GET ${BASE}/${A}`]: message() });
    await reactToMessage("me", A, "👍", false, { json: true, allowWrite: false });
    const out = JSON.parse(logs.join("\n"));
    expect(out).toMatchObject({ dry_run: true, effect: "reaction", method: "PUT", key: "like", path: `/api/chatsvc/amer/v1/users/ME/conversations/48%3Anotes/messages/${A}/properties?name=emotions` });
    expect(out.body.emotions.key).toBe("like");
    expect(typeof out.body.emotions.value).toBe("number");
    expect(writes()).toHaveLength(0);
  });

  test("PUT bodies for a named, a codepoint and a custom key", async () => {
    const custom = "party-blob;0-cus-d6-a89232cada85823f8811d01327be0bcf";
    serve({ [`GET ${BASE}/${A}`]: message(), [`PUT ${BASE}/${A}/properties`]: ok });
    for (const input of ["❤️", "✅", custom]) await reactToMessage("me", A, input, false, { json: false, allowWrite: true });
    expect(writes().map(item => item.method)).toEqual(["PUT", "PUT", "PUT"]);
    expect(writes().map(item => (item.body as any).emotions.key)).toEqual(["heart", "2705_whiteheavycheckmark", custom]);
    for (const item of writes()) {
      expect(item.query).toEqual({ name: "emotions" });
      expect(Object.keys((item.body as any).emotions)).toEqual(["key", "value"]);
      expect(Math.abs((item.body as any).emotions.value - Date.now())).toBeLessThan(10_000);
    }
  });

  test("already reacted: says so and sends no write, even with --allow-write", async () => {
    serve({ [`GET ${BASE}/${A}`]: message({ like: [MINE, OTHER] }), [`PUT ${BASE}/${A}/properties`]: ok });
    await reactToMessage("me", A, "👍", false, { json: false, allowWrite: true });
    expect(logs[0]).toContain('You already reacted with "like"');
    expect(writes()).toHaveLength(0);
  });

  test("another person's reaction under the key does not count as mine", async () => {
    serve({ [`GET ${BASE}/${A}`]: message({ like: [OTHER] }), [`PUT ${BASE}/${A}/properties`]: ok });
    await reactToMessage("me", A, "👍", false, { json: false, allowWrite: true });
    expect(writes()).toHaveLength(1);
  });

  test("a deleted message is refused", async () => {
    serve({ [`GET ${BASE}/${A}`]: message({}, { deletetime: "1791990500000" }) });
    await expect(reactToMessage("me", A, "👍", false, { json: false, allowWrite: true })).rejects.toThrow("deleted");
    expect(writes()).toHaveLength(0);
  });

  test("--verify reads the message back and checks my id is under the key", async () => {
    serve({
      [`GET ${BASE}/${A}`]: [message(), message({ like: [MINE] })],
      [`PUT ${BASE}/${A}/properties`]: ok,
    });
    await reactToMessage("me", A, "👍", false, { json: true, allowWrite: true, verify: true });
    expect(JSON.parse(logs.join("\n")).verification).toEqual({ ok: true });
    expect(process.exitCode).toBe(0);
  });

  test("--verify fails when the read-back lacks the reaction", async () => {
    serve({ [`GET ${BASE}/${A}`]: message(), [`PUT ${BASE}/${A}/properties`]: ok });
    await reactToMessage("me", A, "👍", false, { json: true, allowWrite: true, verify: true });
    expect(JSON.parse(logs.join("\n")).verification).toEqual({ ok: false });
    expect(process.exitCode).toBe(1);
  });

  test("an uncertain write is reported as unknown and not retried", async () => {
    serve({ [`GET ${BASE}/${A}`]: message(), [`PUT ${BASE}/${A}/properties`]: { status: 503, json: {} } });
    await expect(reactToMessage("me", A, "👍", false, { json: false, allowWrite: true })).rejects.toThrow("Outcome unknown");
    expect(writes()).toHaveLength(1);
  });

  test("a stream that is read-only is refused before any request", async () => {
    serve({});
    await expect(reactToMessage("mentions", A, "👍", false, { json: false, allowWrite: true })).rejects.toThrow("read-only");
    expect(fake!.requests).toHaveLength(0);
  });
});

describe("unreact", () => {
  test("sends DELETE with only the key, as text/plain like the web client", async () => {
    serve({ [`GET ${BASE}/${A}`]: message({ like: [MINE] }), [`DELETE ${BASE}/${A}/properties`]: { status: 200, text: "null" } });
    await reactToMessage("me", A, "👍", true, { json: false, allowWrite: true });
    const [request] = writes();
    expect(request).toMatchObject({ method: "DELETE", query: { name: "emotions" }, body: { emotions: { key: "like" } } });
    expect(request!.headers["content-type"]).toBe("text/plain;charset=UTF-8");
    expect(logs[0]).toContain('Removed "like"');
  });

  test("not reacted: says so and sends no write", async () => {
    serve({ [`GET ${BASE}/${A}`]: message({ like: [OTHER] }), [`DELETE ${BASE}/${A}/properties`]: ok });
    await reactToMessage("me", A, "👍", true, { json: false, allowWrite: true });
    expect(logs[0]).toContain('You have no "like" reaction');
    expect(writes()).toHaveLength(0);
  });

  test("the preview sends no DELETE", async () => {
    serve({ [`GET ${BASE}/${A}`]: message({ like: [MINE] }), [`DELETE ${BASE}/${A}/properties`]: ok });
    await reactToMessage("me", A, "👍", true, { json: true, allowWrite: false });
    expect(JSON.parse(logs.join("\n"))).toMatchObject({ dry_run: true, method: "DELETE", body: { emotions: { key: "like" } } });
    expect(writes()).toHaveLength(0);
  });
});
