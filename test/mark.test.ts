import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { horizonValue, markRead } from "../src/commands/mark.ts";
import { history } from "../src/commands/history.ts";
import { setPaceForTests } from "../src/pace.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";

const CHAT = "/api/chatsvc/amer/v1/users/ME/conversations";
const ME = "GET /api/csa/amer/api/v3/teams/users/me";
const CHANNEL = "19:fakechannel@thread.tacv2";
const GROUP = "19:fakegroup@thread.v2";
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
const originalLog = console.log;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}

const page = (...messages: Array<{ id: string; clientmessageid?: string }>) => ({ status: 200, json: { messages: messages.map(item => ({ messagetype: "RichText/Html", content: "<p>x</p>", properties: {}, ...item })) } });
const ok = { status: 200, text: "" };
const writes = () => fake!.requests.filter(item => item.method !== "GET" && item.path.includes("/conversations/"));

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

test("horizonValue joins the message id, now and the client message id", () => {
  expect(horizonValue({ id: "1792420401228", clientMessageId: "870913963303404111" }, 1792421301827)).toBe("1792420401228;1792421301827;870913963303404111");
});

describe("mark", () => {
  test("the preview shows the horizon from the newest message and never writes", async () => {
    serve({
      [`GET ${CHAT}/${CHANNEL}/messages`]: page({ id: "1791990424001", clientmessageid: "111" }, { id: "1791990424009", clientmessageid: "999" }, { id: "1791990424005", clientmessageid: "555" }),
      [`PUT ${CHAT}/${CHANNEL}/properties`]: ok,
    });
    await markRead(CHANNEL, { json: true, allowWrite: false });
    const out = JSON.parse(logs.join("\n"));
    expect(out).toMatchObject({ dry_run: true, effect: "read-state", method: "PUT" });
    expect(out.conversations[0].body.consumptionhorizon).toMatch(/^1791990424009;\d{13};999$/);
    expect(out.conversations[0].path).toContain("/properties?name=consumptionhorizon");
    expect(writes()).toHaveLength(0);
  });

  test("--allow-write sends the horizon to the properties endpoint", async () => {
    serve({
      [`GET ${CHAT}/${CHANNEL}/messages`]: page({ id: "1791990424009", clientmessageid: "999" }),
      [`PUT ${CHAT}/${CHANNEL}/properties`]: ok,
    });
    await markRead(CHANNEL, { json: false, allowWrite: true });
    const [request] = writes();
    expect(request).toMatchObject({ method: "PUT", query: { name: "consumptionhorizon" } });
    expect((request!.body as any).consumptionhorizon).toMatch(/^1791990424009;\d{13};999$/);
    expect(logs.at(-1)).toBe("Marked 1, skipped 0, failed 0.");
  });

  test("a message with no client message id is passed over for the next newest", async () => {
    serve({ [`GET ${CHAT}/${CHANNEL}/messages`]: page({ id: "1791990424009" }, { id: "1791990424005", clientmessageid: "555" }) });
    await markRead(CHANNEL, { json: true, allowWrite: false });
    expect(JSON.parse(logs.join("\n")).conversations[0].body.consumptionhorizon).toMatch(/^1791990424005;/);
  });

  test("an empty conversation is skipped with no write", async () => {
    serve({ [`GET ${CHAT}/${CHANNEL}/messages`]: page(), [`PUT ${CHAT}/${CHANNEL}/properties`]: ok });
    await markRead(CHANNEL, { json: false, allowWrite: true });
    expect(writes()).toHaveLength(0);
    expect(logs.join("\n")).toContain("no messages, skipped");
  });

  test("a read-only stream is refused before any request", async () => {
    serve({});
    await expect(markRead("mentions", { json: false, allowWrite: true })).rejects.toThrow("read-only");
    expect(fake!.requests).toHaveLength(0);
  });

  test("--all-unread previews every conversation that unread lists, and writes none", async () => {
    serve({
      [ME]: { status: 200, json: { teams: [], chats: [{ id: CHANNEL, title: "One", isRead: false }, { id: GROUP, title: "Two", isRead: false }, { id: "19:read@thread.v2", title: "Three", isRead: true }] } },
      [`GET ${CHAT}/${CHANNEL}/messages`]: page({ id: "1791990424001", clientmessageid: "1" }),
      [`GET ${CHAT}/${GROUP}/messages`]: page({ id: "1791990424002", clientmessageid: "2" }),
      [`PUT ${CHAT}/{id}/properties`]: ok,
    });
    await markRead(undefined, { json: true, allowWrite: false, allUnread: true });
    const out = JSON.parse(logs.join("\n"));
    expect(out.conversations.map((item: any) => [item.id, item.status])).toEqual([[CHANNEL, "would mark"], [GROUP, "would mark"]]);
    expect(writes()).toHaveLength(0);
  });

  test("--all-unread with --allow-write marks each one and keeps going after a failure", async () => {
    serve({
      [ME]: { status: 200, json: { teams: [], chats: [{ id: CHANNEL, title: "One", isRead: false }, { id: GROUP, title: "Two", isRead: false }] } },
      [`GET ${CHAT}/${CHANNEL}/messages`]: page({ id: "1791990424001", clientmessageid: "1" }),
      [`GET ${CHAT}/${GROUP}/messages`]: page({ id: "1791990424002", clientmessageid: "2" }),
      [`PUT ${CHAT}/${CHANNEL}/properties`]: { status: 400, json: {} },
      [`PUT ${CHAT}/${GROUP}/properties`]: ok,
    });
    await markRead(undefined, { json: true, allowWrite: true, allUnread: true });
    const out = JSON.parse(logs.join("\n"));
    expect(out.conversations.map((item: any) => item.status)).toEqual(["failed", "marked"]);
    expect(process.exitCode).toBe(1);
  });
});

describe("history mentions --unread", () => {
  const mention = (id: string, thread: string, source: string) => ({
    id, messagetype: "RichText/Html", content: "<p>@you</p>",
    properties: { activity: JSON.stringify({ sourceThreadId: thread, sourceMessageId: Number(source) }) },
  });
  const run = async (options: { unread?: boolean } = { unread: true }) => {
    await history("mentions", { json: true, content: false, ...options });
    return JSON.parse(logs.join("\n")).messages as Array<{ ts: string; source?: any }>;
  };

  test("keeps only mentions newer than the horizon of their conversation", async () => {
    serve({
      [ME]: { status: 200, json: { teams: [], chats: [
        { id: CHANNEL, title: "One", consumptionHorizon: { originalArrivalTime: 1791990000500 } },
        { id: GROUP, title: "Two", consumptionHorizon: { originalArrivalTime: 1791990000500 } },
      ] } },
      [`GET ${CHAT}/48:mentions/messages`]: { status: 200, json: { messages: [
        mention("1791990001000", CHANNEL, "1791990000400"),
        mention("1791990002000", CHANNEL, "1791990000900"),
        mention("1791990003000", GROUP, "1791990000500"),
        mention("1791990004000", GROUP, "1791990000501"),
      ] } },
    });
    const messages = await run();
    expect(messages.map(item => item.ts)).toEqual(["1791990002000", "1791990004000"]);
    expect(messages[0]!.source).toEqual({ conversation_id: CHANNEL, message_id: "1791990000900", label: "One", horizon: "1791990000500" });
  });

  test("a mention in a conversation with no known horizon stays, with a null horizon", async () => {
    serve({
      [ME]: { status: 200, json: { teams: [], chats: [] } },
      [`GET ${CHAT}/48:mentions/messages`]: { status: 200, json: { messages: [mention("1791990001000", CHANNEL, "1791990000400")] } },
    });
    const messages = await run();
    expect(messages).toHaveLength(1);
    expect(messages[0]!.source.horizon).toBeNull();
  });

  test("without --unread every mention is listed and the conversation list is not read", async () => {
    serve({
      [ME]: { status: 200, json: { teams: [], chats: [{ id: CHANNEL, title: "One", consumptionHorizon: { originalArrivalTime: 1791990009999 } }] } },
      [`GET ${CHAT}/48:mentions/messages`]: { status: 200, json: { messages: [mention("1791990001000", CHANNEL, "1791990000400")] } },
    });
    const messages = await run({});
    expect(messages).toHaveLength(1);
    expect(fake!.requests.some(item => item.path.includes("/teams/users/me"))).toBe(false);
  });

  test("--unread is refused for any other conversation", async () => {
    serve({});
    await expect(history("me", { json: true, content: false, unread: true })).rejects.toThrow("--unread works only with the mentions stream");
  });
});

describe("command line", () => {
  test("react and unreact need exactly three arguments, and mark takes one or --all-unread", async () => {
    const { runCli } = await import("./cli.ts");
    expect((await runCli(["react", "me", "123"])).stderr).toContain("Usage: teams-cli react <conversation> <message-id-or-ref> <emoji>");
    expect((await runCli(["unreact", "me"])).stderr).toContain("Usage: teams-cli unreact");
    expect((await runCli(["mark"])).stderr).toContain("Usage: teams-cli mark <conversation>");
    expect((await runCli(["mark", "me", "--all-unread"])).stderr).toContain("Usage: teams-cli mark --all-unread");
  });

  test("--unread is valid only with history, and --all-unread only with mark", async () => {
    const { runCli } = await import("./cli.ts");
    expect((await runCli(["send", "me", "hi", "--unread"])).stderr).toContain("Flag `--unread` is not valid with `teams-cli send`");
    expect((await runCli(["history", "me", "--all-unread"])).stderr).toContain("Flag `--all-unread` is not valid with `teams-cli history`");
  });
});
