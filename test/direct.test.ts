import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { send } from "../src/commands/send.ts";
import { resolveConversation } from "../src/conversations.ts";
import { createDirectChat, directChatId } from "../src/direct.ts";
import { cacheDir } from "../src/config.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";
import { TID } from "./jwt.ts";
import searchV2 from "./fixtures/search-v2.json";
import teamsMe from "./fixtures/teams-me.json";

const ME = "8:orgid:11111111-1111-1111-1111-111111111111";
const ALEX = "8:orgid:33333333-3333-3333-3333-333333333333";
const CHAT = "19:11111111-1111-1111-1111-111111111111_33333333-3333-3333-3333-333333333333@unq.gbl.spaces";
const SEARCH = "POST /api/mt/amer/beta/users/searchV2";
const CSA = "GET /api/csa/amer/api/v3/teams/users/me";
const THREADS = "POST /api/chatsvc/amer/v1/threads";
const MESSAGE = "POST /api/chatsvc/amer/v1/users/ME/conversations/{chat}/messages";
const created = { status: 201, json: { threadResource: { id: CHAT, type: "Thread", properties: { threadType: "chat" } } } };
const sent = { status: 201, json: { OriginalArrivalTime: 1792169149570 } };
const noChats = { ...teamsMe, chats: [] };

let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
const originalLog = console.log;
const originalError = console.error;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams({ [SEARCH]: { status: 200, json: searchV2 }, [CSA]: { status: 200, json: noChats }, ...routes });
}
const calls = (route: string) => fake!.requests.filter(item => `${item.method} ${item.path}`.startsWith(route.split("{")[0]!) && item.method === route.split(" ")[0]);

beforeEach(async () => {
  await resetLocalState();
  await seedProfile();
  logs = [];
  console.log = (...args: unknown[]) => { logs.push(args.map(String).join(" ")); };
  console.error = () => {};
});
afterEach(() => {
  console.log = originalLog;
  console.error = originalError;
  fake?.uninstall();
  fake = undefined;
  process.exitCode = 0;
});

const plain = { json: false, verify: false, allowWrite: false } as const;

describe("the chat id", () => {
  test("it is the two object ids in ascending order, whichever way they are given", () => {
    expect(directChatId(ME, ALEX)).toBe(CHAT);
    expect(directChatId(ALEX, ME)).toBe(CHAT);
    expect(directChatId("8:orgid:BBBBBBBB-0000-0000-0000-000000000000", "8:orgid:aaaaaaaa-0000-0000-0000-000000000000"))
      .toBe("19:aaaaaaaa-0000-0000-0000-000000000000_bbbbbbbb-0000-0000-0000-000000000000@unq.gbl.spaces");
  });
});

describe("creating the chat", () => {
  test("the request is the one the Teams web client sends, and it returns the chat id", async () => {
    serve({ [THREADS]: created });
    expect(await createDirectChat(ME, ALEX)).toBe(CHAT);
    expect(fake!.requests).toHaveLength(1);
    expect(fake!.requests[0]!.body).toEqual({
      members: [{ id: ME, role: "Admin" }, { id: ALEX, role: "Admin" }],
      properties: { threadType: "chat", fixedRoster: true, uniquerosterthread: true },
    });
  });

  test("for a pair that already has a chat the service answers with an empty object, and the id is the derived one", async () => {
    serve({ [THREADS]: { status: 200, json: {} } });
    expect(await createDirectChat(ME, ALEX)).toBe(CHAT);
    serve({ [THREADS]: { status: 200, json: { threadResource: undefined } } });
    expect(await createDirectChat(ALEX, ME)).toBe(CHAT);
  });

  test("a chat id that is not a 1:1 id is an error", async () => {
    serve({ [THREADS]: { status: 200, json: { threadResource: { id: "19:x@thread.v2" } } } });
    await expect(createDirectChat(ME, ALEX)).rejects.toThrow("Teams did not return the new chat.");
    serve({ [THREADS]: { status: 200, json: { threadResource: { id: 5 } } } });
    await expect(createDirectChat(ME, ALEX)).rejects.toThrow("Teams did not return the new chat.");
  });

  test("a failed request is an error and sends nothing else", async () => {
    serve({ [THREADS]: { status: 403, json: { message: "Property [uniquerosterthread] is not allowed for creation." } } });
    await expect(createDirectChat(ME, ALEX)).rejects.toThrow("403");
    expect(fake!.requests).toHaveLength(1);
  });
});

describe("resolving a person with no chat", () => {
  test("a pending chat comes back with the id Teams will give it, and nothing is created", async () => {
    serve({ [THREADS]: created });
    expect(await resolveConversation("@Alex Doe", { allowNew: true })).toEqual({ id: CHAT, label: "Alex Doe", pending: { mri: ALEX } });
    expect(calls(THREADS)).toHaveLength(0);
  });

  test("without allowNew it still fails and never creates one", async () => {
    serve({ [THREADS]: created });
    await expect(resolveConversation("@Alex Doe")).rejects.toThrow("No 1:1 chat with Alex Doe exists in the cache.");
    expect(calls(THREADS)).toHaveLength(0);
  });

  test("a chat that exists is used, with allowNew too", async () => {
    serve({ [THREADS]: created, [CSA]: { status: 200, json: teamsMe } });
    expect(await resolveConversation("@Alex Doe", { allowNew: true })).toEqual({ id: "19:one@unq.gbl.spaces", label: "Alex Doe" });
  });
});

describe("send to a person with no chat", () => {
  test("a preview says a chat would be created, and creates nothing and sends nothing", async () => {
    serve({ [THREADS]: created, [MESSAGE]: sent });
    await send("@Alex Doe", ["hello"], plain);
    expect(logs[0]).toBe("Dry run (effect: message): would send to Alex Doe (new chat, created on send): hello");
    expect(calls(THREADS)).toHaveLength(0);
    expect(calls(MESSAGE)).toHaveLength(0);
    expect(existsSync(join(cacheDir(), TID, "ledger"))).toBe(false);
  });

  test("the JSON preview marks the new chat", async () => {
    serve({ [THREADS]: created, [MESSAGE]: sent });
    await send("@Alex Doe", ["hello"], { ...plain, json: true });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview).toMatchObject({ dry_run: true, new_chat: true, conversation: { id: CHAT, label: "Alex Doe" } });
    expect(calls(THREADS)).toHaveLength(0);
  });

  test("a real send creates the chat first and then sends to the returned chat", async () => {
    serve({ [THREADS]: created, [MESSAGE]: sent });
    await send("@Alex Doe", ["hello"], { ...plain, allowWrite: true });
    const order = fake!.requests.filter(item => item.method === "POST" && (item.path.endsWith("/threads") || item.path.endsWith("/messages"))).map(item => item.path.endsWith("/threads") ? "create" : "message");
    expect(order).toEqual(["create", "message"]);
    expect(calls(MESSAGE)[0]!.path).toContain(CHAT);
    expect((calls(MESSAGE)[0]!.body as any).conversationid).toBe(CHAT);
  });

  test("an existing chat is not created again", async () => {
    serve({ [THREADS]: created, [MESSAGE]: sent, [CSA]: { status: 200, json: teamsMe }, "POST /api/chatsvc/amer/v1/users/ME/conversations/19:one@unq.gbl.spaces/messages": sent });
    await send("@Alex Doe", ["hello"], { ...plain, allowWrite: true });
    expect(calls(THREADS)).toHaveLength(0);
  });

  test("when the chat cannot be created nothing is sent and no send record is left", async () => {
    serve({ [THREADS]: { status: 403, json: { message: "denied" } }, [MESSAGE]: sent });
    await expect(send("@Alex Doe", ["hello"], { ...plain, allowWrite: true })).rejects.toThrow();
    expect(calls(MESSAGE)).toHaveLength(0);
    expect(existsSync(join(cacheDir(), TID, "ledger"))).toBe(false);
  });

  test("a message that cannot be composed creates no chat", async () => {
    serve({ [THREADS]: created, [MESSAGE]: sent });
    await expect(send("@Alex Doe", ["hi @{Nobody Atall}"], { ...plain, allowWrite: true })).rejects.toThrow();
    expect(calls(THREADS)).toHaveLength(0);
    expect(calls(MESSAGE)).toHaveLength(0);
  });

  test("a different chat id in the reply stops the send", async () => {
    serve({ [THREADS]: { status: 200, json: { threadResource: { id: "19:99999999-9999-9999-9999-999999999999_aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa@unq.gbl.spaces" } } }, [MESSAGE]: sent });
    await expect(send("@Alex Doe", ["hello"], { ...plain, allowWrite: true })).rejects.toThrow("Teams returned a different chat than expected. Nothing was sent.");
    expect(calls(MESSAGE)).toHaveLength(0);
  });
});
