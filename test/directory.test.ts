import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { mergeChatService, type ConversationCache } from "../src/conversations.ts";
import { conversations } from "../src/commands/conversations.ts";
import { membersOf } from "../src/commands/members.ts";
import { searchPeople } from "../src/users.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Reply, type Route } from "./fake-teams.ts";
import searchV2 from "./fixtures/search-v2.json";
import teamsMe from "./fixtures/teams-me.json";

const SEARCH = "POST /api/mt/amer/beta/users/searchV2";
const PROFILES = "POST /api/mt/amer/beta/users/fetchShortProfile";
const CSA_ME = "GET /api/csa/amer/api/v3/teams/users/me";
const CHAT_LIST = "GET /api/chatsvc/amer/v1/users/ME/conversations";
const THREAD = "GET /api/chatsvc/amer/v1/threads/19:dev@thread.tacv2";
const ALEX = "8:orgid:33333333-3333-3333-3333-333333333333";
const SAM = "8:orgid:55555555-5555-5555-5555-555555555555";
const found: Reply = { status: 200, json: searchV2 };
const nothing: Reply = { status: 200, json: { value: [] } };
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
let errors: string[] = [];
const originalLog = console.log;
const originalError = console.error;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}

beforeEach(async () => {
  await resetLocalState();
  await seedProfile();
  logs = [];
  errors = [];
  console.log = (...args: unknown[]) => { logs.push(args.map(String).join(" ")); };
  console.error = (...args: unknown[]) => { errors.push(args.map(String).join(" ")); };
});

afterEach(() => {
  console.log = originalLog;
  console.error = originalError;
  fake?.uninstall();
  fake = undefined;
  process.exitCode = 0;
});

describe("users find", () => {
  test("an empty or whitespace query is refused before any request", async () => {
    serve({ [SEARCH]: found });
    await expect(searchPeople("   ")).rejects.toThrow("Give a name or email to search for.");
    expect(fake!.requests).toHaveLength(0);
  });

  test("repeated spaces collapse before the directory is asked", async () => {
    serve({ [SEARCH]: found });
    await searchPeople("alex    doe");
    expect(fake!.requests[0]!.body).toBe("alex doe");
  });

  test("a query of two or more characters asks the directory every time, and dedupes the merge", async () => {
    serve({ [SEARCH]: [found, found] });
    const first = await searchPeople("alex");
    const second = await searchPeople("alex");
    expect(fake!.requests.filter(item => item.path.endsWith("/searchV2"))).toHaveLength(2);
    const mris = second.map(user => user.mri);
    expect(new Set(mris).size).toBe(mris.length);
    expect(first.map(user => user.mri)).toEqual(mris);
  });

  test("cache hits come first, and a cached user is kept when the directory no longer returns them", async () => {
    serve({ [SEARCH]: [found, nothing] });
    await searchPeople("alex");
    const again = await searchPeople("alex");
    expect(again.map(user => user.displayName)).toEqual(["Alex Doe", "Alex Doerr"]);
  });

  test("a one-character query uses the cache only and says so", async () => {
    serve({ [SEARCH]: found });
    await searchPeople("a");
    expect(fake!.requests).toHaveLength(0);
    expect(errors).toEqual(["Searched the cache only. A one-character query is not sent to the directory."]);
  });

  test("--refresh skips the cache, so a cached user the directory does not return is dropped", async () => {
    serve({ [SEARCH]: [found, nothing] });
    await searchPeople("alex");
    const refreshed = await searchPeople("alex", { refresh: true });
    expect(refreshed).toEqual([]);
    await expect(searchPeople("a", { refresh: true })).rejects.toThrow("--refresh needs a query of 2 or more characters.");
  });

  test("without a spaces token it shows the cache and says so", async () => {
    serve({ [SEARCH]: found });
    await searchPeople("alex");
    await seedProfile(["chat", "csa"]);
    const cached = await searchPeople("alex");
    expect(cached.length).toBeGreaterThan(0);
    expect(errors).toContain("No spaces token. Showing cached users only. Run: teams-cli auth login");
  });
});

describe("chat-service merge", () => {
  const cache: ConversationCache = {
    version: 1, updatedAt: "", teams: [{ id: "team", name: "Eng", channels: [{ id: "19:ch@thread.tacv2", name: "General" }] }],
    chats: [{ id: "19:known@thread.v2", title: "Known", type: "meeting", isOneOnOne: false, members: [] }],
  };

  test("adds conversations the csa list lacks, and skips streams, known ids, teams, and empty threads", () => {
    const extra = mergeChatService(cache, [
      { id: "19:known@thread.v2", threadProperties: { threadType: "meeting" } },
      { id: "19:ch@thread.tacv2", threadProperties: { threadType: "topic" } },
      { id: "48:notes", threadProperties: { threadType: "streamofnotes" } },
      { id: "19:gone@thread.skype", threadProperties: { threadType: "space" } },
      { id: "19:sp@thread.skype", threadProperties: { threadType: "space" }, lastMessage: {} },
      { id: "19:new@thread.v2", threadProperties: { threadType: "meeting", topic: "Planning" } },
      { id: "19:empty@thread.v2", threadProperties: { threadType: "meeting" } },
      { id: "19:dm@thread.v2", threadProperties: { threadType: "chat", topic: "Pair" }, lastMessage: {} },
    ]);
    expect(extra).toEqual([
      { id: "19:new@thread.v2", title: "Planning", type: "meeting", isOneOnOne: false, members: [], source: "chatsvc" },
      { id: "19:dm@thread.v2", title: "Pair", type: "chat", isOneOnOne: false, members: [], source: "chatsvc" },
    ]);
  });

  test("a meeting with no title but a last message is merged", () => {
    const extra = mergeChatService(cache, [{ id: "19:m@thread.v2", threadProperties: { threadType: "meeting" }, lastMessage: {} }]);
    expect(extra).toEqual([{ id: "19:m@thread.v2", title: "", type: "meeting", isOneOnOne: false, members: [], source: "chatsvc" }]);
  });
});

describe("conversations", () => {
  const chatList = {
    status: 200,
    json: {
      conversations: [
        { id: "48:notes", threadProperties: { threadType: "streamofnotes" }, lastMessage: {} },
        { id: "19:only@thread.v2", threadProperties: { threadType: "meeting", topic: "Only Here" }, lastMessage: {} },
      ],
    },
  };

  test("lists streams, names small chats, counts large ones, and marks merged chats", async () => {
    const standUpMember = teamsMe.chats.find(chat => chat.title === "Stand-Up")!.members[0]!.mri;
    serve({ [CSA_ME]: { status: 200, json: teamsMe }, [CHAT_LIST]: chatList, [PROFILES]: { status: 200, json: [{ mri: standUpMember, displayName: "Alex Doe" }] } });
    await conversations(true);
    const out = JSON.parse(logs.join("\n"));
    expect(out.streams.map((stream: { id: string }) => stream.id)).toEqual(["48:notes", "48:mentions", "48:notifications"]);
    const merged = out.chats.find((chat: { id: string }) => chat.id === "19:only@thread.v2");
    expect(merged).toMatchObject({ source: "chatsvc", member_count: null, title: "Only Here" });
    const stand = out.chats.find((chat: { title: string }) => chat.title === "Stand-Up");
    expect(stand.member_count).toBe(1);
    expect(stand.members).toHaveLength(1);
    expect(stand.members[0]).toBe("Alex Doe");
  });

  test("without a csa token the chat list is used alone, with a note", async () => {
    await seedProfile(["chat"]);
    serve({ [CHAT_LIST]: chatList });
    await conversations(true);
    const out = JSON.parse(logs.join("\n"));
    expect(out.teams).toEqual([]);
    expect(out.chats.map((chat: { id: string }) => chat.id)).toEqual(["19:only@thread.v2"]);
    expect(errors[0]).toBe("No csa token. Listing chat-service conversations only. Run: teams-cli auth login");
  });
});

describe("members", () => {
  test("lists each member with name, mail, and role", async () => {
    serve({
      [THREAD]: { status: 200, json: { members: [{ id: ALEX, role: "Admin" }, { id: SAM, role: "User" }] } },
      [PROFILES]: { status: 200, json: { type: "x", value: [
        { mri: ALEX, displayName: "Alex Doe", email: "adoe@example.test" },
        { mri: SAM, displayName: "Sam Roe", email: "sroe@example.test" },
      ] } },
    });
    await membersOf("19:dev@thread.tacv2", true);
    expect(JSON.parse(logs.join("\n"))).toEqual([
      { mri: ALEX, name: "Alex Doe", email: "adoe@example.test", role: "Admin" },
      { mri: SAM, name: "Sam Roe", email: "sroe@example.test", role: "User" },
    ]);
  });

  test("a thread that lists one member prints that one", async () => {
    serve({ [THREAD]: { status: 200, json: { members: [{ id: ALEX, role: "Admin" }] } }, [PROFILES]: { status: 200, json: [] } });
    await membersOf("19:dev@thread.tacv2", false);
    expect(logs).toEqual([`${ALEX}\t${ALEX}\t\tAdmin`]);
  });
});
