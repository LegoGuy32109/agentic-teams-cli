import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { editMessage } from "../src/commands/edit.ts";
import { send } from "../src/commands/send.ts";
import { kindOf } from "../src/conversations.ts";
import { resolveMentionTokens } from "../src/users.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";

const API = "/api/chatsvc/amer/v1";
const CSA = "GET /api/csa/amer/api/v3/teams/users/me";
const SEARCH = "POST /api/mt/amer/beta/users/searchV2";
const GROUP = "19:fakegroup@thread.v2";
const MEETING = "19:meeting_fake@thread.v2";
const CHANNEL = "19:fakechannel@thread.tacv2";
const ONE = "19:aaa_bbb@unq.gbl.spaces";
const plain = { json: true, verify: false, allowWrite: false } as const;

let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
const originalLog = console.log;
const originalError = console.error;

const csa = { status: 200, json: { teams: [{ id: "19:team@thread.tacv2", displayName: "Fake Team", channels: [{ id: CHANNEL, displayName: "Renamed Channel" }] }], chats: [{ id: ONE, title: "", threadType: "chat", isOneOnOne: true, members: [] }] } };
function serve(routes: Record<string, Route> = {}) {
  fake?.uninstall();
  fake = installFakeTeams({ [CSA]: csa, ...routes });
}

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
});

const preview = () => JSON.parse(logs.join("\n"));
const span = (name: string) => `<readonly class="skipProofing" itemtype="http://schema.skype.com/Mention" spellcheck="false"><span itemtype="http://schema.skype.com/Mention" itemscope itemid="0">${name}</span></readonly>`;

describe("@{everyone}", () => {
  test("in a group chat and a meeting chat it is the everyone mention of the captured shape", async () => {
    for (const id of [GROUP, MEETING]) {
      serve();
      logs = [];
      await send(id, ["hi @{everyone}"], plain);
      const body = preview().body;
      expect(body.content).toBe(`<p>hi ${span("Everyone")}</p>`);
      expect(body.properties.mentions).toBe(JSON.stringify([{ "@type": "http://schema.skype.com/Mention", itemid: "0", mri: id, mentionType: "everyone", displayName: "Everyone" }]));
      expect(fake!.requests).toHaveLength(0);
    }
  });

  test("any case works, and rich text and --html get the same mention", async () => {
    serve();
    await send(GROUP, ["*hi* @{Everyone}"], { ...plain, format: "rich" });
    expect(preview().body.content).toBe(`<p><strong>hi</strong> ${span("Everyone")}</p>`);
    logs = [];
    await send(GROUP, [], { ...plain, html: "<p>hi @{EVERYONE}</p><ul><li>@{everyone}</li></ul>" });
    const body = preview().body;
    expect(body.content).toContain(span("Everyone"));
    expect(JSON.parse(body.properties.mentions).map((entry: any) => [entry.itemid, entry.mentionType, entry.mri])).toEqual([["0", "everyone", GROUP], ["1", "everyone", GROUP]]);
  });

  test("it is refused in a 1:1 chat, a chat that does not exist yet, Notes to self, and a channel", async () => {
    serve();
    await expect(send(ONE, ["@{everyone}"], plain)).rejects.toThrow("@{everyone} works only in a group or meeting chat. This is a 1:1 chat.");
    await expect(send("me", ["@{everyone}"], plain)).rejects.toThrow("This is Notes to self.");
    await expect(send(CHANNEL, ["@{everyone}"], plain)).rejects.toThrow("This is a channel.");
    expect(fake!.requests.filter(item => item.method !== "GET")).toHaveLength(0);
  });

  test("a 1:1 chat is found from the list too, when its id is not a pair id", async () => {
    serve();
    const cache = { version: 1 as const, updatedAt: "", teams: [], chats: [{ id: "19:odd@thread.v2", title: "", type: "chat", isOneOnOne: true, members: [] }] };
    expect(kindOf(cache, "19:odd@thread.v2")).toBe("oneOnOne");
    expect(kindOf(cache, "19:other@thread.v2")).toBe("group");
    expect(kindOf(cache, "19:x@thread.tacv2")).toBe("channel");
    expect(kindOf(cache, "48:notes")).toBe("notes");
    expect(kindOf(cache, "19:new_pair@unq.gbl.spaces", true)).toBe("oneOnOne");
  });

  test("an edit checks it the same way", async () => {
    serve({ [`GET ${API}/users/ME/conversations/${ONE}/messages/1791990424088`]: { status: 200, json: { from: "8:orgid:11111111-1111-1111-1111-111111111111", content: "x" } } });
    await expect(editMessage(ONE, "1791990424088", ["@{everyone}"], { json: true, allowWrite: false })).rejects.toThrow("@{everyone} works only in a group or meeting chat");
  });
});

describe("@{channel}", () => {
  test("in a channel it is the channel mention with the current channel name", async () => {
    serve();
    await send(CHANNEL, ["hi @{channel}"], plain);
    const body = preview().body;
    expect(body.content).toBe(`<p>hi ${span("Renamed Channel")}</p>`);
    expect(body.properties.mentions).toBe(JSON.stringify([{ "@type": "http://schema.skype.com/Mention", itemid: "0", mri: CHANNEL, mentionType: "channel", displayName: "Renamed Channel" }]));
  });

  test("it works in rich text and --html, and in a thread reply", async () => {
    serve();
    await send(CHANNEL, ["> @{Channel}"], { ...plain, format: "rich" });
    expect(preview().body.content).toBe(`<blockquote><p>${span("Renamed Channel")}</p></blockquote>`);
    logs = [];
    await send(CHANNEL, [], { ...plain, html: "<p>@{channel}</p>", threadTs: "1791990424088" });
    expect(JSON.parse(preview().body.properties.mentions)[0]).toMatchObject({ mentionType: "channel", displayName: "Renamed Channel" });
  });

  test("it is refused in a group chat, a 1:1 chat, and Notes to self, and a channel that is not in the list is an error", async () => {
    serve();
    await expect(send(GROUP, ["@{channel}"], plain)).rejects.toThrow("@{channel} works only in a channel. This is a group chat.");
    await expect(send(ONE, ["@{channel}"], plain)).rejects.toThrow("This is a 1:1 chat.");
    await expect(send("me", ["@{channel}"], plain)).rejects.toThrow("This is Notes to self.");
    await expect(send("19:unknown@thread.tacv2", ["@{channel}"], plain)).rejects.toThrow("is not in your team list");
  });
});

describe("a person with the name Everyone or Channel", () => {
  const people = (name: string, mail: string) => ({ status: 200, json: { value: [{ mri: "8:orgid:0000000b-0000-0000-0000-00000000000b", displayName: name, email: mail, userPrincipalName: mail }] } });

  test("is reached by email, and a longer name is still a person lookup", async () => {
    serve({ [SEARCH]: people("Everyone", "every.one@example.test") });
    await send(GROUP, ["@{every.one@example.test}"], plain);
    expect(JSON.parse(preview().body.properties.mentions)[0]).toMatchObject({ mentionType: "person", mri: "8:orgid:0000000b-0000-0000-0000-00000000000b", displayName: "Everyone" });
    serve({ [SEARCH]: people("Everyone Smith", "es@example.test") });
    logs = [];
    await send(GROUP, ["@{Everyone Smith}"], plain);
    expect(JSON.parse(preview().body.properties.mentions)[0]).toMatchObject({ mentionType: "person", displayName: "Everyone Smith" });
  });

  test("the word alone is never a person lookup, and no directory call is made for it", async () => {
    serve({ [SEARCH]: people("Everyone", "every.one@example.test") });
    await resolveMentionTokens("@{everyone}", { conversationId: GROUP, kind: "group" });
    expect(fake!.requests).toHaveLength(0);
    await expect(resolveMentionTokens("@{everyone}")).rejects.toThrow("needs a conversation");
  });
});
