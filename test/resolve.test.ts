import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { resolveConversation, loadConversations, parseTeamsMe } from "../src/conversations.ts";
import { findUsers, pickUser, resolveMentionTokens, scoreUser, type User } from "../src/users.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";
import { TID } from "./jwt.ts";
import teamsMe from "./fixtures/teams-me.json";
import searchV2 from "./fixtures/search-v2.json";

const ME_PATH = "GET /api/csa/amer/api/v3/teams/users/me";
const SEARCH_PATH = "POST /api/mt/amer/beta/users/searchV2";
let fake: ReturnType<typeof installFakeTeams> | undefined;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
  return fake;
}

beforeEach(async () => {
  await resetLocalState();
});

afterEach(() => {
  fake?.uninstall();
  fake = undefined;
});

describe("conversation references", () => {
  test("an id resolves with no request", async () => {
    await seedProfile(["chat"]);
    serve({});
    expect(await resolveConversation("19:abc@thread.v2")).toEqual({ id: "19:abc@thread.v2", label: "19:abc@thread.v2" });
    expect(fake!.requests.length).toBe(0);
  });

  test("me resolves to notes to self with no request", async () => {
    await seedProfile(["chat"]);
    serve({});
    expect(await resolveConversation("me")).toEqual({ id: "48:notes", label: "Notes to self" });
    expect(fake!.requests.length).toBe(0);
  });

  test("Team/Channel refreshes the cache once and resolves", async () => {
    await seedProfile();
    serve({ [ME_PATH]: { status: 200, json: teamsMe } });
    expect(await resolveConversation("Engineering/Developers")).toEqual({ id: "19:dev@thread.tacv2", label: "Engineering/Developers" });
    expect(fake!.requests.length).toBe(1);
    expect(fake!.requests[0]!.query).toEqual({ isPrefetch: "false", enableMembershipSummary: "true" });
  });

  test("a cache hit makes no request", async () => {
    await seedProfile();
    serve({ [ME_PATH]: { status: 200, json: teamsMe } });
    await resolveConversation("Engineering/Developers");
    await resolveConversation("#General");
    expect(fake!.requests.length).toBe(1);
  });

  test("a bare channel name and a hash name resolve", async () => {
    await seedProfile();
    serve({ [ME_PATH]: { status: 200, json: teamsMe } });
    expect((await resolveConversation("#General")).id).toBe("19:general@thread.tacv2");
    expect((await resolveConversation("developers")).id).toBe("19:dev@thread.tacv2");
  });

  test("a chat title resolves when no channel has the name", async () => {
    await seedProfile();
    serve({ [ME_PATH]: { status: 200, json: teamsMe } });
    expect((await resolveConversation("Stand-Up")).id).toBe("19:chat1@thread.v2");
  });

  test("many candidates fail and list each one with its id", async () => {
    await seedProfile();
    const twoTeams = { teams: [
      { id: "19:a@thread.tacv2", displayName: "Alpha", channels: [{ id: "19:a-gen@thread.tacv2", displayName: "General" }] },
      { id: "19:b@thread.tacv2", displayName: "Beta", channels: [{ id: "19:b-gen@thread.tacv2", displayName: "General" }] },
    ], chats: [] };
    serve({ [ME_PATH]: { status: 200, json: twoTeams } });
    const error = (await resolveConversation("General").then(() => undefined, e => e as Error))!;
    expect(error.message).toBe('"General" matches more than one conversation: Alpha/General (19:a-gen@thread.tacv2), Beta/General (19:b-gen@thread.tacv2). Pass a conversation id.');
  });

  test("a miss with no csa token says how to fix it", async () => {
    await seedProfile(["chat"]);
    serve({});
    await expect(resolveConversation("Engineering/Developers")).rejects.toThrow("Name lookup needs a csa token. Run: teams-cli auth");
    expect(fake!.requests.length).toBe(0);
  });

  test("a name with no match after refresh fails", async () => {
    await seedProfile();
    serve({ [ME_PATH]: { status: 200, json: teamsMe } });
    await expect(resolveConversation("Nowhere")).rejects.toThrow('No conversation matches "Nowhere".');
  });

  test("@Name resolves the user and the cached 1:1 chat", async () => {
    await seedProfile();
    serve({
      [SEARCH_PATH]: { status: 200, json: searchV2 },
      [ME_PATH]: { status: 200, json: teamsMe },
    });
    expect(await resolveConversation("@Alex Doe")).toEqual({ id: "19:one@unq.gbl.spaces", label: "Alex Doe" });
  });

  test("an email resolves the same way", async () => {
    await seedProfile();
    serve({
      [SEARCH_PATH]: { status: 200, json: searchV2 },
      [ME_PATH]: { status: 200, json: teamsMe },
    });
    expect((await resolveConversation("adoe@example.test")).id).toBe("19:one@unq.gbl.spaces");
  });

  test("a user with no cached 1:1 chat fails and never creates one", async () => {
    await seedProfile();
    const noOneOnOne = { ...teamsMe, chats: [] };
    serve({
      [SEARCH_PATH]: { status: 200, json: searchV2 },
      [ME_PATH]: { status: 200, json: noOneOnOne },
    });
    await expect(resolveConversation("@Alex Doe")).rejects.toThrow("No 1:1 chat with Alex Doe exists in the cache. Send to Alex Doe with --allow-write to start one, or pass a chat id.");
    expect(fake!.requests.some(item => item.path.includes("/conversations"))).toBe(false);
  });

  test("the cache file is keyed by tenant", async () => {
    await seedProfile();
    serve({ [ME_PATH]: { status: 200, json: teamsMe } });
    await resolveConversation("#General");
    expect((await loadConversations(TID)).teams.length).toBe(1);
  });

  test("parseTeamsMe ignores missing and unknown fields", () => {
    const parsed = parseTeamsMe({ teams: [{ id: "19:t@thread.tacv2" }, { nope: true }], chats: [{ id: "19:c@thread.v2", extra: 1 }] });
    expect(parsed.teams).toEqual([{ id: "19:t@thread.tacv2", name: "", channels: [] }]);
    expect(parsed.chats).toEqual([{ id: "19:c@thread.v2", title: "", type: "", isOneOnOne: false, members: [] }]);
    expect(parseTeamsMe(undefined)).toEqual({ teams: [], chats: [] });
  });
});

describe("users", () => {
  test("searchV2 uses the query string, a raw body, and keeps only 8:orgid entries", async () => {
    await seedProfile();
    serve({ [SEARCH_PATH]: { status: 200, json: searchV2 } });
    const users = await findUsers("Alex");
    const sent = fake!.requests[0]!;
    expect(sent.query).toEqual({ includeDLs: "false", enableGuest: "false", includeBots: "true", skypeTeamsInfo: "true", source: "shareContact" });
    expect(sent.body).toBe("Alex");
    expect(sent.headers["content-type"]).toBe("text/plain;charset=UTF-8");
    expect(users.map(user => user.mri)).toEqual([
      "8:orgid:33333333-3333-3333-3333-333333333333",
      "8:orgid:44444444-4444-4444-4444-444444444444",
    ]);
  });

  test("a cache hit makes no request", async () => {
    await seedProfile();
    serve({ [SEARCH_PATH]: { status: 200, json: searchV2 } });
    await findUsers("Alex");
    await findUsers("Alex Doe");
    expect(fake!.requests.length).toBe(1);
  });

  test("a cache miss without a spaces token fails with the fix", async () => {
    await seedProfile(["chat", "csa"]);
    serve({});
    await expect(findUsers("Alex")).rejects.toThrow("User lookup needs a spaces token. Run: teams-cli auth");
    expect(fake!.requests.length).toBe(0);
  });

  test("the mention token becomes a placeholder and keeps order", async () => {
    await seedProfile();
    serve({ [SEARCH_PATH]: { status: 200, json: searchV2 } });
    const resolved = await resolveMentionTokens("hi @{Alex Doe}, and @{Alex Doe}");
    expect(resolved.mentions.map(user => user.displayName)).toEqual(["Alex Doe", "Alex Doe"]);
    expect(resolved.text).toBe("hi 0, and 1");
  });

  test("a mention with no matching user fails", async () => {
    await seedProfile();
    serve({ [SEARCH_PATH]: { status: 200, json: { value: [] } } });
    await expect(resolveMentionTokens("@{Nobody Here}")).rejects.toThrow('Could not find a Teams user matching "Nobody Here".');
  });

  test("an ambiguous partial name lists the candidates", async () => {
    await seedProfile();
    serve({ [SEARCH_PATH]: { status: 200, json: searchV2 } });
    await expect(resolveMentionTokens("@{Alex}")).rejects.toThrow(
      'Mention "Alex" is ambiguous: Alex Doe (8:orgid:33333333-3333-3333-3333-333333333333), Alex Doerr (8:orgid:44444444-4444-4444-4444-444444444444).',
    );
  });

  test("an exact name wins over partial matches", () => {
    const users: User[] = [
      { mri: "8:orgid:a", displayName: "Alex Doe", fetchedAt: "" },
      { mri: "8:orgid:b", displayName: "Alex Doerr", fetchedAt: "" },
    ];
    expect(pickUser(users, "Alex Doe").mri).toBe("8:orgid:a");
  });

  test("scoring is exact, then prefix, then substring", () => {
    const user: User = { mri: "8:orgid:a", displayName: "Alex Doe", mail: "adoe@example.test", fetchedAt: "" };
    expect(scoreUser(user, "alex doe")).toBe(0);
    expect(scoreUser(user, "alex")).toBe(1);
    expect(scoreUser(user, "doe@")).toBe(2);
    expect(scoreUser(user, "zed")).toBe(3);
  });
});
