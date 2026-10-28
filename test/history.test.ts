import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { history } from "../src/commands/history.ts";
import { send } from "../src/commands/send.ts";
import { htmlToText, normalizeMessage } from "../src/message.ts";
import { resolveConversation } from "../src/conversations.ts";
import { resolveNames } from "../src/users.ts";
import { TID } from "./jwt.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Reply, type Route } from "./fake-teams.ts";

const CONV = "19:dev@thread.tacv2";
const BASE = "/api/chatsvc/amer/v1/users/ME/conversations";
const CONV_PATH = `GET ${BASE}/${CONV}/messages`;
const NOTES_PATH = `GET ${BASE}/48:notes/messages`;
const PROFILES = "POST /api/mt/amer/beta/users/fetchShortProfile";
const ME = "8:orgid:11111111-1111-1111-1111-111111111111";
const ALEX = "8:orgid:22222222-2222-2222-2222-222222222222";
const SAM = "8:orgid:33333333-3333-3333-3333-333333333333";
const UNKNOWN = "8:orgid:99999999-9999-9999-9999-999999999999";
const PROFILE_REPLY = { status: 200, json: [{ mri: ALEX, displayName: "Alex Doe" }, { mri: SAM, displayName: "Sam Roe" }] };

let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
let errors: string[] = [];
const originalLog = console.log;
const originalError = console.error;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}

const from = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;
const raw = (id: number, extra: Record<string, unknown> = {}) => ({
  id: String(id), messagetype: "RichText/Html", content: `<p>message ${id}</p>`,
  imdisplayname: "Alex Doe", from: from(ALEX), properties: {}, ...extra,
});
const page = (ids: number[], link?: string) => ({
  status: 200,
  json: { messages: ids.map(id => raw(id)), ...(link ? { _metadata: { backwardLink: link } } : {}) },
});
const linkTo = (n: number) => `https://teams.cloud.microsoft${BASE}/${CONV}/messages?pageSize=200&startTime=1&before=${n}`;
const context = { conversationId: CONV, myMri: ME, nameFor: (mri: string) => ({ [ALEX]: "Alex Doe", [SAM]: "Sam Roe" } as Record<string, string>)[mri] };

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

/** Three pages of 200, newest first, with ids 1791000000001 to 1791000000600. */
function threePages(): Reply[] {
  const ids = Array.from({ length: 600 }, (_, index) => 1791000000001 + index);
  const chunk = (start: number) => ids.slice(start, start + 200);
  return [
    page(chunk(400), linkTo(1791000000400)),
    page(chunk(200), linkTo(1791000000200)),
    page(chunk(0)),
  ];
}

describe("paging", () => {
  test("--top follows backward links only until it has enough messages", async () => {
    serve({ [CONV_PATH]: threePages() });
    await history(CONV, { json: true, content: false, top: 250 });
    const out = JSON.parse(logs.join("\n"));
    expect(out.messages).toHaveLength(250);
    expect(out.messages[249].ts).toBe("1791000000600");
    expect(fake!.requests).toHaveLength(2);
  });

  test("--all reads every page to the start", async () => {
    serve({ [CONV_PATH]: threePages() });
    await history(CONV, { json: true, content: false, all: true });
    expect(JSON.parse(logs.join("\n")).messages).toHaveLength(600);
    expect(fake!.requests).toHaveLength(3);
    expect(errors).toEqual([]);
  });

  test("--after-ts pages until the oldest message is at or before the time", async () => {
    serve({ [CONV_PATH]: threePages() });
    await history(CONV, { json: true, content: false, afterTs: "1791000000350" });
    const out = JSON.parse(logs.join("\n"));
    expect(out.messages[0].ts).toBe("1791000000351");
    expect(out.messages).toHaveLength(250);
    expect(fake!.requests).toHaveLength(2);
  });

  test("a paging link off the Teams host is refused", async () => {
    serve({ [CONV_PATH]: page([1791000000001], "https://evil.example.test/api/chatsvc/amer/v1/x") });
    await expect(history(CONV, { json: false, content: false, all: true })).rejects.toThrow("Refused a paging link");
    expect(fake!.requests).toHaveLength(1);
  });

  test("--all stops at 20,000 messages and says so", async () => {
    const pages: Reply[] = Array.from({ length: 101 }, (_, index) => {
      const start = 1791000000000 + index * 200;
      return page(Array.from({ length: 200 }, (_, offset) => start + offset), linkTo(start));
    });
    serve({ [CONV_PATH]: pages });
    await history(CONV, { json: true, content: false, all: true });
    expect(JSON.parse(logs.join("\n"))).toBeDefined();
    expect(fake!.requests).toHaveLength(100);
    expect(errors).toEqual(["Stopped after 20,000 messages. Older messages were not read. Use --after-ts to read a narrower range."]);
  });

  test("--all with --top is refused before any request", async () => {
    serve({ [CONV_PATH]: threePages() });
    await expect(history(CONV, { json: false, content: false, all: true, top: 5 })).rejects.toThrow("Use --all alone");
    expect(fake!.requests).toHaveLength(0);
  });
});

describe("reactions", () => {
  const withReactions = raw(1791000000100, {
    properties: {
      emotions: JSON.stringify([
        { key: "like", users: [{ mri: ALEX, time: 1, value: "1" }, { mri: ME, time: 2, value: "2" }] },
        { key: "custom_key", users: [{ mri: SAM, time: 3, value: "3" }] },
      ]),
    },
  });

  test("reactions show emoji, names from fetchShortProfile, and mine", async () => {
    serve({ [CONV_PATH]: { status: 200, json: { messages: [withReactions] } }, [PROFILES]: PROFILE_REPLY });
    await history(CONV, { json: true, content: false });
    const message = JSON.parse(logs.join("\n")).messages[0];
    expect(message.reactions).toEqual([
      { key: "like", emoji: "👍", count: 2, users: ["Alex Doe", ME], mine: true },
      { key: "custom_key", count: 1, users: ["Sam Roe"], mine: false },
    ]);
  });

  test("the human line lists each reaction with its names", async () => {
    serve({ [CONV_PATH]: { status: 200, json: { messages: [withReactions] } }, [PROFILES]: PROFILE_REPLY });
    await history(CONV, { json: false, content: false });
    expect(logs).toContain(`  reactions: 👍 2 (Alex Doe, ${ME}), custom_key 1 (Sam Roe)`);
  });

  test("the human line names a message's files and the command that saves them", async () => {
    const withFile = raw(1791000000200, {
      content: "",
      properties: { files: JSON.stringify([{ fileName: "plan.pdf", fileInfo: { fileUrl: "https://contoso.sharepoint.com/sites/s/Shared%20Documents/plan.pdf" } }]) },
    });
    serve({ [CONV_PATH]: { status: 200, json: { messages: [withFile] } }, [PROFILES]: PROFILE_REPLY });
    await history(CONV, { json: false, content: false });
    expect(logs).toContain(`  files: plan.pdf (files download ${CONV}#1791000000200)`);
  });

  test("a failed name lookup falls back to the user id", async () => {
    serve({ [CONV_PATH]: { status: 200, json: { messages: [withReactions] } }, [PROFILES]: { status: 500, json: { message: "down" } } });
    await history(CONV, { json: true, content: false });
    expect(JSON.parse(logs.join("\n")).messages[0].reactions[0].users).toEqual([ALEX, ME]);
  });

  test("a second lookup of the same ids uses the cache and sends no request", async () => {
    serve({ [PROFILES]: PROFILE_REPLY });
    const first = await resolveNames(TID, [ALEX, SAM]);
    expect(first).toEqual({ [ALEX]: "Alex Doe", [SAM]: "Sam Roe" });
    const second = await resolveNames(TID, [ALEX, SAM]);
    expect(second).toEqual(first);
    expect(fake!.requests.filter(item => item.path.endsWith("fetchShortProfile"))).toHaveLength(1);
  });
});

describe("message fields", () => {
  test("type, reply_to, pinned and edited", () => {
    const base = { id: "1791000000200", messagetype: "RichText/Html", content: "<p>x</p>", from: from(ALEX), imdisplayname: "Alex Doe" };
    const reply = normalizeMessage({ ...base, rootMessageId: "1791000000100" }, context)!;
    expect(reply).toMatchObject({ type: "message", reply_to: "1791000000100", pinned: false });
    expect(normalizeMessage({ ...base, rootMessageId: "1791000000200" }, context)!.reply_to).toBeUndefined();
    expect(normalizeMessage({ ...base }, context)!.reply_to).toBeUndefined();
    expect(normalizeMessage({ ...base, properties: { pinned: { creatorId: ALEX, pinnedTime: 5 } } }, context)!.pinned).toBe(true);
    expect(normalizeMessage({ ...base, properties: { pinned: { creatorId: "", pinnedTime: 0 } } }, context)!.pinned).toBe(false);
    expect(normalizeMessage({ ...base, properties: { edittime: 1791000000300 } }, context)!.edited).toBe(new Date(1791000000300).toISOString());
  });

  test("a missing sender shows as system", () => {
    const message = normalizeMessage({ id: "1791000000200", messagetype: "Text", content: "hi" }, context)!;
    expect(message.user).toBe("system");
  });

  test("quotes and forwards render as text", () => {
    expect(htmlToText("<blockquote itemtype=\"http://schema.skype.com/Reply\"><strong>Pat Lee</strong><p>first</p></blockquote><p>reply</p>")).toBe("> Pat Lee: first\nreply");
    expect(htmlToText("<div itemtype=\"http://schema.skype.com/Forward\"><p>forward body</p></div><p>mine</p>")).toBe("[forwarded] forward body\nmine");
  });
});

describe("--system", () => {
  const system = { ...context, system: true };
  const callXml = "<ended/><partlist alt=\"x\" count=\"2\"><part identity=\"8:orgid:x\"><name>8:orgid:x</name><displayName>Alex Doe</displayName><duration>1747</duration></part><part identity=\"8:orgid:y\"><name>8:orgid:y</name><displayName>Sam Roe</displayName><duration>60</duration></part></partlist>";

  test("a call lists each participant with seconds", () => {
    const message = normalizeMessage(raw(1791000000300, { messagetype: "Event/Call", content: callXml }), system)!;
    expect(message.type).toBe("call");
    expect(message.call).toEqual({ participants: [{ name: "Alex Doe", seconds: 1747 }, { name: "Sam Roe", seconds: 60 }] });
    expect(message.text).toBe("Call with Alex Doe, Sam Roe");
  });

  test("a recording gives title, url and organizer", () => {
    const xml = "<URIObject><RecordingStatus status=\"x\"/><Title>Stand-Up recording</Title><a href=\"https://example.test/rec?x=1&amp;y=2\">Stand-Up</a><MeetingOrganizerId value=\"22222222-2222-2222-2222-222222222222\" /></URIObject>";
    const message = normalizeMessage(raw(1791000000301, { messagetype: "RichText/Media_CallRecording", content: xml }), system)!;
    expect(message.type).toBe("recording");
    expect(message.recording).toEqual({ title: "Stand-Up recording", url: "https://example.test/rec?x=1&y=2", organizer: "Alex Doe" });
  });

  test("a transcript exposes the call and storage ids from a JSON string", () => {
    const content = "{\\\"callId\\\":\\\"call-1\\\",\\\"storageId\\\":\\\"store-1\\\",\\\"isDeleted\\\":false}";
    const message = normalizeMessage(raw(1791000000302, { messagetype: "RichText/Media_CallTranscript", content }), system)!;
    expect(message.transcript).toEqual({ callId: "call-1", storageId: "store-1" });
  });

  test("activity events get a system type and a readable summary", () => {
    const joined = normalizeMessage(raw(1791000000303, {
      messagetype: "ThreadActivity/MemberJoined",
      content: JSON.stringify({ eventtime: 1, initiator: ALEX, members: [{ id: "8:teamsvisitor:z", friendlyname: "Pat Lee" }] }),
    }), system)!;
    expect(joined).toMatchObject({ type: "system:memberjoined", text: "Pat Lee joined" });
    const added = normalizeMessage(raw(1791000000304, {
      messagetype: "ThreadActivity/AddMember",
      content: `<addmember><eventtime>1</eventtime><initiator>${ALEX}</initiator><target>${SAM}</target></addmember>`,
    }), { ...system, nameFor: mri => ({ [ALEX]: "Alex Doe", [SAM]: "Sam Roe" } as Record<string, string>)[mri] })!;
    expect(added).toMatchObject({ type: "system:addmember", text: "Alex Doe added Sam Roe" });
    const unresolved = normalizeMessage(raw(1791000000305, {
      messagetype: "ThreadActivity/AddMember",
      content: `<addmember><initiator>${ALEX}</initiator><target>${UNKNOWN}</target></addmember>`,
    }), system)!;
    expect(unresolved.text).toBe("member added");
    const renamed = normalizeMessage(raw(1791000000306, {
      messagetype: "ThreadActivity/TopicTopicUpdated", content: JSON.stringify({ oldValue: "a", newValue: "b", user: ALEX }),
    }), system)!;
    expect(renamed).toMatchObject({ type: "system:topictopicupdated", text: "topic changed to \"b\"" });
  });

  test("a deleted message is kept only with --system", () => {
    const gone = raw(1791000000307, { content: "", properties: { deletetime: 1791000000400 } });
    expect(normalizeMessage(gone, context)).toBeUndefined();
    expect(normalizeMessage(gone, system)).toMatchObject({ type: "deleted", deleted: true });
  });

  test("without --system the events and recordings are dropped", () => {
    expect(normalizeMessage(raw(1791000000300, { messagetype: "Event/Call", content: callXml }), context)).toBeUndefined();
    expect(normalizeMessage(raw(1791000000303, { messagetype: "ThreadActivity/MemberJoined", content: "{}" }), context)).toBeUndefined();
  });

  test("history --system shows the call in the human output", async () => {
    serve({ [CONV_PATH]: { status: 200, json: { messages: [raw(1791000000300, { messagetype: "Event/Call", content: callXml })] } }, [PROFILES]: PROFILE_REPLY });
    await history(CONV, { json: false, content: false, system: true });
    expect(logs).toContain("  Call with Alex Doe, Sam Roe");
  });
});

describe("streams", () => {
  test("me in any case reads the notes chat", async () => {
    expect(await resolveConversation("ME")).toEqual({ id: "48:notes", label: "Notes to self" });
    expect(await resolveConversation("Me")).toEqual({ id: "48:notes", label: "Notes to self" });
  });

  test("mentions and notifications are aliases for their streams", async () => {
    expect(await resolveConversation("mentions")).toEqual({ id: "48:mentions", label: "Mentions" });
    expect(await resolveConversation("notifications")).toEqual({ id: "48:notifications", label: "Notifications" });
  });

  test("any 48:* id reads, and a missing sender shows as system", async () => {
    serve({
      [`GET ${BASE}/48:mentions/messages`]: { status: 200, json: { messages: [{ id: "1791000000400", messagetype: "RichText/Html", content: "<p>@you</p>", properties: {} }] } },
    });
    await history("mentions", { json: true, content: false });
    const out = JSON.parse(logs.join("\n"));
    expect(out.channel_id).toBe("48:mentions");
    expect(out.messages[0].user).toBe("system");
    expect(fake!.requests[0]!.path).toBe(`${BASE}/48:mentions/messages`);
  });

  test("a send to mentions or notifications is refused before any request", async () => {
    serve({ [NOTES_PATH]: { status: 200, json: {} } });
    await expect(send("mentions", ["hi"], { json: false, verify: false, allowWrite: true })).rejects.toThrow("That stream is read-only.");
    await expect(send("notifications", ["hi"], { json: false, verify: false, allowWrite: true })).rejects.toThrow("That stream is read-only.");
    expect(fake!.requests).toHaveLength(0);
  });
});
