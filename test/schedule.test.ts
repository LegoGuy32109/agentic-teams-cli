import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { send } from "../src/commands/send.ts";
import { listScheduled, scheduled, scheduledCancel } from "../src/commands/scheduled.ts";
import { parseSendAt } from "../src/schedule.ts";
import { ledgerKey } from "../src/ledger.ts";
import { cacheDir } from "../src/config.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";
import { TID } from "./jwt.ts";
import searchV2 from "./fixtures/search-v2.json";

const DRAFTS = "/api/chatsvc/amer/v1/users/ME/drafts";
const CHANNEL = "19:fakechannel0000000000000000000000@thread.tacv2";
const CREATE = `POST ${DRAFTS}`;
const LIST = `GET ${DRAFTS}`;
const CANCEL = `DELETE ${DRAFTS}/{id}`;
const SEARCH = "POST /api/mt/amer/beta/users/searchV2";
const ME_MRI = "8:orgid:11111111-1111-1111-1111-111111111111";
const CREATED = { status: 201, json: { OriginalArrivalTime: 1792421450913 } };
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
const originalLog = console.log;
const originalError = console.error;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}

const calls = (method: string) => fake!.requests.filter(item => item.method === method);
const ledgerDir = () => join(cacheDir(), TID, "ledger");

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

// A clock in the local zone, so the tests give the same result in any TZ.
const local = (hours: number, minutes: number, seconds = 0) => new Date(2026, 9, 19, hours, minutes, seconds, 0);
const plain = { json: true, verify: false, allowWrite: false } as const;

describe("parseSendAt", () => {
  const now = local(10, 0);

  test("reads +<n>m and +<n>h", () => {
    expect(parseSendAt("+10m", now).getTime()).toBe(now.getTime() + 10 * 60_000);
    expect(parseSendAt("+2h", now).getTime()).toBe(now.getTime() + 2 * 3_600_000);
  });

  test("reads HH:MM as today in the local zone", () => {
    expect(parseSendAt("11:30", now).getTime()).toBe(local(11, 30).getTime());
    expect(parseSendAt("9:05", local(8, 0)).getTime()).toBe(local(9, 5).getTime());
  });

  test("reads HH:MM as tomorrow when it has passed", () => {
    expect(parseSendAt("09:00", now).getTime()).toBe(new Date(2026, 9, 20, 9, 0).getTime());
  });

  test("reads ISO 8601 with Z or an offset", () => {
    expect(parseSendAt("2026-10-20T15:30:00Z", now).toISOString()).toBe("2026-10-20T15:30:00.000Z");
    expect(parseSendAt("2026-10-20T11:30:00-04:00", now).toISOString()).toBe("2026-10-20T15:30:00.000Z");
  });

  test("refuses ISO 8601 without an offset", () => {
    expect(() => parseSendAt("2026-10-20T11:30:00", now)).toThrow("not a valid --at time");
  });

  test("refuses a time in the past", () => {
    expect(() => parseSendAt("2026-10-19T08:00:00Z", new Date("2026-10-19T12:00:00Z"))).toThrow("in the past");
  });

  test("refuses a time less than a minute ahead", () => {
    expect(() => parseSendAt("+0m", now)).toThrow("in the past");
    expect(() => parseSendAt("2026-10-19T14:00:30Z", new Date("2026-10-19T14:00:00Z"))).toThrow("at least 1 minute");
    expect(parseSendAt("2026-10-19T14:01:00Z", new Date("2026-10-19T14:00:00Z")).toISOString()).toBe("2026-10-19T14:01:00.000Z");
  });

  test("a clock time just after midnight rolls to the next day and is checked against the lead", () => {
    // 23:59:30 and "00:00": the next 00:00 is 30 seconds away, which is too soon.
    expect(() => parseSendAt("00:00", local(23, 59, 30))).toThrow("at least 1 minute");
    // 23:30 and "00:15": the next 00:15 is tomorrow, 45 minutes away.
    expect(parseSendAt("00:15", local(23, 30)).getTime()).toBe(new Date(2026, 9, 20, 0, 15).getTime());
  });

  test("refuses text that is not a time", () => {
    for (const bad of ["tomorrow", "25:00", "10:61", "+5", "+m", "2026-10-20"]) expect(() => parseSendAt(bad, now)).toThrow();
  });
});

describe("send --at", () => {
  const at = "2099-01-02T15:30:00Z";

  test("the preview wraps the normal send body and calls nothing that writes", async () => {
    serve({ [CREATE]: CREATED });
    await send(CHANNEL, ["hello", "later"], { ...plain, at });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview).toMatchObject({ ok: true, dry_run: true, effect: "scheduled_message", method: "POST", path: DRAFTS, send_at: "2099-01-02T15:30:00.000Z" });
    expect(preview.send_at_local).toContain("2099");
    const body = preview.body;
    expect(Object.keys(body).sort()).toEqual(["draftDetails", "draftType", "innerThreadId", "message"]);
    expect(body.draftType).toBe("ScheduledDraft");
    expect(body.draftDetails).toEqual({ sendAt: String(Date.parse(at)) });
    expect(body.innerThreadId).toBe(CHANNEL);
    expect(body.message).toMatchObject({
      id: "-1", type: "Message", conversationid: CHANNEL, content: "<p>hello later</p>", messagetype: "RichText/Html",
      threadtype: "streamofdrafts", innerThreadId: CHANNEL, draftDetails: { sendAt: "2099-01-02T15:30:00.000Z" },
      postType: "Standard", from: ME_MRI,
    });
    expect(body.message.properties.draftId).toBe(body.message.clientmessageid);
    expect(body.message.properties).toMatchObject({ formatVariant: "TEAMS", mentions: "[]", files: "[]" });
    expect(calls("POST")).toHaveLength(0);
    expect(calls("DELETE")).toHaveLength(0);
    expect(existsSync(ledgerDir())).toBe(false);
  });

  test("the text preview names the local and UTC time", async () => {
    serve({ [CREATE]: CREATED });
    await send(CHANNEL, ["hello"], { json: false, verify: false, allowWrite: false, at });
    expect(logs[0]).toContain("Dry run (effect: scheduled message): would schedule for");
    expect(logs[0]).toContain("(2099-01-02T15:30:00.000Z): hello");
  });

  test("mentions work in a scheduled body", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 }, [CREATE]: CREATED });
    await send(CHANNEL, ["hi @{Alex Doe}"], { ...plain, at });
    const message = JSON.parse(logs.join("\n")).body.message;
    expect(message.content).toContain("schema.skype.com/Mention");
    const mentions = JSON.parse(message.properties.mentions);
    expect(mentions).toHaveLength(1);
    expect(mentions[0].mri).toMatch(/^8:orgid:/);
    expect(calls("POST").filter(item => item.path === DRAFTS)).toHaveLength(0);
  });

  test("--allow-write posts the wrapped body to the drafts endpoint and records the send", async () => {
    serve({ [CREATE]: CREATED });
    await send(CHANNEL, ["hello"], { ...plain, allowWrite: true, at });
    const posts = calls("POST");
    expect(posts).toHaveLength(1);
    expect(posts[0]!.path).toBe(DRAFTS);
    expect((posts[0]!.body as any).draftType).toBe("ScheduledDraft");
    const receipt = JSON.parse(logs.join("\n"));
    expect(receipt).toMatchObject({ ok: true, effect: "scheduled_message", conversation: CHANNEL, draft_id: "1792421450913", send_at: "2099-01-02T15:30:00.000Z" });
  });

  test("the same text at two times is two records, and the same time twice is refused", async () => {
    serve({ [CREATE]: CREATED });
    const write = { ...plain, allowWrite: true };
    await send(CHANNEL, ["hello"], { ...write, at: "2099-01-02T15:30:00Z" });
    await send(CHANNEL, ["hello"], { ...write, at: "2099-01-02T16:30:00Z" });
    expect(calls("POST")).toHaveLength(2);
    await expect(send(CHANNEL, ["hello"], { ...write, at: "2099-01-02T15:30:00Z" })).rejects.toThrow("Identical message already sent");
    expect(calls("POST")).toHaveLength(2);
    const key = (sendAt?: number) => ledgerKey({ conversationId: CHANNEL, content: "<p>hello</p>", mentionMris: [], contentType: "Text", sendAt });
    expect(key(Date.parse("2099-01-02T15:30:00Z"))).not.toBe(key(Date.parse("2099-01-02T16:30:00Z")));
    expect(key()).not.toBe(key(Date.parse("2099-01-02T15:30:00Z")));
  });

  test("a scheduled reply names its thread, as the web client did", async () => {
    serve({ [CREATE]: CREATED });
    await send(CHANNEL, ["hello"], { ...plain, at, threadTs: "1792421419988", allowWrite: true });
    const body = calls("POST").find(item => item.path === DRAFTS)!.body as any;
    const thread = `${CHANNEL};messageid=1792421419988`;
    expect(body.innerThreadId).toBe(thread);
    expect(body.message.innerThreadId).toBe(thread);
    expect(body.message.replyChainId).toBe("1792421419988");
    expect(body.message.conversationLink).toBe(`blah/${thread}`);
    expect(body.message.conversationid).toBe(CHANNEL);
  });

  test("--image with --at is refused before any request", async () => {
    serve({ [CREATE]: CREATED });
    await expect(send(CHANNEL, ["hello"], { ...plain, at, images: ["missing.png"] })).rejects.toThrow("--at cannot be used with --image");
    expect(fake!.requests).toHaveLength(0);
  });

  test("--verify with --at is refused", async () => {
    await expect(send(CHANNEL, ["hello"], { ...plain, at, verify: true })).rejects.toThrow("--verify cannot be used with --at");
  });

  test("a past time is refused before any request", async () => {
    serve({ [CREATE]: CREATED });
    await expect(send(CHANNEL, ["hello"], { ...plain, at: "2020-01-01T00:00:00Z" })).rejects.toThrow("in the past");
    expect(fake!.requests).toHaveLength(0);
  });
});

const draft = (overrides: Record<string, unknown> = {}) => ({
  sequenceId: 1,
  conversationid: "48:drafts",
  draftType: "ScheduledDraft",
  innerThreadId: CHANNEL,
  draftDetails: { sendAt: "4070908800000", shouldCloneAmsReferences: true },
  id: "1792421450913",
  clientmessageid: "3501233581758942095",
  messagetype: "RichText/Html",
  content: "<p>will send   this later</p>",
  properties: {},
  ...overrides,
});

describe("scheduled", () => {
  test("lists live scheduled messages and skips cancelled ones", async () => {
    serve({
      [LIST]: { status: 200, json: { drafts: [
        draft(),
        draft({ id: "1792421450914", content: "", properties: { deletetime: 1792421475020 } }),
        draft({ id: "1792421450915", draftType: "Draft" }),
      ], _metadata: {} } },
    });
    const items = await listScheduled();
    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({ id: "1792421450913", conversation: CHANNEL, send_at: "2099-01-01T00:00:00.000Z", text: "will send this later" });
    const request = fake!.requests[0]!;
    expect(request.query).toMatchObject({ view: "msnp24Equivalent|supportsMessageProperties", pageSize: "200", startTime: "1" });
  });

  test("a scheduled reply shows as its channel and thread root", async () => {
    serve({ [LIST]: { status: 200, json: { drafts: [draft({ innerThreadId: `${CHANNEL};messageid=1792421419988` })] } } });
    const label = (await listScheduled())[0]!.label;
    expect(label).toEndWith(" (thread 1792421419988)");
    expect(label).not.toContain(";messageid=");
  });

  test("a draft with files lists their addresses, so a cancel can say what stays on SharePoint", async () => {
    const url = "https://fake-my.sharepoint.com/personal/me/Documents/Microsoft%20Teams%20Chat%20Files/a.txt";
    serve({ [LIST]: { status: 200, json: { drafts: [draft({ properties: { files: JSON.stringify([{ fileName: "a.txt", fileInfo: { fileUrl: url } }]) } })] } } });
    expect((await listScheduled())[0]!.files).toEqual([url]);
  });

  test("cuts the text at 80 characters", async () => {
    serve({ [LIST]: { status: 200, json: { drafts: [draft({ content: `<p>${"x".repeat(200)}</p>` })] } } });
    expect((await listScheduled())[0]!.text).toHaveLength(80);
  });

  test("follows syncState pages through the drafts endpoint, not the link", async () => {
    const link = (state: string) => ({ syncState: `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/conversations/48:drafts/messages?startTime=1&syncState=${state}&pageSize=200` });
    serve({
      [LIST]: [
        { status: 200, json: { drafts: [draft()], _metadata: link("one") } },
        { status: 200, json: { drafts: [draft({ id: "1792421450916", content: "<p>second</p>" })], _metadata: link("two") } },
        { status: 200, json: { drafts: [], _metadata: link("two") } },
      ],
    });
    expect((await listScheduled()).map(item => item.id)).toEqual(["1792421450913", "1792421450916"]);
    expect(fake!.requests.map(item => item.query.syncState)).toEqual([undefined, "one", "two"]);
    expect(fake!.requests.every(item => item.path === DRAFTS)).toBe(true);
  });

  test("stops when the sync state does not change, and at the page cap", async () => {
    const next = (state: string) => ({ _metadata: { syncState: `https://teams.cloud.microsoft/x/messages?syncState=${state}` } });
    serve({ [LIST]: { status: 200, json: { drafts: [draft()], ...next("same") } } });
    await listScheduled();
    expect(fake!.requests).toHaveLength(2);
    let n = 0;
    serve({ [LIST]: Array.from({ length: 9 }, () => ({ status: 200, json: { drafts: [draft()], ...next(`s${n++}`) } })) });
    await listScheduled();
    expect(fake!.requests).toHaveLength(5);
  });

  test("--json prints the list", async () => {
    serve({ [LIST]: { status: 200, json: { drafts: [draft()] } } });
    await scheduled(true);
    expect(JSON.parse(logs.join("\n"))[0].id).toBe("1792421450913");
  });
});

describe("scheduled cancel", () => {
  const list = { [LIST]: { status: 200, json: { drafts: [draft()] } } as Route };

  test("refuses an id that is not in the live list", async () => {
    serve({ ...list, [CANCEL]: { status: 200, json: null } });
    await expect(scheduledCancel("1792421450999", { json: false, allowWrite: true })).rejects.toThrow("No scheduled message has the id 1792421450999");
    expect(calls("DELETE")).toHaveLength(0);
  });

  test("refuses an id that is a tombstone", async () => {
    serve({ [LIST]: { status: 200, json: { drafts: [draft({ properties: { deletetime: 1 } })] } }, [CANCEL]: { status: 200 } });
    await expect(scheduledCancel("1792421450913", { json: false, allowWrite: true })).rejects.toThrow("No scheduled message");
    expect(calls("DELETE")).toHaveLength(0);
  });

  test("refuses text that is not a draft id before any request", async () => {
    serve(list);
    await expect(scheduledCancel("../messages", { json: false, allowWrite: true })).rejects.toThrow("is not a draft id");
    expect(fake!.requests).toHaveLength(0);
  });

  test("the preview shows the message and never calls DELETE or POST", async () => {
    serve({ ...list, [CANCEL]: { status: 200 } });
    await scheduledCancel("1792421450913", { json: false, allowWrite: false });
    expect(logs[0]).toContain("Dry run (effect: scheduled cancel): would cancel 1792421450913");
    expect(logs[0]).toContain("will send this later");
    expect(calls("DELETE")).toHaveLength(0);
    expect(calls("POST")).toHaveLength(0);
  });

  test("--allow-write deletes the draft", async () => {
    serve({ ...list, [CANCEL]: { status: 200, text: "null" } });
    await scheduledCancel("1792421450913", { json: true, allowWrite: true });
    expect(calls("DELETE").map(item => item.path)).toEqual([`${DRAFTS}/1792421450913`]);
    expect(JSON.parse(logs.join("\n"))).toMatchObject({ ok: true, effect: "scheduled_cancel", id: "1792421450913" });
  });

  test("an uncertain failure says the outcome is unknown and does not retry", async () => {
    serve({ ...list, [CANCEL]: { status: 503 } });
    await expect(scheduledCancel("1792421450913", { json: false, allowWrite: true })).rejects.toThrow("Outcome unknown");
    expect(calls("DELETE")).toHaveLength(1);
  });
});

describe("command line", () => {
  test("--at is valid for send only, and scheduled cancel takes one id", async () => {
    const { runCli } = await import("./cli.ts");
    expect((await runCli(["history", "me", "--at", "+10m"])).stderr).toContain("Flag `--at` is not valid with `teams-cli history`");
    expect((await runCli(["scheduled", "cancel"])).stderr).toContain("Usage: teams-cli scheduled cancel <draft-id>");
    expect((await runCli(["scheduled", "--allow-write"])).stderr).toContain("not valid with `teams-cli scheduled`");
    expect((await runCli(["--help"])).stdout).toContain("scheduled cancel <draft-id>");
  });
});
