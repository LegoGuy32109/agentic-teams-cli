import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { editMessage } from "../src/commands/edit.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Reply, type Route } from "./fake-teams.ts";
import { OID } from "./jwt.ts";
import searchV2 from "./fixtures/search-v2.json";

const BASE = "/api/chatsvc/amer/v1/users/ME/conversations";
const MSG = "1791990424088";
const MSG_PATH = `${BASE}/48:notes/messages/${MSG}`;
const ME = `8:orgid:${OID}`;
const OTHER = "8:orgid:44444444-4444-4444-4444-444444444444";
const SEARCH = "POST /api/mt/amer/beta/users/searchV2";
const PUT = `PUT ${MSG_PATH}`;
const GET = `GET ${MSG_PATH}`;
const mine = (content: string, extra: Record<string, unknown> = {}): Reply => ({
  status: 200,
  json: { id: MSG, content, messagetype: "RichText/Html", imdisplayname: "Test User", from: `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${ME}`, properties: {}, ...extra },
});
const plain = { json: false, allowWrite: false } as const;
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
const originalLog = console.log;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}
const puts = () => fake!.requests.filter(item => item.method === "PUT");

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
  process.exitCode = 0;
});

describe("edit preview", () => {
  test("sends no PUT and shows the new text", async () => {
    serve({ [GET]: mine("<p>old</p>"), [PUT]: { status: 200, text: "" } });
    await editMessage("me", `48:notes#${MSG}`, ["new", "text"], plain);
    expect(puts()).toHaveLength(0);
    expect(logs[0]).toBe(`Dry run (effect: message): would edit 48:notes#${MSG} in Notes to self to: new text`);
  });

  test("the JSON preview has the edit body and no client message id", async () => {
    serve({ [GET]: mine("<p>old</p>") });
    await editMessage("me", `48:notes#${MSG}`, ["new"], { ...plain, json: true });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview).toMatchObject({ ok: true, dry_run: true, effect: "message", method: "PUT", message_id: MSG });
    expect(Object.keys(preview.body).sort()).toEqual(["contenttype", "content", "imdisplayname", "messagetype", "properties", "skypeeditedid"].sort());
    expect(preview.body.skypeeditedid).toBe(MSG);
    expect(preview.body.content).toBe("<p>new</p>");
    expect(preview.body.clientmessageid).toBeUndefined();
  });
});

describe("edit a thread reply", () => {
  test("adds replyChainId and the thread link from the reply's rootMessageId", async () => {
    serve({ [GET]: mine("<p>old</p>", { rootMessageId: "1791990400000" }) });
    await editMessage("me", `48:notes;messageid=1791990400000#${MSG}`, ["new"], { ...plain, json: true });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.path).toEndWith(`/48%3Anotes/messages/${MSG}`);
    expect(preview.body.replyChainId).toBe("1791990400000");
    expect(preview.body.conversationLink).toBe("conversation/48:notes;messageid=1791990400000");
  });

  test("a root post, whose rootMessageId is its own id, gets no thread fields", async () => {
    serve({ [GET]: mine("<p>old</p>", { rootMessageId: MSG }) });
    await editMessage("me", `48:notes#${MSG}`, ["new"], { ...plain, json: true });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.body.replyChainId).toBeUndefined();
  });

  test("a reply ref from another conversation is refused", async () => {
    serve({ [GET]: mine("<p>old</p>") });
    await expect(editMessage("me", `19:other@thread.tacv2;messageid=1791990400000#${MSG}`, ["new"], plain)).rejects.toThrow("names another conversation");
  });
});

describe("edit write", () => {
  test("sends the PUT with the edit body, then re-reads for the receipt", async () => {
    serve({ [GET]: [mine("<p>old</p>"), mine("<p>new text</p>")], [PUT]: { status: 200, text: "" } });
    await editMessage("me", `48:notes#${MSG}`, ["new", "text"], { ...plain, allowWrite: true, json: true });
    expect(puts()).toHaveLength(1);
    expect(puts()[0]!.path).toBe(MSG_PATH);
    expect(puts()[0]!.body).toMatchObject({ content: "<p>new text</p>", skypeeditedid: MSG, messagetype: "RichText/Html", contenttype: "Text" });
    expect((puts()[0]!.body as any).clientmessageid).toBeUndefined();
    const receipt = JSON.parse(logs.join("\n"));
    expect(receipt).toMatchObject({ ok: true, id: `48:notes#${MSG}`, text: "new text", verification: { ok: true } });
    expect(process.exitCode).toBe(0);
  });

  test("an edit that a mention lands in resolves the mention and sets its entry", async () => {
    serve({ [GET]: [mine("<p>old</p>"), mine("<p>hi</p>")], [SEARCH]: { status: 200, json: searchV2 }, [PUT]: { status: 200, text: "" } });
    await editMessage("me", `48:notes#${MSG}`, ["hi", "@{Alex Doe}"], { ...plain, allowWrite: true, json: true });
    const body = puts()[0]!.body as any;
    const mentions = JSON.parse(body.properties.mentions);
    expect(mentions).toEqual([expect.objectContaining({ mri: "8:orgid:33333333-3333-3333-3333-333333333333", displayName: "Alex Doe" })]);
    expect(body.content).toContain("itemtype=\"http://schema.skype.com/Mention\"");
  });

  test("a message from someone else is refused before any PUT", async () => {
    serve({ [GET]: mine("<p>theirs</p>", { from: `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${OTHER}` }), [PUT]: { status: 200, text: "" } });
    await expect(editMessage("me", `48:notes#${MSG}`, ["x"], { ...plain, allowWrite: true })).rejects.toThrow("You can only edit your own messages.");
    expect(puts()).toHaveLength(0);
  });

  test("a deleted message is refused before any PUT", async () => {
    serve({ [GET]: mine("", { properties: { deletetime: 1 } }), [PUT]: { status: 200, text: "" } });
    await expect(editMessage("me", `48:notes#${MSG}`, ["x"], { ...plain, allowWrite: true })).rejects.toThrow("That message was deleted.");
    expect(puts()).toHaveLength(0);
  });

  test("a ref to another conversation is refused before any request", async () => {
    serve({ [GET]: mine("<p>x</p>") });
    await expect(editMessage("me", `19:dev@thread.tacv2#${MSG}`, ["x"], { ...plain })).rejects.toThrow("names another conversation");
    expect(fake!.requests).toHaveLength(0);
  });

  test("a stream that is read-only is refused", async () => {
    serve({});
    await expect(editMessage("mentions", `48:mentions#${MSG}`, ["x"], { ...plain })).rejects.toThrow("That stream is read-only.");
    expect(fake!.requests).toHaveLength(0);
  });

  test("a 5xx on the PUT is reported as an unknown outcome", async () => {
    serve({ [GET]: mine("<p>old</p>"), [PUT]: { status: 503, text: "busy" } });
    await expect(editMessage("me", `48:notes#${MSG}`, ["x"], { ...plain, allowWrite: true })).rejects.toThrow("Outcome unknown: Teams did not confirm this edit.");
  });
});

describe("edit with --html mentions", () => {
  const MENTION = "http://schema.skype.com/Mention";
  const span = (itemid: string, name: string) => `<readonly class="skipProofing" itemtype="${MENTION}" spellcheck="false"><span itemtype="${MENTION}" itemscope itemid="${itemid}">${name}</span></readonly>`;

  test("the edit body has the same content and mentions as a send", async () => {
    serve({ [GET]: mine("<p>old</p>"), [SEARCH]: { status: 200, json: searchV2 } });
    await editMessage("me", `48:notes#${MSG}`, [], { ...plain, json: true, html: "<ol><li>@{Alex Doe}</li><li><b>@{Alex Doerr}</b></li></ol>" });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.body.content).toBe(`<ol><li>${span("0", "Alex Doe")}</li><li><b>${span("1", "Alex Doerr")}</b></li></ol>`);
    expect(JSON.parse(preview.body.properties.mentions).map((entry: any) => [entry.itemid, entry.displayName])).toEqual([["0", "Alex Doe"], ["1", "Alex Doerr"]]);
  });

  test("hand-written mention markup is refused", async () => {
    serve({ [GET]: mine("<p>old</p>") });
    await expect(editMessage("me", `48:notes#${MSG}`, [], { ...plain, html: `<span itemtype="${MENTION}">x</span>` })).rejects.toThrow("--html cannot contain hand-written mention markup.");
  });

  test("an <ol> report verifies after the write", async () => {
    serve({ [GET]: [mine("<p>old</p>"), mine("<ol><li><p>a</p></li><li><p>b</p></li></ol>")], [PUT]: { status: 200, text: "" } });
    await editMessage("me", `48:notes#${MSG}`, [], { ...plain, allowWrite: true, json: true, html: "<ol><li>a</li><li>b</li></ol>" });
    expect(JSON.parse(logs.join("\n")).verification.ok).toBe(true);
    expect(process.exitCode).toBe(0);
  });
});

describe("edit with mentions in rich text", () => {
  test("a table cell mention resolves, code is untouched, and the link stays in the body", async () => {
    serve({ [GET]: mine("<p>old</p>"), [SEARCH]: { status: 200, json: searchV2 } });
    await editMessage("me", `48:notes#${MSG}`, ["| who |\n|---|\n| @{Alex Doe} |\nsee <https://x.test/a>\n```powershell\n$h = @{Alex Doe=1}\n```"], { ...plain, json: true, format: "rich" });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.body.content).toContain("<td><p><readonly");
    expect(preview.body.content).toContain("$h&nbsp;=&nbsp;@{Alex&nbsp;Doe=1}");
    expect(JSON.parse(preview.body.properties.mentions)).toHaveLength(1);
    expect(JSON.parse(preview.body.properties.links).map((link: any) => link.url)).toEqual(["https://x.test/a"]);
    expect(puts()).toHaveLength(0);
  });
});
