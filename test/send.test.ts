import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { writeFileSync, mkdirSync, existsSync } from "node:fs";
import { join } from "node:path";
import { send, ledgerClear } from "../src/commands/send.ts";
import { ledgerKey, LedgerRefusal, SENT_WINDOW_MS } from "../src/ledger.ts";
import { cacheDir } from "../src/config.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";
import { TID } from "./jwt.ts";
import searchV2 from "./fixtures/search-v2.json";

const SEND = "POST /api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages";
const DEV_SEND = "POST /api/chatsvc/amer/v1/users/ME/conversations/19:dev@thread.tacv2/messages";
const SEARCH = "POST /api/mt/amer/beta/users/searchV2";
const ME_MRI = "8:orgid:11111111-1111-1111-1111-111111111111";
const OK = { status: 201, json: { OriginalArrivalTime: 1791990424088 } };
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
let errors: string[] = [];
const originalLog = console.log;
const originalError = console.error;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}

const postsTo = (suffix: string) => fake!.requests.filter(item => item.method === "POST" && item.path.endsWith(suffix)).length;
const ledgerFile = (key: string) => join(cacheDir(), TID, "ledger", `${key}.json`);
const noteKey = (content: string, mris: string[] = []) => ledgerKey({ conversationId: "48:notes", content, mentionMris: mris, contentType: "Text" });

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

const plain = { json: false, verify: false, allowWrite: false } as const;

describe("preview", () => {
  test("a preview sends no message and creates no ledger file", async () => {
    serve({ [SEND]: OK });
    await send("me", ["hello"], plain);
    expect(postsTo("/messages")).toBe(0);
    expect(logs[0]).toBe("Dry run (effect: message): would send to Notes to self: hello");
    expect(existsSync(join(cacheDir(), TID, "ledger"))).toBe(false);
  });

  test("the JSON preview has the documented shape", async () => {
    serve({ [SEND]: OK });
    await send("me", ["hello"], { ...plain, json: true });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview).toMatchObject({
      ok: true, dry_run: true, effect: "message", method: "POST",
      path: "/api/chatsvc/amer/v1/users/ME/conversations/48%3Anotes/messages",
      conversation: { id: "48:notes", label: "Notes to self" },
    });
    expect(preview.body.content).toBe("<p>hello</p>");
    expect(preview.warning).toBeUndefined();
  });

  test("the preview resolves a mention with a read and still sends nothing", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 }, [DEV_SEND]: OK });
    await send("19:dev@thread.tacv2", ["hi @{Alex Doe}"], plain);
    expect(postsTo("/messages")).toBe(0);
    expect(fake!.requests.some(item => item.path.endsWith("/users/searchV2"))).toBe(true);
  });
});

describe("write", () => {
  test("--allow-write sends exactly one message and prints the receipt", async () => {
    serve({ [SEND]: OK });
    await send("me", ["hello"], { ...plain, allowWrite: true });
    expect(postsTo("/messages")).toBe(1);
    expect(logs[0]).toBe("Sent to Notes to self as 48:notes#1791990424088");
    expect(logs[1]).toBe("https://teams.cloud.microsoft/l/message/48%3Anotes/1791990424088");
    const sent = fake!.requests[0]!.body as any;
    expect(sent.content).toBe("<p>hello</p>");
    expect(sent.from).toBe(ME_MRI);
    expect(sent.imdisplayname).toBe("Test User");
  });

  test("a duplicate send inside the window is refused and not sent", async () => {
    serve({ [SEND]: OK });
    await send("me", ["hello"], { ...plain, allowWrite: true });
    const refusal = await send("me", ["hello"], { ...plain, allowWrite: true }).then(() => undefined, e => e as Error);
    expect(refusal).toBeInstanceOf(LedgerRefusal);
    expect(refusal!.message).toMatch(/^Identical message already sent to Notes to self at .+ as 48:notes#1791990424088\. Not sending again\.$/);
    expect(postsTo("/messages")).toBe(1);
  });

  test("an uncertain send is reported as unknown and blocks a resend", async () => {
    serve({ [SEND]: "network" });
    const first = await send("me", ["hello"], { ...plain, allowWrite: true }).then(() => undefined, e => e as Error);
    expect(first!.message).toBe("Outcome unknown: Teams did not confirm this send. Do not resend. Check with: teams-cli history me --top=5");
    const second = await send("me", ["hello"], { ...plain, allowWrite: true }).then(() => undefined, e => e as Error);
    expect(second).toBeInstanceOf(LedgerRefusal);
    expect(second!.message).toContain("A previous identical send may have posted (outcome unknown).");
    expect(second!.message).toContain(`teams-cli ledger clear ${noteKey("<p>hello</p>")}`);
    expect(postsTo("/messages")).toBe(1);
  });

  test("ledger clear removes the record and allows the send again", async () => {
    serve({ [SEND]: "network" });
    await send("me", ["hello"], { ...plain, allowWrite: true }).catch(() => undefined);
    await ledgerClear(noteKey("<p>hello</p>"));
    expect(logs.at(-1)).toBe("Cleared.");
    serve({ [SEND]: OK });
    await send("me", ["hello"], { ...plain, allowWrite: true });
    expect(postsTo("/messages")).toBe(1);
  });

  test("a leftover reserved file blocks the send until it is old", async () => {
    const key = noteKey("<p>leftover</p>");
    mkdirSync(join(cacheDir(), TID, "ledger"), { recursive: true });
    writeFileSync(ledgerFile(key), JSON.stringify({ state: "reserved", at: new Date().toISOString() }));
    serve({ [SEND]: OK });
    await expect(send("me", ["leftover"], { ...plain, allowWrite: true })).rejects.toThrow(LedgerRefusal);
    expect(postsTo("/messages")).toBe(0);
    writeFileSync(ledgerFile(key), JSON.stringify({ state: "reserved", at: new Date(Date.now() - 25 * 3600_000).toISOString() }));
    await send("me", ["leftover"], { ...plain, allowWrite: true });
    expect(postsTo("/messages")).toBe(1);
  });

  test("a sent record older than the window does not block", async () => {
    const key = noteKey("<p>old</p>");
    mkdirSync(join(cacheDir(), TID, "ledger"), { recursive: true });
    writeFileSync(ledgerFile(key), JSON.stringify({ state: "sent", at: new Date(Date.now() - SENT_WINDOW_MS - 60_000).toISOString(), id: "48:notes#1" }));
    serve({ [SEND]: OK });
    await send("me", ["old"], { ...plain, allowWrite: true });
    expect(postsTo("/messages")).toBe(1);
  });

  test("a definite failure removes the record so the send can be retried", async () => {
    serve({ [SEND]: [{ status: 400, json: { errorCode: 400, message: "Bad request" } }, OK] });
    await expect(send("me", ["hello"], { ...plain, allowWrite: true })).rejects.toThrow("POST /users/ME/conversations/48%3Anotes/messages failed: 400 Bad request");
    await send("me", ["hello"], { ...plain, allowWrite: true });
    expect(postsTo("/messages")).toBe(2);
  });

  test("two racing identical sends post exactly once", async () => {
    serve({ [SEND]: OK });
    const results = await Promise.allSettled([
      send("me", ["race"], { ...plain, allowWrite: true }),
      send("me", ["race"], { ...plain, allowWrite: true }),
    ]);
    expect(results.filter(result => result.status === "fulfilled").length).toBe(1);
    const rejected = results.find(result => result.status === "rejected") as PromiseRejectedResult;
    expect(rejected.reason).toBeInstanceOf(LedgerRefusal);
    expect(postsTo("/messages")).toBe(1);
  });
});

describe("verify", () => {
  test("a matching re-read reports ok", async () => {
    serve({
      [SEND]: OK,
      "GET /api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages/1791990424088": { status: 200, json: { content: "<p>hello</p>" } },
    });
    await send("me", ["hello"], { ...plain, allowWrite: true, verify: true, json: true });
    const receipt = JSON.parse(logs.join("\n"));
    expect(receipt.verification).toEqual({ ok: true, expected: "hello", actual: "hello" });
    expect(process.exitCode).toBe(0);
  });

  test("a different re-read fails with exit code 1", async () => {
    serve({
      [SEND]: OK,
      "GET /api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages/1791990424088": { status: 200, json: { content: "<p>other</p>" } },
    });
    await send("me", ["hello"], { ...plain, allowWrite: true, verify: true });
    expect(logs.some(line => line === "Verification failed. Expected: hello. Actual: other")).toBe(true);
    expect(process.exitCode).toBe(1);
  });
});

describe("mentions and text", () => {
  test("one mention matches the byte-exact browser shape", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 }, [DEV_SEND]: OK });
    await send("19:dev@thread.tacv2", ["hi @{Alex Doe}"], { ...plain, json: true });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.body.content).toBe(`<p>hi <readonly class="skipProofing" itemtype="http://schema.skype.com/Mention" spellcheck="false"><span itemtype="http://schema.skype.com/Mention" itemscope itemid="0">Alex Doe</span></readonly></p>`);
    expect(preview.body.properties.mentions).toBe(`[{"@type":"http://schema.skype.com/Mention","itemid":"0","mri":"8:orgid:33333333-3333-3333-3333-333333333333","mentionType":"person","displayName":"Alex Doe"}]`);
  });

  test("two mentions of one person get item ids 0 and 1", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    await send("19:dev@thread.tacv2", ["@{Alex Doe} and @{Alex Doe}"], { ...plain, json: true });
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.body.content).toContain('itemid="0">Alex Doe</span>');
    expect(preview.body.content).toContain('itemid="1">Alex Doe</span>');
    expect(JSON.parse(preview.body.properties.mentions).map((entry: any) => entry.itemid)).toEqual(["0", "1"]);
  });

  test("an ambiguous mention fails before any request is sent", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    await expect(send("me", ["@{Alex}"], plain)).rejects.toThrow('Mention "Alex" is ambiguous:');
    expect(postsTo("/messages")).toBe(0);
  });

  test("a plain @Word warns on stderr and in the preview", async () => {
    serve({});
    await send("me", ["ping @alex"], { ...plain, json: true });
    expect(errors[0]).toContain("Message contains @alex as plain text");
    expect(JSON.parse(logs.join("\n")).warning).toContain("It will not notify anyone.");
  });

  test("an email address does not warn", async () => {
    serve({});
    await send("me", ["mail adoe@example.test"], plain);
    expect(errors).toEqual([]);
  });

  test("rich format turns bullets into a list", async () => {
    serve({});
    await send("me", ["- one\n- two"], { ...plain, json: true, format: "rich" });
    expect(JSON.parse(logs.join("\n")).body.content).toBe("<ul><li>one</li><li>two</li></ul>");
  });

  test("rich format sends bold, a numbered list and a link with its preview card in the Teams shape", async () => {
    serve({});
    const teams = globalThis.fetch;
    globalThis.fetch = (async (input: any, init: RequestInit) =>
      String(input).startsWith("https://us-prod.asyncgw.teams.microsoft.com/urlp/v1/url/info")
        ? Response.json({ title: "Docs", description: "About docs", thumbnail: "https://us-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail?url=x", thumbnail_meta: { width: 320, height: 180 } })
        : teams(input, init)) as typeof fetch;
    await send("me", ["*Title*\n1. one\n2. two <https://x.test/a|docs>"], { ...plain, json: true, format: "rich" });
    const body = JSON.parse(logs.join("\n")).body;
    expect(body.content).toContain("<p><strong>Title</strong></p><ol><li>one</li><li>two <a href=\"https://x.test/a\"");
    expect(body.content).toContain(">docs</a></li></ol>");
    expect(JSON.parse(body.properties.links)).toEqual([{
      "@type": "http://schema.skype.com/HyperLink", itemid: "0", url: "https://x.test/a", previewenabled: true,
      preview: { previewurl: "https://us-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail?url=x", previewmeta: { width: 320, height: 180 }, title: "Docs", description: "About docs" },
    }]);
    expect(errors).toEqual([]);
  });

  test("when the preview cannot be read the message still sends, without a card, and says so", async () => {
    serve({});
    const teams = globalThis.fetch;
    globalThis.fetch = (async (input: any, init: RequestInit) =>
      String(input).includes("asyncgw") ? new Response("no", { status: 503 }) : teams(input, init)) as typeof fetch;
    await send("me", ["see <https://x.test/a>"], { ...plain, json: true, format: "rich" });
    const body = JSON.parse(logs.join("\n")).body;
    expect(JSON.parse(body.properties.links)).toEqual([{ "@type": "http://schema.skype.com/HyperLink", itemid: "0", url: "https://x.test/a", previewenabled: false }]);
    expect(errors.join("\n")).toContain("Link preview unavailable. The link is sent without a card.");
  });

  test("rich text sends fenced code, a heading and colors, and mention syntax inside code is left alone", async () => {
    serve({});
    await send("me", ["# Plan\n[red]careful[/]\n```powershell\n$h = @{a=1}\nWrite-Host \"hi\"\n```"], { ...plain, json: true, format: "rich" });
    const content = JSON.parse(logs.join("\n")).body.content;
    expect(content).toStartWith('<h1>Plan</h1><p><span style="color:#B6424C">careful</span></p><p itemtype="http://schema.skype.com/CodeBlockEditor"');
    expect(content).toContain('<pre class="language-powershell language-was-manually-selected"');
    expect(content).toContain("$h&nbsp;=&nbsp;@{a=1}<br>");
    expect(content).toContain('<span class="hljs-function">Write-Host</span>');
    expect(errors).toEqual([]);
  });

  test("--html without a token is sent raw", async () => {
    serve({});
    await send("me", [], { ...plain, json: true, html: "<p><b>bold</b></p>" });
    expect(JSON.parse(logs.join("\n")).body.content).toBe("<p><b>bold</b></p>");
  });

  test("invalid flag combinations fail with a clear message", async () => {
    serve({});
    await expect(send("me", ["text"], { ...plain, html: "<p>x</p>" })).rejects.toThrow("Use either <text> or --html.");
    await expect(send("me", ["text"], { ...plain, format: "rich", html: undefined, threadTs: "1" })).rejects.toThrow("--thread-ts needs the root message id");
    await expect(send("me", ["text"], { ...plain, format: "fancy" })).rejects.toThrow("--format must be plain or rich.");
    await expect(send("me", [], plain)).rejects.toThrow("Usage: teams-cli send");
  });

  test("a message over 100,000 bytes fails before any request", async () => {
    serve({});
    await expect(send("me", ["x".repeat(101_000)], plain)).rejects.toThrow("above the 100,000 byte limit");
    expect(fake!.requests.length).toBe(0);
  });
});

const MENTION = "http://schema.skype.com/Mention";
const mentionSpan = (itemid: string, name: string) => `<readonly class="skipProofing" itemtype="${MENTION}" spellcheck="false"><span itemtype="${MENTION}" itemscope itemid="${itemid}">${name}</span></readonly>`;
const DOE = "8:orgid:33333333-3333-3333-3333-333333333333";
const DOERR = "8:orgid:44444444-4444-4444-4444-444444444444";
const NOTES_MESSAGE = "GET /api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages/1791990424088";

/** Previews an --html message and returns the parsed JSON preview. */
async function previewHtml(html: string) {
  logs.length = 0;
  await send("me", [], { ...plain, json: true, html });
  return JSON.parse(logs.join("\n"));
}

describe("mentions in --html", () => {
  test("a token in a paragraph becomes the browser mention shape", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const preview = await previewHtml("<p>hi @{Alex Doe}, welcome</p>");
    expect(preview.body.content).toBe(`<p>hi ${mentionSpan("0", "Alex Doe")}, welcome</p>`);
    expect(preview.body.properties.mentions).toBe(`[{"@type":"${MENTION}","itemid":"0","mri":"${DOE}","mentionType":"person","displayName":"Alex Doe"}]`);
  });

  test("properties.mentions is identical to the text mode JSON", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const htmlPreview = await previewHtml("<p>hi @{Alex Doe} and @{Alex Doerr}</p>");
    logs.length = 0;
    await send("me", ["hi @{Alex Doe} and @{Alex Doerr}"], { ...plain, json: true });
    const textPreview = JSON.parse(logs.join("\n"));
    expect(htmlPreview.body.properties.mentions).toBe(textPreview.body.properties.mentions);
    expect(htmlPreview.body.content).toBe(textPreview.body.content);
  });

  test("a token in a list item and in bold resolves", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const preview = await previewHtml("<ol><li>owner @{Alex Doe}</li><li><b>@{Alex Doerr}</b> reviews</li></ol>");
    expect(preview.body.content).toBe(`<ol><li>owner ${mentionSpan("0", "Alex Doe")}</li><li><b>${mentionSpan("1", "Alex Doerr")}</b> reviews</li></ol>`);
    expect(JSON.parse(preview.body.properties.mentions).map((entry: any) => [entry.itemid, entry.mri])).toEqual([["0", DOE], ["1", DOERR]]);
  });

  test("the same person twice gets item ids 0 and 1", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const preview = await previewHtml("<p>@{Alex Doe}</p><p>again @{Alex Doe}</p>");
    expect(preview.body.content).toContain('itemid="0">Alex Doe</span>');
    expect(preview.body.content).toContain('itemid="1">Alex Doe</span>');
    expect(JSON.parse(preview.body.properties.mentions).map((entry: any) => entry.itemid)).toEqual(["0", "1"]);
  });

  test("a token in an attribute, <code>, <pre>, or a comment is not resolved", async () => {
    serve({});
    const html = `<p><a href="https://x.test/?u=@{Alex Doe}" title='@{Nobody}'>link</a> <code>@{Alex Doe}</code></p><pre>@{Alex Doe}</pre><!-- @{Alex Doe} -->`;
    const preview = await previewHtml(html);
    expect(preview.body.content).toBe(html);
    expect(preview.body.properties.mentions).toBe("[]");
    expect(fake!.requests.length).toBe(0);
  });

  test("a token after a skipped one still resolves, with item id 0", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const preview = await previewHtml("<code>@{Alex Doe}</code><p>@{Alex Doe}</p>");
    expect(preview.body.content).toBe(`<code>@{Alex Doe}</code><p>${mentionSpan("0", "Alex Doe")}</p>`);
  });

  test("a missing name fails before any request is sent", async () => {
    serve({ [SEARCH]: { status: 200, json: { value: [] } } });
    await expect(send("me", [], { ...plain, html: "<p>@{Nobody Here}</p>" })).rejects.toThrow('Could not find a Teams user matching "Nobody Here".');
    expect(postsTo("/messages")).toBe(0);
  });

  test("an ambiguous name fails, and so does an unclosed token", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    await expect(send("me", [], { ...plain, html: "<p>@{Alex}</p>" })).rejects.toThrow('Mention "Alex" is ambiguous:');
    await expect(send("me", [], { ...plain, html: "<p>@{Alex <b>Doe</b>}</p>" })).rejects.toThrow("Mention is not closed");
    expect(postsTo("/messages")).toBe(0);
  });

  test("hand-written mention markup is still refused", async () => {
    serve({});
    const message = "--html cannot contain hand-written mention markup. Write @{Name} in the text and the CLI builds the mention.";
    await expect(send("me", [], { ...plain, html: `<span itemtype="${MENTION}">x</span>` })).rejects.toThrow(message);
    await expect(send("me", [], { ...plain, html: `<span itemtype='${MENTION.toUpperCase()}'>x</span>` })).rejects.toThrow(message);
    expect(fake!.requests.length).toBe(0);
  });

  test("@@{ writes a literal @{ without a lookup", async () => {
    serve({});
    const preview = await previewHtml("<p>Write @@{Name} to mention. Also @@{Name}.</p>");
    expect(preview.body.content).toBe("<p>Write @{Name} to mention. Also @{Name}.</p>");
    expect(preview.body.properties.mentions).toBe("[]");
    expect(fake!.requests.length).toBe(0);
  });

  test("an entity in the name is decoded for the lookup", async () => {
    serve({ [SEARCH]: { status: 200, json: { value: [{ displayName: "Doe & Co", mri: DOE }, { displayName: "Doe & Cohen", mri: DOERR }] } } });
    const preview = await previewHtml("<p>@{Doe &amp; Co}</p>");
    expect(preview.body.content).toBe(`<p>${mentionSpan("0", "Doe &amp; Co")}</p>`);
    expect(JSON.parse(preview.body.properties.mentions)[0].displayName).toBe("Doe & Co");
  });

  test("the preview shows @Display Name and the mention does not warn", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    await send("me", [], { ...plain, html: "<p>hi @{Alex Doe}</p>" });
    expect(logs[0]).toBe("Dry run (effect: message): would send to Notes to self: hi @Alex Doe");
    expect(errors).toEqual([]);
  });

  test("a plain @Word in --html still warns", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const preview = await previewHtml("<p>@{Alex Doe} and @alex</p>");
    expect(preview.warning).toContain("Message contains @alex as plain text");
    expect(preview.warning).not.toContain("@Alex Doe");
  });

  test("a real send carries the entries and the verify read-back matches", async () => {
    const read = `<p>hi ${mentionSpan("0", "Alex Doe")}</p>`;
    serve({ [SEARCH]: { status: 200, json: searchV2 }, [SEND]: OK, [NOTES_MESSAGE]: { status: 200, json: { content: read } } });
    await send("me", [], { ...plain, json: true, html: "<p>hi @{Alex Doe}</p>", allowWrite: true, verify: true });
    const sent = fake!.requests.find(item => item.method === "POST" && item.path.endsWith("/messages"))!.body as any;
    expect(JSON.parse(sent.properties.mentions)).toHaveLength(1);
    expect(JSON.parse(logs.join("\n")).verification).toEqual({ ok: true, expected: "hi @Alex Doe", actual: "hi @Alex Doe" });
  });
});

describe("verify with list and break markup", () => {
  test.each([
    ["<ol><li>a</li><li>b</li></ol>", "<ol><li><p>a</p></li><li><p>b</p></li></ol>"],
    ["<ul><li>a</li><li>b</li></ul><p>x<br>y</p>", "<ul>\r\n<li>a</li>\r\n<li>b</li>\r\n</ul><p>x<br />y</p>"],
    ["<p>one<br>two</p>", "<p>one</p><p>two</p>"],
  ])("%s reads back as ok", async (sent, read) => {
    serve({ [SEND]: OK, [NOTES_MESSAGE]: { status: 200, json: { content: read } } });
    await send("me", [], { ...plain, json: true, html: sent, allowWrite: true, verify: true });
    expect(JSON.parse(logs.join("\n")).verification.ok).toBe(true);
    expect(process.exitCode).toBe(0);
  });

  test("a changed word still fails verification", async () => {
    serve({ [SEND]: OK, [NOTES_MESSAGE]: { status: 200, json: { content: "<ol><li>a</li><li>c</li></ol>" } } });
    await send("me", [], { ...plain, json: true, html: "<ol><li>a</li><li>b</li></ol>", allowWrite: true, verify: true });
    expect(JSON.parse(logs.join("\n")).verification.ok).toBe(false);
    expect(process.exitCode).toBe(1);
  });
});

const RICH_PREVIEW_URL = "https://us-prod.asyncgw.teams.microsoft.com/urlp/v1/url/info";
const mentionEntries = (preview: any) => JSON.parse(preview.body.properties.mentions) as Array<{ itemid: string; mri: string; displayName: string }>;

describe("mentions with rich text", () => {
  async function previewRich(text: string) {
    logs.length = 0;
    await send("me", [text], { ...plain, json: true, format: "rich" });
    return JSON.parse(logs.join("\n"));
  }

  test("a mention works in a table cell, a quote, a colored span and a list item, in document order", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const preview = await previewRich("| who | role |\n|---|---|\n| @{Alex Doe} | owner |\n> quoted @{Alex Doerr}\n[red]careful @{Alex Doe}[/]\n- item @{Alex Doerr}");
    const content: string = preview.body.content;
    expect(content).toContain(`<td><p>${mentionSpan("0", "Alex Doe")}</p></td>`);
    expect(content).toContain(`<blockquote><p>quoted ${mentionSpan("1", "Alex Doerr")}</p></blockquote>`);
    expect(content).toContain(`<p><span style="color:#B6424C">careful ${mentionSpan("2", "Alex Doe")}</span></p>`);
    expect(content).toContain(`<ul><li>item ${mentionSpan("3", "Alex Doerr")}</li></ul>`);
    expect(mentionEntries(preview).map(entry => [entry.itemid, entry.mri])).toEqual([["0", DOE], ["1", DOERR], ["2", DOE], ["3", DOERR]]);
    expect(errors).toEqual([]);
  });

  test("a table, a code block and one mention give exactly one entry, and the code is not touched", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const preview = await previewRich("| a | b |\n|---|---|\n| @{Alex Doe} | 1 |\n```powershell\n$h = @{Alex Doe=1}\n```");
    expect(mentionEntries(preview)).toHaveLength(1);
    expect(preview.body.content).toContain("$h&nbsp;=&nbsp;@{Alex&nbsp;Doe=1}");
    expect(preview.body.content).toContain(mentionSpan("0", "Alex Doe"));
    expect(fake!.requests.filter(item => item.path.endsWith("/users/searchV2")).length).toBe(1);
  });

  test("a token only inside a code block makes no lookup and no entry", async () => {
    serve({});
    const preview = await previewRich("```\n@{Nobody Here}\n```");
    expect(preview.body.content).toContain("@{Nobody&nbsp;Here}");
    expect(preview.body.properties.mentions).toBe("[]");
    expect(fake!.requests.length).toBe(0);
  });

  test("text that a link preview card returns is never resolved as a mention", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const teams = globalThis.fetch;
    globalThis.fetch = (async (input: any, init: RequestInit) =>
      String(input).startsWith(RICH_PREVIEW_URL)
        ? Response.json({ title: "@{Alex Doerr}", description: "@{Alex Doerr}", thumbnail: "https://us-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail?url=x", thumbnail_meta: { width: 1, height: 1 } })
        : teams(input, init)) as typeof fetch;
    const preview = await previewRich("@{Alex Doe} see <https://x.test/a|docs>");
    expect(mentionEntries(preview).map(entry => entry.mri)).toEqual([DOE]);
    expect(preview.body.content).not.toContain("Alex Doerr");
    const links = JSON.parse(preview.body.properties.links);
    expect(links).toHaveLength(1);
    expect(links[0].preview.title).toBe("@{Alex Doerr}");
    expect(fake!.requests.filter(item => item.path.endsWith("/users/searchV2")).length).toBe(1);
  });
});

describe("mentions in rich --html", () => {
  const TEAMS_CODE = '<p itemtype="http://schema.skype.com/CodeBlockEditor"><pre class="language-powershell"><code>$h&nbsp;=&nbsp;<span class="hljs-built_in">@{Alex Doe=1}</span></code></pre></p>';

  test("a table, a quote, a colored span and a code block resolve only the text mention", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const html = `<table><thead><tr><th><p>who</p></th></tr></thead><tbody><tr><td><p>@{Alex Doe}</p></td></tr></tbody></table><blockquote><p>@{Alex Doerr} said</p></blockquote><p><span style="color:#B6424C">@{Alex Doe}</span></p>${TEAMS_CODE}`;
    const preview = await previewHtml(html);
    expect(mentionEntries(preview).map(entry => [entry.itemid, entry.mri])).toEqual([["0", DOE], ["1", DOERR], ["2", DOE]]);
    expect(preview.body.content).toContain(`<td><p>${mentionSpan("0", "Alex Doe")}</p></td>`);
    expect(preview.body.content).toContain(`<span style="color:#B6424C">${mentionSpan("2", "Alex Doe")}</span>`);
    expect(preview.body.content).toEndWith(TEAMS_CODE);
  });

  test("an <a> href is not text, and --html sends no link preview request", async () => {
    serve({ [SEARCH]: { status: 200, json: searchV2 } });
    const preview = await previewHtml(`<p>@{Alex Doe} <a href="https://x.test/?u=@{Alex Doerr}">docs</a></p>`);
    expect(preview.body.content).toContain('href="https://x.test/?u=@{Alex Doerr}"');
    expect(mentionEntries(preview)).toHaveLength(1);
    expect(preview.body.properties.links).toBe("[]");
    expect(fake!.requests.some(item => item.path.includes("urlp"))).toBe(false);
  });
});
