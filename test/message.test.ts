import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { htmlToText, normalizeMessage, sortByTime } from "../src/message.ts";
import { history, parseAfterTs } from "../src/commands/history.ts";
import { resetLocalState, installFakeTeams, seedProfile, type Route } from "./fake-teams.ts";
import messages from "./fixtures/messages.json";

const CONV = "19:dev@thread.tacv2";
const CONV_PATH = "GET /api/chatsvc/amer/v1/users/ME/conversations/19:dev@thread.tacv2/messages";
const ME = "8:orgid:11111111-1111-1111-1111-111111111111";
const context = { conversationId: CONV, myMri: ME, nameFor: () => undefined };
const normalizedAll = () => (messages.messages as any[]).map(raw => normalizeMessage(raw, context));

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
});

describe("htmlToText", () => {
  test("a mention span becomes @Name and the readonly wrapper goes", () => {
    expect(htmlToText("<p><readonly class=\"x\" itemtype=\"http://schema.skype.com/Mention\"><span itemtype=\"http://schema.skype.com/Mention\" itemid=\"0\">Alex Doe</span></readonly>we</p>")).toBe("@Alex Doewe");
  });

  test("the code block marker paragraph is not text, whether or not the server added a newline inside it", () => {
    const marker = (inner: string) => `<p itemtype="http://schema.skype.com/CodeBlockEditor" id="codeBlockEditor-1">${inner}</p>`;
    const sent = `<p>before</p>${marker("&nbsp;")}<pre class="language-python"><code>a=1</code></pre>`;
    const stored = `<p>before</p>\r\n${marker("\r\n&nbsp;")}\r\n<pre class="language-python"><code>a=1</code></pre>`;
    expect(htmlToText(sent)).toBe(htmlToText(stored));
    expect(htmlToText(sent)).toBe("before\na=1");
  });

  test("line breaks, paragraphs, and list items", () => {
    expect(htmlToText("<p>one</p><p>two<br>three</p><ul><li>a</li><li>b</li></ul>")).toBe("one\ntwo\nthree\n- a- b");
  });

  test("a link with the same text as its href shows once", () => {
    expect(htmlToText("<a href=\"https://example.test/x\">https://example.test/x</a>")).toBe("https://example.test/x");
    expect(htmlToText("<a href=\"https://example.test/x\">docs</a>")).toBe("docs (https://example.test/x)");
  });

  test("an image becomes [image] and a quote becomes > lines", () => {
    expect(htmlToText("<img src=\"x.png\">")).toBe("[image]");
    expect(htmlToText("<blockquote itemtype=\"http://schema.skype.com/Reply\"><p>first</p><p>second</p></blockquote><p>reply</p>")).toBe("> first second\nreply");
  });

  test("entities decode once, and &amp; decodes last", () => {
    expect(htmlToText("<p>a &amp;lt; b &lt;c&gt; &quot;d&quot; &#39;e&#39;&nbsp;f</p>")).toBe("a &lt; b <c> \"d\" 'e' f");
  });

  test("CRLF between paragraphs becomes one line break", () => {
    expect(htmlToText("<p>line one</p>\r\n<p>&nbsp;</p>\r\n<p>line three</p>")).toBe("line one\n \nline three");
  });

  test("blank runs collapse to one blank line", () => {
    expect(htmlToText("<p>a</p><p></p><p></p><p></p><p>b</p>")).toBe("a\n\nb");
  });
});

describe("normalizeMessage", () => {
  test("keeps text and rich messages and skips system, call, and deleted", () => {
    const kept = normalizedAll().filter(Boolean);
    expect(kept.map(message => message!.ts).sort()).toEqual(["1791988883194", "1791988884000", "1791988885000", "1791988889000"]);
  });

  test("the id joins the conversation and the message id", () => {
    const hello = normalizedAll().find(message => message?.ts === "1791988883194")!;
    expect(hello.id).toBe("19:dev@thread.tacv2#1791988883194");
    expect(hello.user).toBe("Test User");
    expect(hello.user_id).toBe(ME);
    expect(hello.created).toBe(new Date(1791988883194).toISOString());
    expect(hello.reactions).toEqual([]);
  });

  test("a mention message reports the mention and mentioned", () => {
    const mention = normalizedAll().find(message => message?.ts === "1791988884000")!;
    expect(mention.text).toBe("@Test Userwe");
    expect(mention.mentions).toEqual(["Test User"]);
    expect(mention.mentioned).toBe(true);
  });

  test("an edited message gets an edited time", () => {
    const edited = normalizedAll().find(message => message?.ts === "1791988885000")!;
    expect(edited.edited).toBe(new Date(1791988886000).toISOString());
  });

  test("a quote and link message keeps the quote lines in text", () => {
    const quote = normalizedAll().find(message => message?.ts === "1791988889000")!;
    expect(quote.text).toBe("> first line second line\nReply & https://example.test/x[image]");
  });

  test("content adds card and file text without duplicates", () => {
    const message = normalizeMessage({
      id: "1791988883194",
      messagetype: "RichText/Html",
      content: "<p>see file</p>",
      imdisplayname: "Test User",
      conversationid: CONV,
      properties: {
        files: JSON.stringify([{ fileName: "plan.docx", objectUrl: "https://example.test/plan" }, { nope: true }]),
        cards: JSON.stringify([{ content: { text: "see file", title: "Card title" } }]),
      },
    }, context)!;
    expect(message.content).toBe("see file\nplan.docx\nCard title");
    expect(message.files).toEqual([{ name: "plan.docx", url: "https://example.test/plan" }]);
  });

  test("malformed JSON properties do not throw", () => {
    const message = normalizeMessage({
      id: "1791988883194", messagetype: "Text", content: "hi", conversationid: CONV,
      properties: { mentions: "{not json", files: "[oops", cards: "nope" },
    }, context)!;
    expect(message.mentions).toEqual([]);
    expect(message.files).toEqual([]);
    expect(message.content).toBe("hi");
  });

  test("messages sort by numeric id", () => {
    const sorted = sortByTime(normalizedAll().filter(Boolean) as any[]);
    expect(sorted.map(message => message.ts)).toEqual(["1791988883194", "1791988884000", "1791988885000", "1791988889000"]);
  });
});

describe("history command", () => {
  test("default reads the last 20 and prints the label and count", async () => {
    serve({ [CONV_PATH]: { status: 200, json: messages } });
    await history(CONV, { json: false, content: false });
    expect(logs[0]).toBe(`${CONV} (4 messages)`);
    expect(fake!.requests[0]!.query).toEqual({ view: "msnp24Equivalent|supportsMessageProperties", pageSize: "200", startTime: "1" });
  });

  test("--top keeps the newest N in ascending order", async () => {
    serve({ [CONV_PATH]: { status: 200, json: messages } });
    await history(CONV, { json: true, content: false, top: 2 });
    const out = JSON.parse(logs.join("\n"));
    expect(out.messages.map((message: any) => message.ts)).toEqual(["1791988885000", "1791988889000"]);
    expect(out.channel_id).toBe(CONV);
  });

  test("--top rejects values outside 1 to 1000", async () => {
    await expect(history(CONV, { json: false, content: false, top: 1001 })).rejects.toThrow("--top must be a whole number from 1 to 1000.");
  });

  test("--top and --after-ts together fail", async () => {
    await expect(history(CONV, { json: false, content: false, top: 3, afterTs: "1791988883000" })).rejects.toThrow("Use either --top or --after-ts, not both.");
  });

  test("--after-ts keeps messages newer than the time and prints no window warning", async () => {
    serve({ [CONV_PATH]: { status: 200, json: messages } });
    await history(CONV, { json: true, content: false, afterTs: "1791988884500" });
    const out = JSON.parse(logs.join("\n"));
    expect(out.messages.map((message: any) => message.ts)).toEqual(["1791988885000", "1791988889000"]);
    expect(errors).toEqual([]);
    await history(CONV, { json: true, content: false, afterTs: "1791988000000" });
    expect(errors).toEqual([]);
  });

  test("--after-ts accepts ISO text and 12 or more digits", () => {
    expect(parseAfterTs("1791988884000")).toBe(1791988884000);
    expect(parseAfterTs("2026-10-14T14:41:24Z")).toBe(Date.parse("2026-10-14T14:41:24Z"));
    expect(() => parseAfterTs("yesterday-ish")).toThrow("--after-ts must be an ISO datetime");
  });

  test("--thread builds the path with an encoded messageid", async () => {
    serve({ [`GET /api/chatsvc/amer/v1/users/ME/conversations/${CONV};messageid=1791988883194/messages`]: { status: 200, json: messages } });
    await history(CONV, { json: false, content: false, thread: "1791988883194" });
    expect(fake!.requests[0]!.path).toBe(`/api/chatsvc/amer/v1/users/ME/conversations/${CONV};messageid=1791988883194/messages`);
  });

  test("--thread works only with channels", async () => {
    await expect(history("48:notes", { json: false, content: false, thread: "1791988883194" })).rejects.toThrow("--thread works only with channels.");
  });

  test("a 403 gives the access message", async () => {
    serve({ [CONV_PATH]: { status: 403, json: { errorCode: 403, message: "Forbidden" } } });
    await expect(history(CONV, { json: false, content: false })).rejects.toThrow(
      "You are not a member of this conversation, or the chat token lacks access.",
    );
  });

  test("the human output shows text, and --content shows the full content", async () => {
    serve({ [CONV_PATH]: { status: 200, json: messages } });
    await history(CONV, { json: false, content: true, top: 4 });
    expect(logs.join("\n")).toContain("  @Test Userwe");
    expect(logs.join("\n")).toContain("(id: 19:dev@thread.tacv2#1791988884000)");
  });
});
