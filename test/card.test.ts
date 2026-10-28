import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { cardFromProperties, cardMarker, cardOutline } from "../src/card.ts";
import { ACCEPTED, cardSend } from "../src/commands/card.ts";
import { history } from "../src/commands/history.ts";
import { MAX_CARD_BYTES } from "../src/limits.ts";
import { normalizeMessage } from "../src/message.ts";
import { redact } from "../src/tokens.ts";
import { runCli } from "./cli.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";

const CHANNEL = "19:fake@thread.tacv2";
const CHAT = "19:fake@thread.v2";
const ROOT = "1791990000000";
const SECRET = "SECRETSIG-0123456789";
const URL_PATH = "/workflows/abc123/triggers/manual/paths/invoke";
const ROUTER = `https://prod-00.westus.logic.azure.com:443${URL_PATH}?api-version=2016-06-01&sig=${SECRET}`;
const ROUTE = `POST ${URL_PATH}`;

// The shape of a card that the live flow posted (2026-10-19). Only the structure is real. Every text is invented.
const CARD = {
  type: "AdaptiveCard",
  version: "1.5",
  body: [
    { type: "Container", items: [{ type: "TextBlock", text: "Build status", weight: "bolder" }, { type: "TextBlock", text: "All green", isSubtle: true }] },
    { type: "ColumnSet", columns: [{ type: "Column", items: [{ type: "TextBlock", text: "Tests" }] }, { type: "Column", items: [{ type: "TextBlock", text: "Lint" }] }] },
    { type: "ActionSet", actions: [{ type: "Action.ToggleVisibility", title: "Show details" }, { type: "Action.OpenUrl", title: "Open run", url: "https://example.test/run" }] },
  ],
};

let dir: string;
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
let errors: string[] = [];
const originalLog = console.log;
const originalError = console.error;

function cardFile(content: unknown = CARD, name = "card.json"): string {
  const path = join(dir, name);
  writeFileSync(path, typeof content === "string" ? content : JSON.stringify(content));
  return path;
}
function serve(routes: Record<string, Route> = {}) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
  return fake;
}
const posts = () => fake!.requests.filter(request => /\.(logic\.azure|powerplatform)\.com$/.test(request.host));
const send = { json: false, allowWrite: true } as const;
/** Runs a call that must fail and returns its error. */
async function failure(call: Promise<unknown>): Promise<Error> {
  try { await call; } catch (error) { return error as Error; }
  throw new Error("The call did not fail.");
}
const output = () => [...logs, ...errors].join("\n");

beforeEach(async () => {
  dir = mkdtempSync(join(tmpdir(), "teams-cli-card-"));
  logs = [];
  errors = [];
  console.log = (value?: unknown) => { logs.push(String(value)); };
  console.error = (value?: unknown) => { errors.push(String(value)); };
  process.env.TEAMS_CARD_ROUTER = ROUTER;
  await resetLocalState();
  await seedProfile(["chat", "csa", "spaces"]);
});

afterEach(() => {
  console.log = originalLog;
  console.error = originalError;
  delete process.env.TEAMS_CARD_ROUTER;
  fake?.uninstall();
  fake = undefined;
  process.exitCode = 0;
});

describe("card send preview", () => {
  test("makes no request and shows the target, size, outline, and reminder", async () => {
    serve();
    await cardSend(CHANNEL, cardFile(), { json: false, allowWrite: false, thread: `${CHANNEL}#${ROOT}` });
    expect(posts()).toHaveLength(0);
    const text = logs.join("\n");
    expect(text).toContain(`a reply in thread ${ROOT} of`);
    expect(text).toContain("version: 1.5");
    expect(text).toContain(`${Buffer.byteLength(JSON.stringify(CARD))} bytes`);
    expect(text).toContain("title: Build status");
    expect(text).toContain("TextBlock 4");
    expect(text).toContain("Container 1");
    expect(text).toContain("actions: Show details, Open run");
    expect(text).toContain("--draft --allow-write first");
    expect(text).not.toContain(SECRET);
  });

  test("a preview with no thread says it is a new post, and the JSON preview carries the outline", async () => {
    serve();
    await cardSend(CHANNEL, cardFile(), { json: true, allowWrite: false });
    const preview = JSON.parse(logs[0]!);
    expect(preview.mode).toBe("post");
    expect(preview.dry_run).toBe(true);
    expect(preview.outline).toEqual(cardOutline(CARD));
    expect(posts()).toHaveLength(0);
  });

  test("a preview needs no webhook variable", async () => {
    delete process.env.TEAMS_CARD_ROUTER;
    serve();
    await cardSend(CHANNEL, cardFile(), { json: false, allowWrite: false });
    expect(logs.join("\n")).toContain("a new post in");
  });
});

describe("card send request bodies", () => {
  test("--draft sends mode draft with no channel, whatever the channel", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(), { ...send, draft: true });
    expect(posts()).toHaveLength(1);
    expect(posts()[0]!.body).toEqual({
      type: "message",
      target: { mode: "draft" },
      attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", contentUrl: null, content: CARD }],
    });
    expect(logs.join("\n")).toContain(ACCEPTED);
  });

  test("a post sends mode post with the channel id", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(), send);
    expect((posts()[0]!.body as any).target).toEqual({ mode: "post", channel: CHANNEL });
  });

  test("--thread with a full ref sends mode reply with the root id", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(), { ...send, thread: `${CHANNEL}#${ROOT}` });
    expect((posts()[0]!.body as any).target).toEqual({ mode: "reply", channel: CHANNEL, replyTo: ROOT });
  });

  test("a bare root id works, and a bad root fails before any request", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(), { ...send, thread: ROOT });
    expect((posts()[0]!.body as any).target.replyTo).toBe(ROOT);
    await expect(cardSend(CHANNEL, cardFile(), { ...send, thread: "bad" })).rejects.toThrow("--thread needs the root message id");
    expect(posts()).toHaveLength(1);
  });

  test("a full message payload file is accepted and sent as the same card", async () => {
    serve({ [ROUTE]: { status: 202 } });
    const payload = { type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", contentUrl: null, content: CARD }] };
    await cardSend(CHANNEL, cardFile(payload), send);
    expect((posts()[0]!.body as any).attachments[0].content).toEqual(CARD);
  });

  test("a chat is refused for a post, and accepted for a draft", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await expect(cardSend(CHAT, cardFile(), send)).rejects.toThrow("Cards to chats go through --draft only for now.");
    expect(posts()).toHaveLength(0);
    await cardSend(CHAT, cardFile(), { ...send, draft: true });
    expect(posts()).toHaveLength(1);
  });
});

describe("card file checks", () => {
  test.each([
    ["not JSON", "{nope", "is not valid JSON"],
    ["a wrong type", { type: "Other", version: "1.5", body: [] }, '"type": "AdaptiveCard"'],
    ["no version", { type: "AdaptiveCard", body: [] }, "needs a \"version\""],
    ["a body that is not an array", { type: "AdaptiveCard", version: "1.5", body: {} }, "\"body\" that is an array"],
    ["a payload with two attachments", { type: "message", attachments: [{}, {}] }, "exactly one attachment"],
  ])("%s fails before any request", async (_name, content, message) => {
    serve();
    await expect(cardSend(CHANNEL, cardFile(content as any), send)).rejects.toThrow(message as string);
    expect(posts()).toHaveLength(0);
  });

  test("a missing file fails with its path", async () => {
    serve();
    await expect(cardSend(CHANNEL, join(dir, "none.json"), send)).rejects.toThrow("Cannot read the card file");
  });

  test("a payload over the limit fails in a preview and in a send", async () => {
    serve();
    const big = { ...CARD, body: [{ type: "TextBlock", text: "x".repeat(MAX_CARD_BYTES) }] };
    await expect(cardSend(CHANNEL, cardFile(big), { json: false, allowWrite: false })).rejects.toThrow(`${MAX_CARD_BYTES} bytes at most`);
    await expect(cardSend(CHANNEL, cardFile(big), send)).rejects.toThrow("Make the card smaller");
    expect(posts()).toHaveLength(0);
  });
});

describe("the webhook variable", () => {
  test("a missing variable names TEAMS_CARD_ROUTER and the flow", async () => {
    delete process.env.TEAMS_CARD_ROUTER;
    serve();
    await expect(cardSend(CHANNEL, cardFile(), send)).rejects.toThrow(/TEAMS_CARD_ROUTER is not set.*card router/);
    expect(posts()).toHaveLength(0);
  });

  test.each([
    ["http", `http://prod-00.logic.azure.com${URL_PATH}?sig=${SECRET}`],
    ["another host", `https://evil.example.com${URL_PATH}?sig=${SECRET}`],
    ["a host that only ends in the name", `https://logic.azure.com.evil.test${URL_PATH}?sig=${SECRET}`],
    ["text that is not a URL", `not a url ${SECRET}`],
  ])("%s is refused, with no request and no URL in the error", async (_name, value) => {
    process.env.TEAMS_CARD_ROUTER = value;
    serve();
    const error = await failure(cardSend(CHANNEL, cardFile(), send));
    expect(error.message).toContain("is not accepted");
    expect(error.message).not.toContain(SECRET);
    expect(error.message).not.toContain("evil");
    expect(fake!.requests).toHaveLength(0);
  });

  test("a powerplatform.com host is accepted", async () => {
    process.env.TEAMS_CARD_ROUTER = `https://abc.environment.api.powerplatform.com${URL_PATH}?sig=${SECRET}`;
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(), send);
    expect(posts()).toHaveLength(1);
  });
});

describe("router answers", () => {
  test("202 is accepted and prints the allowlist sentence", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(), send);
    expect(logs[0]).toBe(ACCEPTED);
    expect(logs[1]).toContain("To update this card later, find its ref with --wait");
    expect(ACCEPTED).toContain("A channel that is not on the flow's allowlist is refused in the flow.");
  });

  test("400 is an error with the status and at most 200 characters of the body, and the send can be tried again", async () => {
    serve({ [ROUTE]: [{ status: 400, text: "bad ".repeat(100) }, { status: 202 }] });
    const error = await failure(cardSend(CHANNEL, cardFile(), send));
    expect(error.message).toStartWith("The card router refused the card: status 400 bad bad");
    expect(error.message.length).toBeLessThan("The card router refused the card: status 400 ".length + 201);
    // A definite failure leaves no ledger record, so the same card can go again.
    await cardSend(CHANNEL, cardFile(), send);
    expect(posts()).toHaveLength(2);
  });

  test("500 is uncertain, is not retried, and blocks a repeat", async () => {
    serve({ [ROUTE]: { status: 500, text: "boom" } });
    await expect(cardSend(CHANNEL, cardFile(), send)).rejects.toThrow("Outcome unknown");
    expect(posts()).toHaveLength(1);
    await expect(cardSend(CHANNEL, cardFile(), send)).rejects.toThrow("may have posted");
    expect(posts()).toHaveLength(1);
  });

  test("408 and a network error are uncertain", async () => {
    serve({ [ROUTE]: { status: 408 } });
    await expect(cardSend(CHANNEL, cardFile(), send)).rejects.toThrow("status 408");
    serve({ [ROUTE]: "network" });
    const error = await failure(cardSend(CHANNEL, cardFile({ ...CARD, version: "1.4" }, "other.json"), send));
    expect(error.message).toContain("Outcome unknown");
    expect(error.message).toContain("a network error");
    expect(posts()).toHaveLength(1);
  });
});

describe("the send-once ledger", () => {
  test("refuses an identical card to the same target", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(), send);
    await expect(cardSend(CHANNEL, cardFile(), send)).rejects.toThrow("Identical message already sent");
    expect(posts()).toHaveLength(1);
  });

  test("a draft, a post, and a reply of the same card are different sends", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(), { ...send, draft: true });
    await cardSend(CHANNEL, cardFile(), send);
    await cardSend(CHANNEL, cardFile(), { ...send, thread: ROOT });
    await cardSend(CHANNEL, cardFile(), { ...send, thread: "1791990000001" });
    expect(posts()).toHaveLength(4);
  });

  test("a changed card is a different send", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(), send);
    await cardSend(CHANNEL, cardFile({ ...CARD, version: "1.4" }, "v14.json"), send);
    expect(posts()).toHaveLength(2);
  });
});

describe("the webhook URL stays secret", () => {
  test("it is in no request body, output, or error, even when the flow echoes it", async () => {
    serve({ [ROUTE]: { status: 400, text: `Invalid call to ${ROUTER} here` } });
    const error = await failure(cardSend(CHANNEL, cardFile(), send));
    expect(error.message).toContain("<card-router-url>");
    expect(error.message).not.toContain(SECRET);
    expect(JSON.stringify(posts()[0]!.body)).not.toContain(SECRET);
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, cardFile(CARD, "again.json"), { ...send, json: true });
    await cardSend(CHANNEL, cardFile(), { json: false, allowWrite: false });
    expect(output()).not.toContain(SECRET);
    expect(output()).not.toContain("logic.azure.com");
  });

  test("redact() removes the exact value and any Workflows URL", () => {
    expect(redact(`failed at ${ROUTER}.`)).toBe("failed at <card-router-url>.");
    delete process.env.TEAMS_CARD_ROUTER;
    expect(redact("see https://x.y.powerplatform.com/a?sig=zzz now")).toBe("see <card-router-url> now");
  });

  test("a real run through the CLI prints no URL on an error", async () => {
    const result = await runCli(["card", "send", CHANNEL, join(dir, "nope.json")], { env: { TEAMS_CARD_ROUTER: ROUTER } });
    expect(result.code).toBe(1);
    expect(result.stdout + result.stderr).not.toContain(SECRET);
  });
});

describe("card messages in history and context", () => {
  // The same structure as the live message: properties.cards is a JSON string of [{cardClientId, content, contentType}].
  const cardMessage = (id: number) => ({
    id: String(id), messagetype: "RichText/Html", imdisplayname: "Alex Doe", properties: {
      cards: JSON.stringify([{ cardClientId: "c1", contentType: "application/vnd.microsoft.card.adaptive", content: CARD }]),
    },
    content: '<span itemid="c1" itemscope="" itemtype="http://schema.skype.com/InputExtension"><span itemprop="cardId"></span></span>',
  });
  const context = { conversationId: CHANNEL, myMri: "me", nameFor: () => undefined };

  test("normalizeMessage reads the card from properties.cards", () => {
    const message = normalizeMessage(cardMessage(1791990000100), context)!;
    expect(message.card).toEqual({ title: "Build status", elements: { Container: 1, TextBlock: 4, ColumnSet: 1, Column: 2, ActionSet: 1, "Action.ToggleVisibility": 1, "Action.OpenUrl": 1 } });
    expect(message.text).toBe("");
    expect(cardMarker(message.card!)).toBe("[card: Build status]");
  });

  test("a card with no TextBlock is just [card]", () => {
    const info = cardFromProperties(JSON.stringify([{ content: { type: "AdaptiveCard", version: "1.5", body: [{ type: "Image", url: "https://example.test/a.png" }] } }]))!;
    expect(info.title).toBeNull();
    expect(cardMarker(info)).toBe("[card]");
    expect(cardFromProperties("[]")).toBeUndefined();
    expect(cardFromProperties("nope")).toBeUndefined();
  });

  test("history prints the marker where the empty line was, and --json adds card", async () => {
    serve({
      [`GET /api/chatsvc/amer/v1/users/ME/conversations/${CHANNEL}/messages`]: { status: 200, json: { messages: [cardMessage(1791990000100)] } },
      "POST /api/mt/amer/beta/users/fetchShortProfile": { status: 200, json: [] },
    });
    await history(CHANNEL, { json: false, content: false });
    const lines = logs.filter(line => line.startsWith("  "));
    expect(lines).toEqual(["  [card: Build status]"]);
    logs = [];
    await history(CHANNEL, { json: true, content: false });
    const parsed = JSON.parse(logs[0]!);
    expect(parsed.messages[0].card.title).toBe("Build status");
    expect(parsed.messages[0].card.elements.TextBlock).toBe(4);
  });
});
