import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { cardHash } from "../src/card.ts";
import { setWaitSleepForTests, WAIT_POLLS } from "../src/card-ref.ts";
import { cardSend, cardUpdate } from "../src/commands/card.ts";
import { saveCardRouterUrl } from "../src/config.ts";
import { saveConversations } from "../src/conversations.ts";
import { listRecords } from "../src/ledger.ts";
import { redact } from "../src/tokens.ts";
import { OID, TID } from "./jwt.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";

const CHANNEL = "19:fake@thread.tacv2";
const GROUP = "33333333-3333-3333-3333-333333333333";
const BOT_CHAT = "19:aaaa_bbbb@unq.gbl.spaces";
const OTHER_CHAT = "19:cccc@thread.v2";
const ROOT = "1791990000000";
const REPLY = "1791990000555";
const POSTED = "1791990001234";
const SECRET = "SECRETSIG-0123456789";
const URL_PATH = "/workflows/abc123/triggers/manual/paths/invoke";
const ROUTER = `https://prod-00.westus.logic.azure.com:443${URL_PATH}?api-version=2016-06-01&sig=${SECRET}`;
const ROUTE = `POST ${URL_PATH}`;
const MESSAGES = "GET /api/chatsvc/amer/v1/users/ME/conversations/{id}/messages";
const CARD = { type: "AdaptiveCard", version: "1.5", body: [{ type: "TextBlock", text: "Build status", wrap: true }, { type: "Container", items: [] }] };
const CARD_TYPE = "application/vnd.microsoft.card.adaptive";
// A 1x1 PNG header and a GIF header. Only the first bytes are read.
const PNG = Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13]), Buffer.from("IHDR"), Buffer.from([0, 0, 0, 1, 0, 0, 0, 1]), Buffer.alloc(20)]);
const SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>';

let dir: string;
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
let sleeps: number[] = [];
const originalLog = console.log;
const originalError = console.error;

function file(content: unknown = CARD, name = "card.json"): string {
  const path = join(dir, name);
  writeFileSync(path, typeof content === "string" ? content : JSON.stringify(content));
  return path;
}
function serve(routes: Record<string, Route> = {}) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
  return fake;
}
const posts = () => fake!.requests.filter(request => request.host.endsWith(".logic.azure.com"));
const reads = () => fake!.requests.filter(request => request.method === "GET");
const write = { json: false, allowWrite: true, noWait: true } as const;
const from = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;
const message = (id: string, mri: string, card: unknown, extra: Record<string, unknown> = {}) => ({
  id, from: from(mri), imdisplayname: "someone",
  properties: { cards: JSON.stringify([{ cardClientId: "c1", contentType: CARD_TYPE, content: card }]) }, ...extra,
});
async function failure(call: Promise<unknown>): Promise<Error> {
  try { await call; } catch (error) { return error as Error; }
  throw new Error("The call did not fail.");
}

beforeEach(async () => {
  dir = mkdtempSync(join(tmpdir(), "teams-cli-card-update-"));
  logs = [];
  sleeps = [];
  console.log = (value?: unknown) => { logs.push(String(value)); };
  console.error = () => {};
  process.env.TEAMS_CARD_ROUTER = ROUTER;
  setWaitSleepForTests(async ms => { sleeps.push(ms); });
  await resetLocalState();
  await seedProfile(["chat", "csa", "spaces"]);
  await saveConversations(TID, {
    version: 1, updatedAt: "", chats: [
      { id: BOT_CHAT, title: "", type: "chat", isOneOnOne: true, members: [`8:orgid:${OID}`, "28:bot-guid"] },
      { id: OTHER_CHAT, title: "Group", type: "chat", isOneOnOne: false, members: [] },
    ],
    teams: [{ id: "t1", name: "Eng", groupId: GROUP, channels: [{ id: CHANNEL, name: "General" }] }],
  });
});

afterEach(() => {
  console.log = originalLog;
  console.error = originalError;
  delete process.env.TEAMS_CARD_ROUTER;
  setWaitSleepForTests(ms => Bun.sleep(ms));
  fake?.uninstall();
  fake = undefined;
  process.exitCode = 0;
});

/** The Workflows chat is found from its bot messages, so a route answers for it. */
const botMessages = (messages: unknown[] = []): Route => request => request.path.includes(BOT_CHAT)
  ? { status: 200, json: { messages: [{ id: "1", from: from("8:bot-guid"), imdisplayname: "Workflows", content: "hi" }, ...messages] } }
  : { status: 200, json: { messages: [] } };

describe("card update", () => {
  test("a channel post sends mode update with channel and the bare message id", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardUpdate(`${CHANNEL}#${POSTED}`, file(), write);
    expect(posts()).toHaveLength(1);
    expect(posts()[0]!.body).toEqual({
      type: "message",
      target: { mode: "update", channel: CHANNEL, messageId: POSTED },
      attachments: [{ contentType: CARD_TYPE, contentUrl: null, content: CARD }],
    });
    expect(logs.join("\n")).toContain("Accepted by Workflows");
  });

  test("a thread reply sends messageId as <root>/replies/<reply>, never the bare reply id", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardUpdate(`${CHANNEL};messageid=${ROOT}#${REPLY}`, file(), write);
    const target = (posts()[0]!.body as any).target;
    expect(target).toEqual({ mode: "update", channel: CHANNEL, messageId: `${ROOT}/replies/${REPLY}` });
  });

  test("a draft ref sends mode update-draft with the message id and no channel", async () => {
    serve({ [ROUTE]: { status: 202 }, [MESSAGES]: botMessages() });
    await cardUpdate(`draft#${POSTED}`, file(), write);
    expect((posts()[0]!.body as any).target).toEqual({ mode: "update-draft", messageId: POSTED });
  });

  test("a ref in the Workflows chat is a draft, and a ref in another chat is refused", async () => {
    serve({ [ROUTE]: { status: 202 }, [MESSAGES]: botMessages() });
    await cardUpdate(`${BOT_CHAT}#${POSTED}`, file(), write);
    expect((posts()[0]!.body as any).target.mode).toBe("update-draft");
    const error = await failure(cardUpdate(`${OTHER_CHAT}#${POSTED}`, file(), write));
    expect(error.message).toContain("is not a channel or your Workflows chat");
    expect(posts()).toHaveLength(1);
  });

  test("the preview says what changes and makes no request to the router", async () => {
    serve();
    await cardUpdate(`${CHANNEL};messageid=${ROOT}#${REPLY}`, file(), { json: false, allowWrite: false });
    expect(posts()).toHaveLength(0);
    expect(fake!.requests.filter(request => request.method !== "GET")).toHaveLength(0);
    const text = logs.join("\n");
    expect(text).toContain(`would replace the card ${ROOT}/replies/${REPLY} in Eng/General`);
    expect(text).toContain("title: Build status");
    expect(text).not.toContain(SECRET);
  });

  test("a ref that is not a posted card fails before any request", async () => {
    serve();
    const error = await failure(cardUpdate("not-a-ref", file(), write));
    expect(error.message).toContain("is not the ref of a posted card");
    expect(fake!.requests).toHaveLength(0);
  });

  test("the ledger keys the update by target and card hash", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardUpdate(`${CHANNEL}#${POSTED}`, file(), write);
    await expect(cardUpdate(`${CHANNEL}#${POSTED}`, file(), write)).rejects.toThrow("Identical message already sent");
    expect(posts()).toHaveLength(1);
  });

  test("the same card on another target, and another card on the same target, are separate updates", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardUpdate(`${CHANNEL}#${POSTED}`, file(), write);
    await cardUpdate(`${CHANNEL}#${POSTED}1`, file(), write);
    await cardUpdate(`${CHANNEL}#${POSTED}`, file({ ...CARD, version: "1.4" }, "b.json"), write);
    expect(posts()).toHaveLength(3);
  });

  test("a 500 is uncertain and is not retried", async () => {
    serve({ [ROUTE]: { status: 500 } });
    await expect(cardUpdate(`${CHANNEL}#${POSTED}`, file(), write)).rejects.toThrow("Outcome unknown");
    expect(posts()).toHaveLength(1);
  });

  test("the stored router URL is used when TEAMS_CARD_ROUTER is not set, and the environment wins when it is", async () => {
    delete process.env.TEAMS_CARD_ROUTER;
    await saveCardRouterUrl(`https://stored.logic.azure.com${URL_PATH}?sig=stored`);
    serve({ [ROUTE]: { status: 202 } });
    await cardUpdate(`${CHANNEL}#${POSTED}`, file(), write);
    expect(posts()[0]!.host).toBe("stored.logic.azure.com");
    process.env.TEAMS_CARD_ROUTER = ROUTER;
    await cardUpdate(`${CHANNEL}#${POSTED}1`, file(), write);
    expect(posts()[1]!.host).toBe("prod-00.westus.logic.azure.com");
  });
});

describe("card send --wait", () => {
  const waiting = { ...write, wait: true } as const;

  test("finds the new post by card title, prints its ref, and keeps it in the ledger record", async () => {
    const now = Date.now();
    const mine = `8:orgid:${OID}`;
    // Teams rewrites a stored card (key order, enum case), so the title is what matches. The other messages must not match.
    const reordered = { body: [{ text: "Build status", weight: "bolder", type: "TextBlock" }], version: "1.5", type: "AdaptiveCard" };
    serve({
      [ROUTE]: { status: 202 },
      [MESSAGES]: { status: 200, json: { messages: [
        message(String(now + 10), mine, reordered),
        message(String(now + 20), mine, { ...CARD, body: [{ type: "TextBlock", text: "Another card" }] }),
        message(String(now + 30), "8:orgid:someone-else", CARD),
        message(String(now - 3_600_000), mine, CARD),
      ] } },
    });
    await cardSend(CHANNEL, file(), waiting);
    expect(logs).toContain(`Ref: ${CHANNEL}#${now + 10}`);
    expect(sleeps).toHaveLength(0);
    const [record] = await listRecords(TID);
    expect(record!.state).toBe("sent");
    expect(record!.id).toBe(`${CHANNEL}#${now + 10}`);
    // The ref works as an update target.
    await cardUpdate(logs.find(line => line.startsWith("Ref: "))!.slice(5), file({ ...CARD, version: "1.4" }, "b.json"), write);
    expect((posts()[1]!.body as any).target).toEqual({ mode: "update", channel: CHANNEL, messageId: String(now + 10) });
  });

  test("a thread reply is read from the thread and its ref has the thread form", async () => {
    const now = Date.now();
    serve({
      [ROUTE]: { status: 202 },
      [MESSAGES]: request => ({ status: 200, json: { messages: request.path.includes(`;messageid=${ROOT}`) ? [message(String(now + 5), `8:orgid:${OID}`, CARD)] : [] } }),
    });
    await cardSend(CHANNEL, file(), { ...waiting, thread: ROOT });
    expect(logs).toContain(`Ref: ${CHANNEL};messageid=${ROOT}#${now + 5}`);
  });

  test("a draft is matched in the Workflows chat, from the bot and not from you", async () => {
    const now = Date.now();
    serve({
      [ROUTE]: { status: 202 },
      [MESSAGES]: botMessages([message(String(now + 7), "8:bot-guid", CARD), message(String(now + 9), `8:orgid:${OID}`, CARD)]),
    });
    await cardSend(CHANNEL, file(), { ...waiting, draft: true });
    expect(logs).toContain(`Ref: ${BOT_CHAT}#${now + 7}`);
  });

  test("a card that never shows ends after the poll count, says so, and still counts as sent", async () => {
    serve({ [ROUTE]: { status: 202 }, [MESSAGES]: { status: 200, json: { messages: [message(String(Date.now()), `8:orgid:${OID}`, { ...CARD, body: [{ type: "TextBlock", text: "Another card" }] })] } } });
    await cardSend(CHANNEL, file(), waiting);
    expect(sleeps).toHaveLength(WAIT_POLLS - 1);
    expect(sleeps.reduce((sum, ms) => sum + ms, 0)).toBe(20_000);
    expect(logs.join("\n")).toContain("did not show within 20 s");
    expect(logs.join("\n")).not.toContain("Ref:");
    expect(posts()).toHaveLength(1);
    expect((await listRecords(TID))[0]!.state).toBe("sent");
  });

  test("a Workflows bot card is read from the Swift payload in the message content", async () => {
    const now = Date.now();
    const swift = Buffer.from(JSON.stringify({ attachments: [{ content: { type: "AdaptiveCard", body: [{ text: "Build status", type: "TextBlock" }] } }] })).toString("base64");
    serve({
      [ROUTE]: { status: 202 },
      [MESSAGES]: botMessages([{ ...message(String(now + 5), "8:bot-guid", CARD), messagetype: "RichText/Media_Card", content: `<URIObject type="SWIFT.1"><Swift b64="${swift}"/></URIObject>`, properties: {} }]),
    });
    await cardSend(CHANNEL, file(), { ...waiting, draft: true });
    expect(logs).toContain(`Ref: ${BOT_CHAT}#${now + 5}`);
  });

  test("without --wait it reads nothing and says how to find the ref", async () => {
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, file(), write);
    expect(reads()).toHaveLength(0);
    expect(logs.join("\n")).toContain("history");
    expect(logs.join("\n")).toContain("--json");
  });

  test("the hash ignores key order", () => {
    expect(cardHash({ a: 1, b: { c: 2, d: [1, { e: 1, f: 2 }] } })).toBe(cardHash({ b: { d: [1, { f: 2, e: 1 }], c: 2 }, a: 1 }));
    expect(cardHash({ a: 1 })).not.toBe(cardHash({ a: 2 }));
  });
});

describe("images in cards", () => {
  const imageCard = (url: string) => ({ ...CARD, body: [{ type: "Image", url }] });

  test("a file: url becomes a data URI relative to the card file, and the preview lists it", async () => {
    writeFileSync(join(dir, "pic.png"), PNG);
    serve({ [ROUTE]: { status: 202 } });
    await cardSend(CHANNEL, file(imageCard("file:pic.png")), { json: false, allowWrite: false });
    expect(logs.join("\n")).toContain(`image: pic.png, ${PNG.length} bytes, png`);
    await cardSend(CHANNEL, file(imageCard("file:pic.png")), write);
    expect((posts()[0]!.body as any).attachments[0].content.body[0].url).toBe(`data:image/png;base64,${PNG.toString("base64")}`);
  });

  test("backgroundImage as a string and as an object are both replaced, and a https url is left alone", async () => {
    writeFileSync(join(dir, "pic.png"), PNG);
    serve({ [ROUTE]: { status: 202 } });
    const card = { ...CARD, backgroundImage: "file:pic.png", body: [
      { type: "Container", backgroundImage: { url: "file:pic.png", fillMode: "cover" }, items: [{ type: "Image", url: "https://example.test/a.png" }] },
    ] };
    await cardSend(CHANNEL, file(card), write);
    const sent = (posts()[0]!.body as any).attachments[0].content;
    expect(sent.backgroundImage).toStartWith("data:image/png;base64,");
    expect(sent.body[0].backgroundImage).toEqual({ url: `data:image/png;base64,${PNG.toString("base64")}`, fillMode: "cover" });
    expect(sent.body[0].items[0].url).toBe("https://example.test/a.png");
  });

  test("SVG is refused with the render advice, by extension and by content", async () => {
    writeFileSync(join(dir, "chart.svg"), SVG);
    writeFileSync(join(dir, "chart.png"), SVG);
    serve();
    for (const name of ["chart.svg", "chart.png"]) {
      const error = await failure(cardSend(CHANNEL, file(imageCard(`file:${name}`)), write));
      expect(error.message).toContain("Teams does not show SVG in cards. Render it to PNG first (rsvg-convert, or a headless browser for Mermaid)");
    }
    expect(fake!.requests).toHaveLength(0);
  });

  test("a payload over the limit fails with its size, in the preview and in a send", async () => {
    writeFileSync(join(dir, "big.png"), Buffer.concat([PNG, Buffer.alloc(200 * 1024)]));
    serve();
    const error = await failure(cardSend(CHANNEL, file(imageCard("file:big.png")), { json: false, allowWrite: false }));
    expect(error.message).toMatch(/The payload is \d+ bytes\. The card router accepts 262144 bytes at most/);
    await expect(cardSend(CHANNEL, file(imageCard("file:big.png")), write)).rejects.toThrow("Make the card smaller");
    expect(posts()).toHaveLength(0);
  });

  test("a missing file and a file of another type fail before any request", async () => {
    writeFileSync(join(dir, "note.txt"), "plain text");
    serve();
    expect((await failure(cardSend(CHANNEL, file(imageCard("file:none.png")), write))).message).toContain("Cannot read the card image file");
    expect((await failure(cardSend(CHANNEL, file(imageCard("file:note.txt")), write))).message).toContain("is not a PNG, JPEG, or GIF");
    expect(fake!.requests).toHaveLength(0);
  });

  test("an error never shows the router URL", async () => {
    serve({ [ROUTE]: { status: 400, text: `bad ${SECRET}` } });
    const error = await failure(cardSend(CHANNEL, file(), write));
    expect(redact(error.message)).not.toContain(SECRET.repeat(2));
    expect(error.message).not.toContain(URL_PATH);
  });
});
