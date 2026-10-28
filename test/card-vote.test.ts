import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { setWaitSleepForTests, WAIT_POLLS } from "../src/card-ref.ts";
import { buildTemplateCard, findTemplate } from "../src/card-templates/templates.ts";
import { cardSend, cardUpdate } from "../src/commands/card.ts";
import { cardVotes, countVotes, voteOptions } from "../src/commands/card-templates.ts";
import { saveConversations } from "../src/conversations.ts";
import { listRecords } from "../src/ledger.ts";
import { setPaceForTests } from "../src/pace.ts";
import { OID, TID } from "./jwt.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";

const CHANNEL = "19:fake@thread.tacv2";
const GROUP = "33333333-3333-3333-3333-333333333333";
const BOT_CHAT = "19:aaaa_bbbb@unq.gbl.spaces";
const POSTED = "1791990001234";
const ROUTER = "https://prod-00.westus.logic.azure.com:443/workflows/abc123/triggers/manual/paths/invoke?sig=SECRETSIG-0123456789";
const ROUTE = "POST /workflows/abc123/triggers/manual/paths/invoke";
const LIST = "GET /api/chatsvc/amer/v1/users/ME/conversations/{id}/messages";
const ONE = "GET /api/chatsvc/amer/v1/users/ME/conversations/{id}/messages/{mid}";
const PUT = "PUT /api/chatsvc/amer/v1/users/ME/conversations/{id}/messages/{mid}/properties";
const CARD_TYPE = "application/vnd.microsoft.card.adaptive";
const MINE = `8:orgid:${OID}`;
const OTHER_A = "8:orgid:00000000-0000-0000-0000-00000000000a";
const OTHER_B = "8:orgid:00000000-0000-0000-0000-00000000000b";
const OTHER_C = "8:orgid:00000000-0000-0000-0000-00000000000c";
const POLL_KEYS = ["happy_person_raising_one_hand", "handshake", "26d4_noentry"];
const ok = { status: 200, text: "" };

let dir: string;
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
let sleeps: number[] = [];
let clock = 0;
const originalLog = console.log;
const originalError = console.error;

function dataFile(name: string, data: unknown = findTemplate(name).example): { template: string; data: string } {
  const path = join(dir, `${name}.json`);
  writeFileSync(path, JSON.stringify(data));
  return { template: name, data: path };
}
function serve(routes: Record<string, Route> = {}) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
  return fake;
}
const posts = () => fake!.requests.filter(request => request.host.endsWith(".logic.azure.com"));
const puts = () => fake!.requests.filter(request => request.method === "PUT");
const writes = () => fake!.requests.filter(request => request.method !== "GET");
const putKeys = () => puts().map(request => (request.body as any).emotions.key);
const write = { json: false, allowWrite: true } as const;
const botFrom = "https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/8:bot-guid";
const myFrom = `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${MINE}`;

/** The card that a template builds, as Teams stores it on a message. */
const stored = (name: string, id: string, from = myFrom) => ({
  id, from, imdisplayname: "me",
  properties: { cards: JSON.stringify([{ cardClientId: "c1", contentType: CARD_TYPE, content: buildTemplateCard(name, findTemplate(name).example).card }]) },
});
const emotions = (users: Record<string, string[]>) =>
  JSON.stringify(Object.entries(users).map(([key, mris]) => ({ key, users: mris.map(mri => ({ mri, time: 1, value: "x" })) })));
/** A route for the list read that `--wait` polls, and one for the single-message read that `react` does. */
function posted(name: string, users: Record<string, string[]> = {}): Record<string, Route> {
  return {
    [ROUTE]: { status: 202 },
    [LIST]: request => ({ status: 200, json: { messages: request.path.includes(BOT_CHAT) ? [{ id: "1", from: botFrom, imdisplayname: "Workflows", content: "hi" }, stored(name, String(Date.now() + 10), botFrom)] : [stored(name, String(Date.now() + 10))] } }),
    [ONE]: { status: 200, json: { id: POSTED, properties: { emotions: emotions(users) } } },
    [PUT]: ok,
  };
}
async function failure(call: Promise<unknown>): Promise<Error> {
  try { await call; } catch (error) { return error as Error; }
  throw new Error("The call did not fail.");
}

beforeEach(async () => {
  dir = mkdtempSync(join(tmpdir(), "teams-cli-card-vote-"));
  logs = [];
  sleeps = [];
  clock = Date.now();
  console.log = (value?: unknown) => { logs.push(String(value)); };
  console.error = () => {};
  process.env.TEAMS_CARD_ROUTER = ROUTER;
  setWaitSleepForTests(async ms => { sleeps.push(ms); });
  // The write pacer allows 3 requests in 5 s. A clock that moves when the pacer sleeps keeps the test fast and shows the waits.
  setPaceForTests({ now: () => clock, sleep: async ms => { clock += ms; } });
  await resetLocalState();
  await seedProfile(["chat", "csa", "spaces"]);
  await saveConversations(TID, {
    version: 1, updatedAt: "", chats: [{ id: BOT_CHAT, title: "", type: "chat", isOneOnOne: true, members: [`8:orgid:${OID}`, "28:bot-guid"] }],
    teams: [{ id: "t1", name: "Eng", groupId: GROUP, channels: [{ id: CHANNEL, name: "General" }] }],
  });
});

afterEach(() => {
  console.log = originalLog;
  console.error = originalError;
  delete process.env.TEAMS_CARD_ROUTER;
  setWaitSleepForTests(ms => Bun.sleep(ms));
  setPaceForTests({});
  fake?.uninstall();
  fake = undefined;
  process.exitCode = 0;
});

describe("card send --template: the preview", () => {
  test("it names the template and lists the reactions it would add, and it sends and writes nothing", async () => {
    serve();
    await cardSend(CHANNEL, dataFile("poll"), { json: false, allowWrite: false });
    const text = logs.join("\n");
    expect(text).toContain("template: poll");
    expect(text).toContain("title: Who takes the release this week?");
    expect(text).toContain("reactions: after the post, would add 🙋 (happy_person_raising_one_hand), 🤝 (handshake), ⛔ (26d4_noentry) as your votes. --no-react skips this.");
    expect(fake!.requests).toHaveLength(0);
  });

  test("the JSON preview has the template, the reactions with their keys, and --no-react empties them", async () => {
    serve();
    await cardSend(CHANNEL, dataFile("decision"), { json: true, allowWrite: false });
    const out = JSON.parse(logs.join("\n"));
    expect(out).toMatchObject({ dry_run: true, template: "decision" });
    expect(out.reactions).toEqual([
      { emoji: "1️⃣", key: "keycapdigitone" }, { emoji: "2️⃣", key: "keycapdigittwo" }, { emoji: "3️⃣", key: "keycapdigitthree" },
    ]);
    logs = [];
    await cardSend(CHANNEL, dataFile("decision"), { json: true, allowWrite: false, noReact: true });
    expect(JSON.parse(logs.join("\n")).reactions).toBeUndefined();
  });

  test("a template with no vote has no reactions line", async () => {
    serve();
    await cardSend(CHANNEL, dataFile("standup"), { json: false, allowWrite: false });
    expect(logs.join("\n")).toContain("template: standup");
    expect(logs.join("\n")).not.toContain("reactions:");
  });

  test("bad data fails in the preview, before any request, and names the field", async () => {
    serve();
    const error = await failure(cardSend(CHANNEL, dataFile("poll", { question: "Q", options: [] }), { json: false, allowWrite: false }));
    expect(error.message).toContain('The data for template "poll" is not valid: options needs at least 2 items.');
    expect(fake!.requests).toHaveLength(0);
  });

  test("a missing data file and a data file that is not JSON fail clearly", async () => {
    serve();
    expect((await failure(cardSend(CHANNEL, { template: "poll", data: join(dir, "none.json") }, write))).message).toContain("Cannot read the data file");
    writeFileSync(join(dir, "bad.json"), "{nope");
    expect((await failure(cardSend(CHANNEL, { template: "poll", data: join(dir, "bad.json") }, write))).message).toContain("is not valid JSON");
    expect(fake!.requests).toHaveLength(0);
  });

  test("a card file over the Teams size limit is refused, in the preview and in a send", async () => {
    serve();
    const path = join(dir, "big.json");
    writeFileSync(path, JSON.stringify({ type: "AdaptiveCard", version: "1.5", body: [{ type: "TextBlock", text: "x".repeat(30 * 1024), wrap: true }] }));
    const error = await failure(cardSend(CHANNEL, path, { json: false, allowWrite: false }));
    expect(error.message).toMatch(/The card is \d+ bytes\. Teams refuses a card over 28672 bytes \(28 KB\)/);
    await expect(cardSend(CHANNEL, path, write)).rejects.toThrow("less than 28KB");
    expect(posts()).toHaveLength(0);
  });
});

describe("card send --template: the vote reactions", () => {
  test("after the post, each option emoji is added as your reaction, in order, to the new message", async () => {
    serve(posted("poll"));
    await cardSend(CHANNEL, dataFile("poll"), write);
    const order = writes().map(request => `${request.method} ${request.host.endsWith(".logic.azure.com") ? "router" : request.path.split("/").slice(-2).join("/")}`);
    expect(order[0]).toBe("POST router");
    expect(order.slice(1).every(item => item.endsWith("/properties"))).toBe(true);
    expect(order).toHaveLength(4);
    expect(putKeys()).toEqual(POLL_KEYS);
    for (const request of puts()) {
      expect(request.query).toEqual({ name: "emotions" });
      expect(request.path).toContain(`/conversations/${CHANNEL}/messages/`);
    }
    expect(logs.join("\n")).toContain("Added your vote reaction 🙋 (happy_person_raising_one_hand).");
    expect(logs.join("\n")).toContain("Ref: ");
    expect(process.exitCode ?? 0).toBe(0);
  });

  test("the reactions go through the write pacer, so a burst waits instead of failing", async () => {
    serve(posted("poll"));
    const start = clock;
    await cardSend(CHANNEL, dataFile("poll"), write);
    expect(puts()).toHaveLength(3);
    // The post and three reactions are four writes. The service allows 3 in 5 s, so the last one waited.
    expect(clock - start).toBeGreaterThanOrEqual(4_000);
  });

  test("a decision card with no given emoji reacts with the keycap digits, which have Teams keys", async () => {
    serve(posted("decision"));
    await cardSend(CHANNEL, dataFile("decision"), write);
    expect(putKeys()).toEqual(["keycapdigitone", "keycapdigittwo", "keycapdigitthree"]);
  });

  test("--no-react posts the card and reads and writes nothing else", async () => {
    serve(posted("poll"));
    await cardSend(CHANNEL, dataFile("poll"), { ...write, noReact: true });
    expect(posts()).toHaveLength(1);
    expect(puts()).toHaveLength(0);
    expect(fake!.requests.filter(request => request.method === "GET")).toHaveLength(0);
    expect(logs.join("\n")).not.toContain("vote reaction");
  });

  test("a template with no vote adds no reaction and does not wait", async () => {
    serve(posted("standup"));
    await cardSend(CHANNEL, dataFile("standup"), write);
    expect(posts()).toHaveLength(1);
    expect(puts()).toHaveLength(0);
    expect(fake!.requests.filter(request => request.method === "GET")).toHaveLength(0);
  });

  test("a card file never reacts", async () => {
    serve(posted("poll"));
    const path = join(dir, "card.json");
    writeFileSync(path, JSON.stringify(buildTemplateCard("poll", findTemplate("poll").example).card));
    await cardSend(CHANNEL, path, write);
    expect(puts()).toHaveLength(0);
  });

  test("a ref that is not found is said, no reaction is made, and the react commands are printed", async () => {
    serve({ ...posted("poll"), [LIST]: { status: 200, json: { messages: [] } } });
    await cardSend(CHANNEL, dataFile("poll"), write);
    expect(sleeps).toHaveLength(WAIT_POLLS - 1);
    expect(puts()).toHaveLength(0);
    const text = logs.join("\n");
    expect(text).toContain("did not show within 20 s");
    expect(text).toContain("The vote reactions were not added, because the ref of the card was not found.");
    for (const emoji of ["🙋", "🤝", "⛔"]) expect(text).toContain(`teams-cli react ${CHANNEL} <ref> ${emoji} --allow-write`);
    expect((await listRecords(TID))[0]!.state).toBe("sent");
  });

  test("a ref that is not found shows in the JSON, with the commands", async () => {
    serve({ ...posted("poll"), [LIST]: { status: 200, json: { messages: [] } } });
    await cardSend(CHANNEL, dataFile("poll"), { ...write, json: true });
    const out = JSON.parse(logs.join("\n"));
    expect(out.ref).toBeNull();
    expect(out.reactions.map((item: any) => item.status)).toEqual(["no-ref", "no-ref", "no-ref"]);
    expect(out.react_by_hand).toHaveLength(3);
    expect(out.find_ref).toContain("history");
  });

  test("the ledger: a second identical send is refused, so it posts and reacts only once", async () => {
    serve(posted("poll"));
    await cardSend(CHANNEL, dataFile("poll"), write);
    await expect(cardSend(CHANNEL, dataFile("poll"), write)).rejects.toThrow("Identical message already sent");
    expect(posts()).toHaveLength(1);
    expect(puts()).toHaveLength(3);
    const [record] = await listRecords(TID);
    expect(record!.state).toBe("sent");
    expect(record!.id).toStartWith(`${CHANNEL}#`);
  });

  test("a reaction you already have is not added again", async () => {
    serve(posted("poll", { handshake: [MINE] }));
    await cardSend(CHANNEL, dataFile("poll"), write);
    expect(putKeys()).toEqual(["happy_person_raising_one_hand", "26d4_noentry"]);
    expect(logs.join("\n")).toContain("You already had the reaction 🤝 (handshake). Nothing to add.");
  });

  test("one reaction that fails does not stop the others, and the exit code is 1", async () => {
    serve({ ...posted("poll"), [PUT]: request => ((request.body as any).emotions.key === "handshake" ? { status: 500 } : ok) });
    await cardSend(CHANNEL, dataFile("poll"), write);
    expect(putKeys()).toEqual(POLL_KEYS);
    expect(logs.join("\n")).toContain("Could not add 🤝 (handshake): Outcome unknown");
    expect(logs.join("\n")).toContain("teams-cli react");
    expect(process.exitCode).toBe(1);
  });

  test("a draft reacts in your Workflows chat", async () => {
    serve(posted("poll"));
    await cardSend(CHANNEL, dataFile("poll"), { ...write, draft: true });
    expect(posts()).toHaveLength(1);
    expect(putKeys()).toEqual(POLL_KEYS);
    for (const request of puts()) expect(request.path).toContain(`/conversations/${BOT_CHAT}/messages/`);
  });

  test("a router refusal posts nothing and reacts to nothing", async () => {
    serve({ ...posted("poll"), [ROUTE]: { status: 400, text: "bad" } });
    await expect(cardSend(CHANNEL, dataFile("poll"), write)).rejects.toThrow("The card router refused the card");
    expect(puts()).toHaveLength(0);
  });
});

describe("card update --template", () => {
  test("the preview names the template, and an update sends the built card with mode update", async () => {
    serve({ [ROUTE]: { status: 202 } });
    const source = dataFile("incident");
    await cardUpdate(`${CHANNEL}#${POSTED}`, source, { json: false, allowWrite: false });
    expect(logs.join("\n")).toContain("template: incident");
    expect(posts()).toHaveLength(0);
    await cardUpdate(`${CHANNEL}#${POSTED}`, source, write);
    const body = posts()[0]!.body as any;
    expect(body.target).toEqual({ mode: "update", channel: CHANNEL, messageId: POSTED });
    expect(body.attachments[0].content).toEqual(buildTemplateCard("incident", findTemplate("incident").example).card);
  });

  test("a longer timeline is a new card, so the ledger lets the second update through", async () => {
    serve({ [ROUTE]: { status: 202 } });
    const data: any = structuredClone(findTemplate("incident").example);
    await cardUpdate(`${CHANNEL}#${POSTED}`, dataFile("incident", data), write);
    data.timeline.push({ time: "4:55 PM", text: "Fixed", kind: "resolved" });
    data.resolved = true;
    await cardUpdate(`${CHANNEL}#${POSTED}`, dataFile("incident", data), write);
    expect(posts()).toHaveLength(2);
    expect(puts()).toHaveLength(0);
    await expect(cardUpdate(`${CHANNEL}#${POSTED}`, dataFile("incident", data), write)).rejects.toThrow("Identical message already sent");
  });
});

describe("card votes", () => {
  const poll = () => buildTemplateCard("poll", findTemplate("poll").example).card;
  const message = (users: Record<string, string[]>, extra: Record<string, unknown> = {}) => ({
    status: 200, json: { id: POSTED, properties: { cards: JSON.stringify([{ content: poll() }]), emotions: emotions(users), ...extra } },
  });
  const run = async (routes: Record<string, Route>, json = true) => {
    serve(routes);
    await cardVotes(`${CHANNEL}#${POSTED}`, { json });
    return json ? JSON.parse(logs.join("\n")) : logs.join("\n");
  };

  test("it counts other people's reactions per option, and leaves out your own seed reaction", async () => {
    const out = await run({ [ONE]: message({
      happy_person_raising_one_hand: [MINE, OTHER_A, OTHER_B], handshake: [MINE, OTHER_C], "26d4_noentry": [MINE],
    }) });
    expect(out.options.map((item: any) => [item.emoji, item.label, item.votes])).toEqual([["🙋", "I'll take it", 2], ["🤝", "I can back up", 1], ["⛔", "I have a conflict", 0]]);
    expect(out.total_votes).toBe(3);
    expect(out.own_reactions_excluded.sort()).toEqual([...POLL_KEYS].sort());
    expect(out.options_found).toBe(true);
    // It only reads.
    expect(writes()).toHaveLength(0);
  });

  test("without your seed, the counts are the same, and the same person under two options counts under each", async () => {
    const out = await run({ [ONE]: message({ happy_person_raising_one_hand: [OTHER_A], handshake: [OTHER_A, OTHER_B] }) });
    expect(out.options.map((item: any) => item.votes)).toEqual([1, 2, 0]);
    expect(out.own_reactions_excluded).toEqual([]);
  });

  test("a skin-tone variant counts for its option, and a reaction that is not an option is listed apart", async () => {
    const out = await run({ [ONE]: message({ "happy_person_raising_one_hand-tone3": [OTHER_A], like: [OTHER_B, OTHER_C], handshake: [MINE] }) });
    expect(out.options.map((item: any) => item.votes)).toEqual([1, 0, 0]);
    expect(out.other).toEqual([{ key: "like", emoji: "👍", votes: 2 }]);
    expect(out.total_votes).toBe(1);
  });

  test("the text output has one line per option and the total", async () => {
    const text = await run({ [ONE]: message({ handshake: [OTHER_A, OTHER_B] }) }, false);
    expect(text).toContain("Your own reactions are not counted.");
    expect(text).toMatch(/🙋 I'll take it\s+0/);
    expect(text).toMatch(/🤝 I can back up\s+2/);
    expect(text).toContain("Total votes: 2");
  });

  test("a decision card is read the same way, with the keycap options", async () => {
    serve({ [ONE]: { status: 200, json: { id: POSTED, properties: {
      cards: JSON.stringify([{ content: buildTemplateCard("decision", findTemplate("decision").example).card }]),
      emotions: emotions({ keycapdigitone: [MINE, OTHER_A], keycapdigitthree: [OTHER_A, OTHER_B] }),
    } } } });
    await cardVotes(`${CHANNEL}#${POSTED}`, { json: true });
    const out = JSON.parse(logs.join("\n"));
    expect(out.options.map((item: any) => [item.emoji, item.label, item.votes])).toEqual([["1️⃣", "Admins only", 1], ["2️⃣", "Audited", 0], ["3️⃣", "As is", 2]]);
  });

  test("a message without vote options lists every reaction", async () => {
    serve({ [ONE]: { status: 200, json: { id: POSTED, properties: { emotions: emotions({ like: [OTHER_A, MINE] }) } } } });
    await cardVotes(`${CHANNEL}#${POSTED}`, { json: true });
    const out = JSON.parse(logs.join("\n"));
    expect(out).toMatchObject({ options_found: false, options: [], total_votes: 1 });
    expect(out.other).toEqual([{ key: "like", emoji: "👍", votes: 1 }]);
  });

  test("a card from the Workflows bot is read from its Swift payload", async () => {
    const swift = Buffer.from(JSON.stringify({ attachments: [{ content: poll() }] })).toString("base64");
    const out = await run({ [ONE]: { status: 200, json: { id: POSTED, content: `<URIObject><Swift b64="${swift}"/></URIObject>`, properties: { emotions: emotions({ handshake: [OTHER_A] }) } } } });
    expect(out.options.map((item: any) => item.votes)).toEqual([0, 1, 0]);
  });

  test("a deleted message, and a ref that is not a message, fail", async () => {
    serve({ [ONE]: message({}, { deletetime: "1791990500000" }) });
    expect((await failure(cardVotes(`${CHANNEL}#${POSTED}`, { json: true }))).message).toContain("deleted");
    expect((await failure(cardVotes("nope", { json: true }))).message).toContain("is not the ref of a message");
  });

  test("voteOptions reads the ids in order, and countVotes does the sums with your id left out", () => {
    expect(voteOptions([poll()]).map(item => item.emoji)).toEqual(["🙋", "🤝", "⛔"]);
    const result = countVotes(
      [{ key: "handshake", mris: [MINE, OTHER_A, OTHER_A] }, { key: "like", mris: [MINE] }],
      [{ emoji: "🤝", label: "x" }, { emoji: "⛔", label: "y" }], MINE);
    expect(result.options.map(item => item.votes)).toEqual([1, 0]);
    expect(result.other).toEqual([]);
    expect(result.ownExcluded).toEqual(["handshake", "like"]);
  });
});
