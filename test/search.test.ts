import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, test } from "bun:test";
import {
  dateRange, fileNames, fileQuery, fileSearchBody, hitMarkers, messageQuery, normalizeFile, normalizeHit, resolveSender, search, searchBody,
} from "../src/commands/search.ts";
import { refreshConversations } from "../src/conversations.ts";
import { seedProfile, installFakeTeams, resetLocalState, type FakeRequest, type Reply } from "./fake-teams.ts";

// Hand-built shapes. The field names follow the live captures of 2026-10-19 (docs/search.md). No capture text is copied.
const CHAT = "19:chat0000000000000000000000000000@thread.v2";
const CHANNEL = "19:channel00000000000000000000000000@thread.tacv2";
const SEARCH = "POST /searchservice/api/v2/query";
const USERS = "POST /api/mt/amer/beta/users/searchV2";
const ME = "GET /api/csa/amer/api/v3/teams/users/me";
const EXT = "SkypeSpaces_ConversationPost_Extension_";

let fake: ReturnType<typeof installFakeTeams> | undefined;
let originalLog: typeof console.log;
let originalTZ: string | undefined;
let lines: string[] = [];

beforeAll(() => { originalTZ = process.env.TZ; process.env.TZ = "America/Indianapolis"; });
afterAll(() => { if (originalTZ === undefined) delete process.env.TZ; else process.env.TZ = originalTZ; });
beforeEach(async () => {
  originalLog = console.log;
  lines = [];
  console.log = (value?: unknown) => { lines.push(String(value)); };
  await resetLocalState();
  await seedProfile(["chat", "csa", "spaces", "substrate"]);
});
afterEach(() => { console.log = originalLog; fake?.uninstall(); fake = undefined; });

function serve(routes: Parameters<typeof installFakeTeams>[0]) { fake = installFakeTeams(routes); return fake; }
const sent = (index = 0): any => (fake!.find("POST", "/searchservice/api/v2/query")[index]!.body as any);
const messageRequest = (index = 0) => sent(index).entityRequests[0];
const source = (id: string, extra: Record<string, unknown> = {}, extensions: Record<string, unknown> = {}) => ({
  Source: {
    ClientThreadId: CHAT, InternetMessageId: id, DateTimeSent: "2026-10-19T20:11:24.188Z", Preview: `text ${id}`,
    From: { EmailAddress: { Name: "Pat Example", Address: "pat@example.test" } }, MentionsPreview: { IsMentioned: false },
    Extensions: { [`${EXT}FileData`]: null, [`${EXT}ImageSrc`]: null, [`${EXT}AmsReferences`]: null, [`${EXT}ParentMessageId`]: id, ...extensions },
    ...extra,
  },
});
const messageReply = (results: unknown[], more = false, total = results.length): Reply =>
  ({ status: 200, json: { EntitySets: [{ ResultSets: [{ Total: total, MoreResultsAvailable: more, Results: results }] }] } });
const person = (name: string, upn: string, n: number) =>
  ({ mri: `8:orgid:00000000-0000-0000-0000-00000000000${n}`, displayName: name, givenName: name.split(" ")[0], surname: name.split(" ")[1], userPrincipalName: upn });

describe("query strings", () => {
  test("plain words skip deleted messages", () => {
    expect(messageQuery({ text: "florida" })).toBe("NOT (isClientSoftDeleted:TRUE) florida");
  });

  test("a channel uses the clientthreadid form", () => {
    expect(messageQuery({ text: "florida", thread: CHANNEL })).toBe(`clientthreadid:${CHANNEL} AND NOT (isClientSoftDeleted:TRUE) florida`);
  });

  test("a chat, and Notes to self, use the chat form that hides topics and deleted messages", () => {
    const chat = `clientthreadid:${CHAT} AND NOT (Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String:(topic OR space)) AND NOT (isClientSoftDeleted:TRUE) florida`;
    expect(messageQuery({ text: "florida", thread: CHAT })).toBe(chat);
    expect(messageQuery({ text: "teams-cli", thread: "48:notes" })).toBe(chat.replace(CHAT, "48:notes").replace("florida", "teams-cli"));
  });

  test("--from adds the email as a plain term", () => {
    expect(messageQuery({ text: "florida", email: "pat@example.test" })).toBe("NOT (isClientSoftDeleted:TRUE) pat@example.test florida");
    expect(messageQuery({ text: "florida", email: "pat@example.test", thread: CHAT })).toEndWith("AND NOT (isClientSoftDeleted:TRUE) pat@example.test florida");
  });

  test("a query of filters only has no stray words", () => {
    expect(messageQuery({ text: "", thread: CHANNEL })).toBe(`clientthreadid:${CHANNEL} AND NOT (isClientSoftDeleted:TRUE)`);
    expect(messageQuery({ text: "", email: "pat@example.test" })).toBe("NOT (isClientSoftDeleted:TRUE) pat@example.test");
  });

  test("dates use sent bounds at local midnight, in UTC", () => {
    expect(messageQuery({ text: "florida", email: "pat@example.test", on: "2026-10-18" }))
      .toBe("sent >= 2026-10-18T04:00:00.000Z AND sent < 2026-10-19T04:00:00.000Z AND NOT (isClientSoftDeleted:TRUE) pat@example.test florida");
    expect(messageQuery({ text: "x", thread: CHANNEL, after: "2026-10-18" })).toBe(`clientthreadid:${CHANNEL} AND sent >= 2026-10-18T04:00:00.000Z AND NOT (isClientSoftDeleted:TRUE) x`);
  });

  test("file queries use LastModifiedTime, the email as a term, and a type filter", () => {
    expect(fileQuery({ text: "florida", email: "pat@example.test", on: "2026-10-18" }))
      .toBe("LastModifiedTime >= 2026-10-18T04:00:00.000Z AND LastModifiedTime < 2026-10-19T04:00:00.000Z AND pat@example.test florida");
    expect(fileQuery({ text: "plan", type: ".Loop" })).toBe("(filetype:loop OR fileExtension:loop) AND plan");
    expect(fileQuery({ text: "florida" })).toBe("florida");
    expect(() => fileQuery({ text: "x", type: "a b" })).toThrow("--type");
  });
});

describe("date flags", () => {
  test("--on is one local day, and --after and --before are local midnights", () => {
    expect(dateRange({ on: "2026-10-18" })).toEqual({ from: "2026-10-18T04:00:00.000Z", to: "2026-10-19T04:00:00.000Z" });
    expect(dateRange({ after: "2026-10-18" })).toEqual({ from: "2026-10-18T04:00:00.000Z" });
    expect(dateRange({ before: "2026-10-18" })).toEqual({ to: "2026-10-18T04:00:00.000Z" });
  });

  test("the offset follows daylight saving time", () => {
    // The clocks in Indiana go back on 2026-11-01, so that day is 25 hours long.
    expect(dateRange({ on: "2026-11-01" })).toEqual({ from: "2026-11-01T04:00:00.000Z", to: "2026-11-02T05:00:00.000Z" });
  });

  test("bad dates and impossible ranges fail", () => {
    expect(() => dateRange({ after: "2026-02-30" })).toThrow("YYYY-MM-DD");
    expect(() => dateRange({ after: "2026-10-03", before: "2026-10-01" })).toThrow("earlier");
    expect(() => dateRange({ after: "2026-10-03", before: "2026-10-03" })).toThrow("earlier");
    expect(() => dateRange({ on: "2026-10-03", after: "2026-10-01" })).toThrow("--on");
  });
});

describe("--from", () => {
  test("an email is used as typed, with no lookup", async () => {
    serve({});
    expect(await resolveSender("pat@example.test")).toBe("pat@example.test");
    expect(fake!.requests).toHaveLength(0);
  });

  test("a name that matches one person gives that person's email", async () => {
    serve({ [USERS]: { status: 200, json: { value: [person("Pat Example", "pat@example.test", 1)] } } });
    expect(await resolveSender("Pat")).toBe("pat@example.test");
  });

  test("several matches fail and list each person", async () => {
    serve({ [USERS]: { status: 200, json: { value: [person("Pat Example", "pat@example.test", 1), person("Pat Other", "pat.other@example.test", 2)] } } });
    const error = (await resolveSender("Pat").catch(failure => failure)) as Error;
    expect(error.message).toContain("matches 2 people");
    expect(error.message).toContain("Pat Example <pat@example.test>");
    expect(error.message).toContain("Pat Other <pat.other@example.test>");
  });

  test("an exact name wins over longer names", async () => {
    serve({ [USERS]: { status: 200, json: { value: [person("Pat Example", "pat@example.test", 1), person("Pat Examples", "pe@example.test", 2)] } } });
    expect(await resolveSender("Pat Example")).toBe("pat@example.test");
  });

  test("no match fails", async () => {
    serve({ [USERS]: { status: 200, json: { value: [] } } });
    await expect(resolveSender("Nobody")).rejects.toThrow("matches nobody");
  });

  test("search sends the resolved email, and sends nothing when the name is ambiguous", async () => {
    serve({
      [USERS]: { status: 200, json: { value: [person("Pat Example", "pat@example.test", 1)] } },
      [SEARCH]: messageReply([source("1791990000000")]),
    });
    await search("florida", { from: "Pat", json: true });
    expect(messageRequest().query.queryString).toBe("NOT (isClientSoftDeleted:TRUE) pat@example.test florida");
    fake!.uninstall();
    serve({
      [USERS]: { status: 200, json: { value: [person("Pat Example", "pat@example.test", 1), person("Pat Other", "po@example.test", 2)] } },
      [SEARCH]: messageReply([]),
    });
    await expect(search("florida", { from: "Pat" })).rejects.toThrow("matches 2 people");
    expect(fake!.find("POST", "/searchservice/api/v2/query")).toHaveLength(0);
  });
});

describe("the Messages request", () => {
  test("uses the Messages-tab scenario, size 25, and the extension fields", async () => {
    serve({ [SEARCH]: messageReply([source("1791990000000")]) });
    await search("florida", { json: true });
    const body = sent();
    expect(body.scenario).toEqual({
      Name: "powerbar",
      Dimensions: [{ DimensionName: "QueryType", DimensionValue: "Messages" }, { DimensionName: "FormFactor", DimensionValue: "general.web.reactSearch" }],
    });
    const request = body.entityRequests[0];
    expect(request).toMatchObject({ entityType: "Message", contentSources: ["Teams"], propertySet: "Optimized", from: 0, size: 25 });
    expect(request.query).toEqual({ queryString: "NOT (isClientSoftDeleted:TRUE) florida", displayQueryString: "florida" });
    expect(request.fields).toEqual(["FromSkypeInternalId", "FileData", "ThreadType", "SkypeGroupId", "SenderTenantId", "ImageSrc", "ParentMessageId", "AmsReferences"]
      .map(name => `Extension_${EXT}${name}_${name === "AmsReferences" ? "StringArray" : "String"}`));
    // The typeahead scenario of the old code must not come back.
    expect(JSON.stringify(body)).not.toContain("PowerbarTypeAheadMsg");
    expect(searchBody("x", 0, 25)).toHaveProperty("entityRequests.0.topResultsCount", 5);
  });

  test("--in takes a channel id, a Team/Channel name, and me", async () => {
    serve({
      [ME]: { status: 200, json: { chats: [], teams: [{ id: "19:team@thread.tacv2", displayName: "Eng", channels: [{ id: CHANNEL, displayName: "Deploys" }] }] } },
      [SEARCH]: messageReply([source("1791990000000")]),
    });
    await search("florida", { in: CHANNEL, json: true });
    await search("florida", { in: "Eng/Deploys", json: true });
    await search("teams-cli", { in: "me", json: true });
    expect(messageRequest(0).query.queryString).toBe(`clientthreadid:${CHANNEL} AND NOT (isClientSoftDeleted:TRUE) florida`);
    expect(messageRequest(1).query.queryString).toBe(`clientthreadid:${CHANNEL} AND NOT (isClientSoftDeleted:TRUE) florida`);
    expect(messageRequest(2).query.queryString).toStartWith("clientthreadid:48:notes AND NOT (Extension_");
  });

  test("a filters-only query needs --from or --in", async () => {
    serve({ [SEARCH]: messageReply([]) });
    await expect(search("", {})).rejects.toThrow("--from, or --in");
    await expect(search("", { on: "2026-10-18" })).rejects.toThrow("--from, or --in");
    await search("", { in: CHANNEL, json: true });
    expect(messageRequest().query.queryString).toBe(`clientthreadid:${CHANNEL} AND NOT (isClientSoftDeleted:TRUE)`);
  });

  test("pages with from, keeps the page size, and removes duplicates", async () => {
    let call = 0;
    serve({
      [SEARCH]: (request: FakeRequest) => {
        call++;
        void request;
        return messageReply(call === 1 ? [source("1791990000000"), source("1791990000001")] : [source("1791990000001"), source("1791990000002")], call === 1, 3);
      },
    });
    await search("hello", { count: 3, json: true });
    expect(JSON.parse(lines.at(-1)!).messages.map((message: any) => message.ref.split("#")[1])).toEqual(["1791990000000", "1791990000001", "1791990000002"]);
    expect([0, 1].map(index => messageRequest(index).from)).toEqual([0, 2]);
    expect([0, 1].map(index => messageRequest(index).size)).toEqual([25, 25]);
  });
});

describe("message hits", () => {
  const fileData = (items: unknown) => JSON.stringify([{ type: "Reference", rawJson: JSON.stringify(items) }]);

  test("an empty FileData is not a file, and a plain hit has no markers", () => {
    const hit = normalizeHit(source("1791990000000", {}, { [`${EXT}FileData`]: fileData([]) }));
    expect(hit.has_files).toBe(false);
    expect(hitMarkers(hit)).toEqual([]);
  });

  test("file names come from FileData, and [files] covers a file with no name", () => {
    const named = normalizeHit(source("1791990000000", {}, { [`${EXT}FileData`]: fileData([{ fileName: "plan.pdf" }, { title: "notes.docx" }]) }));
    expect(hitMarkers(named)).toEqual(["[files: plan.pdf, notes.docx]"]);
    const unnamed = normalizeHit(source("1791990000000", {}, { [`${EXT}FileData`]: fileData([{ id: "x" }]) }));
    expect(hitMarkers(unnamed)).toEqual(["[files]"]);
    expect(fileNames("not json")).toEqual({ has: false, names: [] });
  });

  test("an image comes from ImageSrc or AmsReferences", () => {
    expect(hitMarkers(normalizeHit(source("1791990000000", {}, { [`${EXT}ImageSrc`]: "https://example.test/img" })))).toEqual(["[image]"]);
    expect(hitMarkers(normalizeHit(source("1791990000000", {}, { [`${EXT}AmsReferences`]: ["0-id"] })))).toEqual(["[image]"]);
  });

  test("a reply names its thread root as a ref that context accepts", () => {
    const reply = normalizeHit(source("1791990000009", { ClientThreadId: CHANNEL }, { [`${EXT}ParentMessageId`]: "1791990000001" }));
    expect(reply.reply_in_thread).toBe(`${CHANNEL}#1791990000001`);
    expect(hitMarkers(reply)).toEqual([`reply in thread ${CHANNEL}#1791990000001`]);
    expect(normalizeHit(source("1791990000009", { ClientThreadId: CHANNEL })).reply_in_thread).toBeUndefined();
  });

  test("a mention of you is marked", () => {
    expect(hitMarkers(normalizeHit(source("1791990000000", { MentionsPreview: { IsMentioned: true } })))).toEqual(["[mentions you]"]);
  });

  test("text and JSON output show the label, sender, local time, ref, and markers", async () => {
    serve({
      [ME]: { status: 200, json: { chats: [{ id: CHAT, title: "Florida team", isOneOnOne: false }], teams: [] } },
      [SEARCH]: messageReply([source("1791990000000", {}, { [`${EXT}ImageSrc`]: "https://example.test/img", [`${EXT}ParentMessageId`]: "1791990000001" })]),
    });
    await refreshConversations();
    await search("florida", {});
    expect(lines[0]).toBe("[2026-10-19 16:11] Pat Example in Florida team\n  text 1791990000000");
    expect(lines[1]).toBe(`  id: ${CHAT}#1791990000000  [image]  reply in thread ${CHAT}#1791990000001`);
    lines = [];
    await search("florida", { json: true });
    expect(JSON.parse(lines[0]!).messages[0]).toMatchObject({
      ref: `${CHAT}#1791990000000`, conversation: "Florida team", sender: "Pat Example", sent_local: "2026-10-19 16:11",
      has_image: true, has_files: false, reply_in_thread: `${CHAT}#1791990000001`,
    });
  });
});

describe("file search", () => {
  const file = (name: string, extra: Record<string, unknown> = {}) => ({
    Source: {
      FileName: name, FileType: "pdf", LastModifiedTime: "2026-10-18T15:30:00Z", ModifiedBy: "Pat Example", Size: 1024,
      DefaultEncodingUrl: `https://example.sharepoint.com/sites/eng/Shared%20Documents/${name}`, SiteName: "https://example.sharepoint.com/sites/eng",
      ParentLink: "https://example.sharepoint.com/sites/eng/Shared Documents", UniqueId: `{${name}-id}`, ...extra,
    },
  });
  const fileReply = (results: unknown[]): Reply => ({ status: 200, json: { EntitySets: [{ ResultSets: [{ Total: results.length, Results: results }] }] } });

  test("sends the Files-tab request shape", async () => {
    serve({ [SEARCH]: fileReply([file("a.pdf")]) });
    await search("florida", { files: true, from: "pat@example.test", on: "2026-10-18", type: "pdf", json: true });
    const body = sent();
    expect(body.scenario.Dimensions).toEqual([{ DimensionName: "QueryType", DimensionValue: "Files" }, { DimensionName: "FormFactor", DimensionValue: "general.cdlworker.reactSearch" }]);
    const request = body.entityRequests[0];
    expect(request).toMatchObject({
      entityType: "File", contentSources: ["OneDriveBusiness", "SharePoint", "Exchange"], size: 25, from: 0,
      Sort: [{ Field: "PersonalScore", SortDirection: "Desc" }], extendedQueries: [{ SearchProvider: "SharePoint" }],
    });
    expect(request.query.queryString).toBe(
      "LastModifiedTime >= 2026-10-18T04:00:00.000Z AND LastModifiedTime < 2026-10-19T04:00:00.000Z AND (filetype:pdf OR fileExtension:pdf) AND pat@example.test florida");
    expect(request.fields).toContain("DefaultEncodingUrl");
    expect(request.fields).toContain("ModifiedBy");
    expect(request.RefiningQueries[0].RefinerString).toStartWith("or(andnot(IsDocument:true");
  });

  test("prints the name, folder, modified time and person, and link", async () => {
    serve({ [SEARCH]: fileReply([file("a.pdf"), file("Florida EVV", { FileType: "onepart", LinkingUrl: "https://example.sharepoint.com/x.one?d=w1" })]) });
    await search("florida", { files: true });
    expect(lines).toEqual([
      "a.pdf", "  in: https://example.sharepoint.com/sites/eng/Shared Documents", "  modified: 2026-10-18 11:30 by Pat Example",
      "  link: https://example.sharepoint.com/sites/eng/Shared%20Documents/a.pdf",
      "Florida EVV", "  in: https://example.sharepoint.com/sites/eng/Shared Documents", "  modified: 2026-10-18 11:30 by Pat Example",
      "  link: https://example.sharepoint.com/x.one?d=w1",
    ]);
    expect(normalizeFile(file("a.pdf")).unique_id).toBe("a.pdf-id");
  });

  test("pages with from", async () => {
    let call = 0;
    serve({ [SEARCH]: () => { call++; return call === 1 ? { status: 200, json: { EntitySets: [{ ResultSets: [{ MoreResultsAvailable: true, Results: [file("a.pdf"), file("b.pdf")] }] }] } } : fileReply([file("c.pdf")]); } });
    await search("florida", { files: true, count: 3, json: true });
    expect(JSON.parse(lines[0]!).files.map((item: any) => item.name)).toEqual(["a.pdf", "b.pdf", "c.pdf"]);
    expect([0, 1].map(index => messageRequest(index).from)).toEqual([0, 2]);
  });

  test("--in with --files scopes to the channel's files folder, as the Files tab does", async () => {
    await seedProfile(["chat", "csa", "spaces", "substrate", "sharepoint"]);
    const team = { id: "19:team@thread.tacv2", displayName: "Eng", teamSiteInformation: { groupId: "group-1", sharepointSiteUrl: "https://fake.sharepoint.com/sites/eng" },
      channels: [{ id: CHANNEL, displayName: "Deploys", defaultFileSettings: { filesRelativePath: "/sites/eng/Shared Documents/Deploys" } }] };
    serve({ [ME]: { status: 200, json: { chats: [], teams: [team] } }, [SEARCH]: fileReply([file("a.pdf")]) });
    await search("plan", { files: true, in: "Eng/Deploys", json: true });
    const request = messageRequest();
    const folder = "https://fake.sharepoint.com/sites/eng/Shared Documents/Deploys";
    expect(request.contentSources).toEqual(["SharePoint"]);
    expect(request.query.queryString).toBe(`(ParentLink:"${folder}" OR ParentLink:"${folder}/*") AND plan`);
    expect(request.query.queryAnnotations).toEqual([{ type: "Channel", threadId: CHANNEL, text: "dummyText", groupId: "group-1", path: `${folder}/`, tenantId: expect.any(String) }]);
  });

  test("a file search without --in keeps every source and has no annotation", () => {
    const request = (fileSearchBody("plan", 0, 25) as any).entityRequests[0];
    expect(request.contentSources).toEqual(["OneDriveBusiness", "SharePoint", "Exchange"]);
    expect(request.query).not.toHaveProperty("queryAnnotations");
    expect(fileQuery({ text: "plan", folder: "https://x.sharepoint.com/sites/a/Shared Documents" }))
      .toBe('(ParentLink:"https://x.sharepoint.com/sites/a/Shared Documents" OR ParentLink:"https://x.sharepoint.com/sites/a/Shared Documents/*") AND plan');
  });

  test("--in with --files refuses a chat, --type needs --files, and a filters-only query needs --from or --type", async () => {
    serve({ [SEARCH]: fileReply([]) });
    await expect(search("x", { files: true, in: "me" })).rejects.toThrow("--in with --files takes a channel");
    await expect(search("x", { type: "pdf" })).rejects.toThrow("--type works only with --files");
    await expect(search("", { files: true })).rejects.toThrow("--from, or --type");
    expect(fake!.requests).toHaveLength(0);
  });
});
