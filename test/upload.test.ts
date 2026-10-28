import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { existsSync, mkdtempSync, readdirSync, truncateSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { send } from "../src/commands/send.ts";
import { cacheDir } from "../src/config.ts";
import { loadConversations, saveConversations } from "../src/conversations.ts";
import { ledgerKey } from "../src/ledger.ts";
import { checkFileSet, CHUNK_BYTES, channelTarget, chatFileEntry, channelFileEntry, inviteRecipients, loadFile, siteOfFolder } from "../src/upload.ts";
import { installFakeTeams, resetLocalState, seedProfile, type FakeRequest, type Route } from "./fake-teams.ts";
import { runCli } from "./cli.ts";
import { TID } from "./jwt.ts";

const ME = "8:orgid:11111111-1111-1111-1111-111111111111";
const PEOPLE = [1, 2, 3].map(n => `8:orgid:0000000${n}-0000-0000-0000-000000000000`);
const GONE = "8:orgid:0000000a-0000-0000-0000-00000000000a";
const PERSONAL_HOST = "fake-my.sharepoint.com";
const TEAM_HOST = "fake.sharepoint.com";
const SITE = "/personal/fake_example_test";
const CHAT = "19:fakechat@thread.v2";
const CHAN = "19:fakechan@thread.tacv2";
const NOTES = "48:notes";
const API = "/api/chatsvc/amer/v1";
const OK_SEND = { status: 201, json: { OriginalArrivalTime: 1791990424088 } };
const PROPERTIES = `GET ${API}/users/ME/properties`;
const CSA = "GET /api/csa/amer/api/v3/teams/users/me";
const PROFILES = "POST /api/mt/amer/beta/users/fetchShortProfile";
const ONEDRIVE = (name: string) => `PUT ${SITE}/_api/v2.0/drive/root:/Microsoft Teams Chat Files/${name}:/content`;
const INVITE = `POST ${SITE}/_api/v2.0/sites/root/items/{id}/driveItem/invite`;
const FOLDER_READ = "GET /sites/fakesite/_api/web/GetFolderByServerRelativePath(decodedurl=@v)";
const FILE_READ = "GET /sites/fakesite/_api/web/GetFileByServerRelativePath(decodedurl=@v)";
const STUB = "POST /sites/fakesite/_api/web/GetFolderByServerRelativeUrl(@v)/Files/addStub(@fn)";
const FILE_CALL = "POST /sites/fakesite/_api/web/{file}/{call}";
const FOLDER = "/sites/fakesite/Shared Documents/Original Name";

const dir = mkdtempSync(join(tmpdir(), "teams-upload-"));
function file(name: string, bytes: Uint8Array | string = "hello"): string {
  const path = join(dir, name);
  writeFileSync(path, bytes);
  return path;
}

let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
const originalLog = console.log;
const originalError = console.error;
const plain = { json: false, verify: false, allowWrite: false } as const;
const live = { json: false, verify: false, allowWrite: true } as const;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams({ [PROPERTIES]: { status: 200, json: { personalFileSite: JSON.stringify({ personalSiteUrl: `https://${PERSONAL_HOST}${SITE}/` }) } }, ...routes });
}

beforeEach(async () => {
  await resetLocalState();
  await seedProfile(["chat", "csa", "spaces", "sharepoint"]);
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

const messages = (conversation: string) => fake!.requests.filter(item => item.method === "POST" && item.path === `${API}/users/ME/conversations/${conversation}/messages`);
const writes = () => fake!.requests.filter(item => item.method !== "GET" && item.path !== "/api/mt/amer/beta/users/fetchShortProfile");
const ledgerDir = () => join(cacheDir(), TID, "ledger");

/** A OneDrive upload answer in the shape of capture 375, with fake values. */
function driveItem(name: string, n = 1) {
  const guid = `0000000${n}-aaaa-bbbb-cccc-000000000000`;
  return {
    id: `01FAKEITEM${n}`,
    name,
    webDavUrl: `https://${PERSONAL_HOST}${SITE}/Documents/Microsoft%20Teams%20Chat%20Files/${encodeURIComponent(name)}`,
    parentReference: { driveId: "b!fakedrive" },
    sharepointIds: { listId: "fake-list", listItemUniqueId: guid, siteId: "fake-site", siteUrl: `https://${PERSONAL_HOST}${SITE}`, webId: "fake-web" },
  };
}

/** The csa team list: one team with the channels that the folder rules need. */
function teamsMe(channels: any[] = [{ id: CHAN, displayName: "Renamed Channel", defaultFileSettings: { filesRelativePath: FOLDER } }]) {
  return { teams: [{ id: "19:faketeam@thread.tacv2", displayName: "Fake Team", teamSiteInformation: { sharepointSiteUrl: `https://${TEAM_HOST}/sites/fakesite` }, channels }], chats: [] };
}

function threadWith(members: string[]) {
  return { status: 200, json: { id: CHAT, members: members.map(id => ({ id, role: "Admin" })) } };
}

function profile(mri: string, name: string) {
  return { mri, objectId: mri.slice("8:orgid:".length), userPrincipalName: `${name.toLowerCase()}@example.test`, displayName: name, type: "ADUser" };
}

const goodProfiles = { status: 200, json: { type: "Thread", value: [profile(PEOPLE[0]!, "Ann"), profile(PEOPLE[1]!, "Bob"), profile(ME, "Test User"), { mri: GONE, isShortProfile: true, accountEnabled: false }] } };

/** A SharePoint upload that tracks the bytes it has taken, as the real calls did in capture 11. */
function channelUploadRoutes(options: { name: string; failOn?: string } = { name: "note.txt" }) {
  let taken = 0;
  const calls: string[] = [];
  const route = (request: FakeRequest) => {
    const call = request.path.split("/").pop()!;
    calls.push(call);
    if (options.failOn && call.startsWith(options.failOn)) return { status: 500 };
    const bytes = (request.body as Uint8Array | undefined)?.length ?? 0;
    if (call.startsWith("StartUpload")) { taken = bytes; return { status: 200, json: { value: String(taken) } }; }
    if (call.startsWith("ContinueUpload")) {
      expect(Number(/fileOffset=(\d+)/.exec(call)![1])).toBe(taken);
      taken += bytes;
      return { status: 200, json: { value: String(taken) } };
    }
    return { status: 200, json: { UniqueId: "00000009-aaaa-bbbb-cccc-000000000000", Name: options.name, ServerRelativeUrl: `${FOLDER}/${options.name}`, Length: String(taken) } };
  };
  return { calls, routes: {
    [CSA]: { status: 200, json: teamsMe() },
    [FOLDER_READ]: { status: 200, json: { Exists: true, ServerRelativeUrl: FOLDER } },
    [FILE_READ]: { status: 404 },
    [STUB]: { status: 200, json: { UniqueId: "00000009-aaaa-bbbb-cccc-000000000000", Name: options.name, Length: "0" } },
    [FILE_CALL]: route,
    [`POST ${API}/users/ME/conversations/${CHAN}/messages`]: OK_SEND,
  } satisfies Record<string, Route> };
}

describe("loading a file", () => {
  test("reads the size and SHA-256, and refuses a missing file, a folder, an empty file, and a file over 250 MB", async () => {
    const loaded = await loadFile(file("a.txt", "hello"));
    expect(loaded).toMatchObject({ name: "a.txt", size: 5, digest: "2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824" });
    await expect(loadFile(join(dir, "nope.txt"))).rejects.toThrow("Cannot read the file");
    await expect(loadFile(dir)).rejects.toThrow("is not a file");
    await expect(loadFile(file("empty.txt", ""))).rejects.toThrow("is empty");
    const big = file("big.bin", "x");
    truncateSync(big, 251 * 1024 * 1024);
    await expect(loadFile(big)).rejects.toThrow("The limit is 250 MB");
  });

  test("refuses a name that SharePoint rejects, and two files with one name", async () => {
    for (const name of ["a:b.txt", "a?b.txt", "a*b.txt", "a|b.txt", "a\"b.txt", "~$lock.docx", "dot."]) {
      await expect(loadFile(file(name))).rejects.toThrow("cannot be uploaded");
    }
    const one = await loadFile(file("same.txt"));
    expect(() => checkFileSet([one, one])).toThrow("Two files are named same.txt");
    expect(() => checkFileSet(Array.from({ length: 11 }, (_, index) => ({ ...one, name: `${index}.txt` })))).toThrow("10 files at most");
  });
});

describe("the channel folder", () => {
  test("a folder inside the team site belongs to the team site, and another folder to its own site", () => {
    const team = `https://${TEAM_HOST}/sites/fakesite`;
    expect(siteOfFolder("/sites/fakesite/Shared Documents/General", team)).toBe("/sites/fakesite");
    expect(siteOfFolder("/sites/fakesite-dev/Shared Documents", team)).toBe("/sites/fakesite-dev");
    expect(siteOfFolder("/teams/other/Shared Documents/Files", team)).toBe("/teams/other");
    expect(siteOfFolder("/odd/Shared Documents", team)).toBeUndefined();
    expect(siteOfFolder("/sites/fakesite/Shared Documents/x", undefined)).toBe("/sites/fakesite");
  });

  test("a channel renamed after it was created uploads to the folder that Teams records, not one made from its name", async () => {
    serve({ [CSA]: { status: 200, json: teamsMe() } });
    const target = await channelTarget(CHAN);
    expect(target).toEqual({ host: TEAM_HOST, sitePath: "/sites/fakesite", siteUrl: `https://${TEAM_HOST}/sites/fakesite`, folder: FOLDER, label: "Fake Team/Renamed Channel" });
  });

  test("General, and a private channel with its own site, resolve from the same rule", async () => {
    serve({ [CSA]: { status: 200, json: teamsMe([
      { id: "19:general@thread.tacv2", displayName: "General", defaultFileSettings: { filesRelativePath: "/sites/fakesite/Shared Documents/General" } },
      { id: "19:private@thread.tacv2", displayName: "Private", defaultFileSettings: { filesRelativePath: "/sites/fakesite-private/Shared Documents" } },
    ]) } });
    expect(await channelTarget("19:general@thread.tacv2")).toMatchObject({ sitePath: "/sites/fakesite", folder: "/sites/fakesite/Shared Documents/General" });
    expect(await channelTarget("19:private@thread.tacv2")).toMatchObject({ host: TEAM_HOST, sitePath: "/sites/fakesite-private", folder: "/sites/fakesite-private/Shared Documents" });
  });

  test("the folder is cached, and a cached channel needs no list request", async () => {
    serve({ [CSA]: { status: 200, json: teamsMe() } });
    await channelTarget(CHAN);
    const requests = fake!.requests.length;
    await channelTarget(CHAN);
    expect(fake!.requests.length).toBe(requests);
    expect((await loadConversations(TID)).teams[0]!.channels[0]!.filesPath).toBe(FOLDER);
  });

  test("a channel with no recorded folder reads the thread once, caches the answer, and otherwise refuses", async () => {
    const channels = [{ id: CHAN, displayName: "Fresh", defaultFileSettings: { filesRelativePath: null } }];
    serve({ [CSA]: { status: 200, json: teamsMe(channels) }, [`GET ${API}/threads/${CHAN}`]: { status: 200, json: { properties: { channelDocsFolderRelativeUrl: "/sites/fakesite/Shared Documents/Fresh" } } } });
    expect(await channelTarget(CHAN)).toMatchObject({ folder: "/sites/fakesite/Shared Documents/Fresh" });
    expect((await loadConversations(TID)).teams[0]!.channels[0]!.filesPath).toBe("/sites/fakesite/Shared Documents/Fresh");

    await resetLocalState();
    await seedProfile(["chat", "csa", "spaces", "sharepoint"]);
    serve({ [CSA]: { status: 200, json: teamsMe(channels) }, [`GET ${API}/threads/${CHAN}`]: { status: 200, json: { properties: {} } } });
    await expect(channelTarget(CHAN)).rejects.toThrow("Teams has no files folder recorded for Fake Team/Fresh. Open the channel's Files tab");
  });

  test("a shared channel, a site on another host, and an unknown channel are refused", async () => {
    serve({ [CSA]: { status: 200, json: teamsMe([{ id: CHAN, displayName: "Shared", isShared: true, defaultFileSettings: { filesRelativePath: FOLDER } }]) } });
    await expect(channelTarget(CHAN)).rejects.toThrow("is a shared channel");
    await resetLocalState();
    await seedProfile(["chat", "csa", "spaces", "sharepoint"]);
    const other = teamsMe();
    other.teams[0]!.teamSiteInformation.sharepointSiteUrl = "https://evil.sharepoint.com/sites/fakesite";
    serve({ [CSA]: { status: 200, json: other } });
    await expect(channelTarget(CHAN)).rejects.toThrow("not a SharePoint host of your organization");
    await expect(channelTarget("19:missing@thread.tacv2")).rejects.toThrow("is not in your team list");
  });
});

describe("invite recipients", () => {
  test("skips you and a member with no directory user, and reads the profile to prove it", async () => {
    serve({ [PROFILES]: goodProfiles });
    const set = await inviteRecipients([ME, ...PEOPLE.slice(0, 2), GONE, "28:bot"], ME);
    expect(set.recipients.map(item => item.objectId)).toEqual(PEOPLE.slice(0, 2).map(mri => mri.slice(8)));
    expect(set.skipped).toEqual([{ mri: "28:bot", reason: "not a directory user" }, { mri: GONE, reason: "account disabled or deleted" }]);
    expect(fake!.find("POST", "/api/mt/amer/beta/users/fetchShortProfile")[0]!.body).toEqual([...PEOPLE.slice(0, 2), GONE]);
  });

  test("a member that the profile read does not return is skipped, and a failed read stops the send", async () => {
    serve({ [PROFILES]: { status: 200, json: { value: [profile(PEOPLE[0]!, "Ann")] } } });
    const set = await inviteRecipients([ME, PEOPLE[0]!, PEOPLE[1]!], ME);
    expect(set.skipped).toEqual([{ mri: PEOPLE[1]!, reason: "no directory entry" }]);
    serve({ [PROFILES]: { status: 500 } });
    await expect(inviteRecipients([ME, PEOPLE[0]!], ME)).rejects.toThrow("500");
  });

  test("a disabled account that still has a name is skipped too", async () => {
    serve({ [PROFILES]: { status: 200, json: { value: [{ ...profile(PEOPLE[0]!, "Ann"), accountEnabled: false }, profile(PEOPLE[1]!, "Bob")] } } });
    const set = await inviteRecipients([ME, PEOPLE[0]!, PEOPLE[1]!], ME);
    expect(set.recipients.map(item => item.name)).toEqual(["Bob"]);
    expect(set.skipped).toEqual([{ mri: PEOPLE[0]!, reason: "account disabled or deleted" }]);
  });

  test("a chat with only you has nobody to invite and asks no profile", async () => {
    serve({ [PROFILES]: goodProfiles });
    expect(await inviteRecipients([ME], ME)).toEqual({ recipients: [], skipped: [] });
    expect(fake!.requests).toHaveLength(0);
  });
});

describe("OneDrive upload to a chat", () => {
  test("uploads, invites each other member to every file, then posts the message, in that order", async () => {
    serve({
      [`GET ${API}/threads/${CHAT}`]: threadWith([ME, ...PEOPLE.slice(0, 2), GONE]),
      [PROFILES]: goodProfiles,
      [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt", 1) },
      [ONEDRIVE("b.txt")]: { status: 201, json: driveItem("b.txt", 2) },
      [INVITE]: { status: 200, json: { value: [] } },
      [`POST ${API}/users/ME/conversations/${CHAT}/messages`]: OK_SEND,
    });
    await send(CHAT, ["see files"], { ...live, files: [file("a.txt"), file("b.txt")] });
    const order = writes().map(item => `${item.method} ${item.path.replace(/^.*\/(\w+)$/, "$1")}`);
    expect(order).toEqual(["PUT content", "PUT content", "POST invite", "POST invite", "POST messages"]);

    const put = fake!.requests.find(item => item.method === "PUT")!;
    expect(put.host).toBe(PERSONAL_HOST);
    expect(put.rawPath).toBe(`${SITE}/_api/v2.0/drive/root:/Microsoft%20Teams%20Chat%20Files/a.txt:/content`);
    expect(put.query).toEqual({ "@name.conflictBehavior": "rename", $select: "*,sharepointIds,webDavUrl" });
    expect(put.headers).toMatchObject({ scenario: "ShareUploadFile", scenariotype: "AUO", application: "Teams_Web", prefer: "manualRedirect,pacToken=N", "content-type": "application/octet-stream" });
    expect(new TextDecoder().decode(put.body as Uint8Array)).toBe("hello");

    const invites = fake!.requests.filter(item => item.path.endsWith("/invite"));
    expect(invites.map(item => item.path)).toEqual([
      `${SITE}/_api/v2.0/sites/root/items/00000001-aaaa-bbbb-cccc-000000000000/driveItem/invite`,
      `${SITE}/_api/v2.0/sites/root/items/00000002-aaaa-bbbb-cccc-000000000000/driveItem/invite`,
    ]);
    expect(invites[0]!.host).toBe(PERSONAL_HOST);
  });

  test("by default the invite is the web client's: edit access for each member and the default link, put in the message", async () => {
    serve({
      [`GET ${API}/threads/${CHAT}`]: threadWith([ME, ...PEOPLE.slice(0, 2), GONE]),
      [PROFILES]: goodProfiles,
      [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt") },
      [INVITE]: { status: 200, json: { value: [{ id: "perm-1", roles: ["write"], link: { scope: "anonymous", type: "edit", webUrl: "https://fake-my.sharepoint.com/:t:/g/personal/x/link" } }] } },
      [`POST ${API}/users/ME/conversations/${CHAT}/messages`]: OK_SEND,
    });
    await send(CHAT, [], { ...live, files: [file("a.txt")] });
    const invite = fake!.requests.find(item => item.path.endsWith("/invite"))!;
    expect(JSON.stringify(invite.body)).toBe(`{"recipients":[{"objectId":"${PEOPLE[0]!.slice(8)}"},{"objectId":"${PEOPLE[1]!.slice(8)}"}],"roles":["write"]}`);
    expect(invite.headers).toMatchObject({ scenario: "TeamsFiles", application: "Teams_Web", prefer: "ReturnSpecificErrorCode,getDefaultLink", "content-type": "application/json" });
    const entry = JSON.parse((messages(CHAT)[0]!.body as any).properties.files)[0];
    expect(entry.fileInfo).toMatchObject({ shareUrl: "https://fake-my.sharepoint.com/:t:/g/personal/x/link", shareId: "perm-1" });
  });  test("--access=people invites each member with sign-in and no email, asks no link, and --role=view gives read", async () => {
    serve({
      [`GET ${API}/threads/${CHAT}`]: threadWith([ME, ...PEOPLE.slice(0, 2), GONE]),
      [PROFILES]: goodProfiles,
      [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt") },
      [INVITE]: { status: 200, json: { value: [] } },
      [`POST ${API}/users/ME/conversations/${CHAT}/messages`]: OK_SEND,
    });
    await send(CHAT, [], { ...live, files: [file("a.txt")], access: "people", role: "view" });
    const invite = fake!.requests.find(item => item.path.endsWith("/invite"))!;
    expect(JSON.stringify(invite.body)).toBe(`{"recipients":[{"objectId":"${PEOPLE[0]!.slice(8)}"},{"objectId":"${PEOPLE[1]!.slice(8)}"}],"roles":["read"],"requireSignIn":true,"sendInvitation":false}`);
    expect(invite.headers).toMatchObject({ scenario: "TeamsFiles", application: "Teams_Web", prefer: "ReturnSpecificErrorCode", "content-type": "application/json" });
    expect(JSON.stringify(invite)).not.toContain("getDefaultLink");
  });  test("--access=private uploads, reads no members, and invites nobody", async () => {
    serve({
      [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt") },
      [`POST ${API}/users/ME/conversations/${CHAT}/messages`]: OK_SEND,
    });
    await send(CHAT, ["mine"], { ...live, files: [file("a.txt")], access: "private" });
    expect(fake!.requests.some(item => item.path.endsWith("/invite"))).toBe(false);
    expect(fake!.requests.some(item => item.path.startsWith(`${API}/threads/`))).toBe(false);
    expect(messages(CHAT)).toHaveLength(1);
  });

  test("a bad --access or --role is refused before any request", async () => {
    serve({});
    await expect(send(CHAT, ["x"], { ...plain, files: [file("a.txt")], access: "public" })).rejects.toThrow("is not a file access");
    await expect(send(CHAT, ["x"], { ...plain, files: [file("a.txt")], role: "owner" })).rejects.toThrow("is not a file role");
  });


  test("the message carries properties.files in the captured shape, with no share link", async () => {
    serve({
      [`GET ${API}/threads/${CHAT}`]: threadWith([ME, PEOPLE[0]!]),
      [PROFILES]: { status: 200, json: { value: [profile(PEOPLE[0]!, "Ann")] } },
      [ONEDRIVE("a b.csv")]: { status: 201, json: driveItem("a b.csv") },
      [INVITE]: { status: 200, json: {} },
      [`POST ${API}/users/ME/conversations/${CHAT}/messages`]: OK_SEND,
    });
    await send(CHAT, [], { ...live, files: [file("a b.csv")] });
    const body: any = messages(CHAT)[0]!.body;
    expect(body.content).toBe("");
    const site = `https://${PERSONAL_HOST}${SITE}`;
    const url = `${site}/Documents/Microsoft%20Teams%20Chat%20Files/a%20b.csv`;
    const guid = "00000001-aaaa-bbbb-cccc-000000000000";
    expect(JSON.parse(body.properties.files)).toEqual([{
      itemid: guid, fileName: "a b.csv", fileType: "csv",
      fileInfo: { itemId: "01FAKEITEM1", fileUrl: url, siteUrl: `${site}/`, serverRelativeUrl: "", shareUrl: null, shareId: null },
      fileChicletState: { serviceName: "p2p", state: "active" },
      "@type": "http://schema.skype.com/File", version: 2, id: guid, baseUrl: `${site}/`, objectUrl: url, openUrl: null, type: "csv", title: "a b.csv", state: "active",
      chicletBreadcrumbs: null, providerData: "", botFileProperties: {}, isUploadError: null, progressComplete: null, permissionScope: null,
      filePreview: { previewUrl: "", previewHeight: 0, previewWidth: 0 },
      sharepointIds: { driveId: "b!fakedrive", listId: "fake-list", listItemUniqueId: guid, siteId: "fake-site", siteUrl: site, webId: "fake-web" },
      publication: null, site: null,
    }]);
    expect(JSON.parse(body.properties.files)[0]).toEqual(chatFileEntry({ name: "a b.csv", itemId: "01FAKEITEM1", uniqueId: guid, fileUrl: url, siteUrl: site, driveId: "b!fakedrive", listId: "fake-list", siteId: "fake-site", webId: "fake-web" }));
  });

  test("Notes to self uploads and invites nobody", async () => {
    serve({ [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt") }, [`POST ${API}/users/ME/conversations/${NOTES}/messages`]: OK_SEND });
    await send("me", ["mine"], { ...live, files: [file("a.txt")] });
    expect(fake!.requests.filter(item => item.path.endsWith("/invite") || item.path.includes("/threads/") || item.path.endsWith("fetchShortProfile"))).toHaveLength(0);
    expect(messages(NOTES)).toHaveLength(1);
  });

  test("a chat with only you uploads and invites nobody", async () => {
    serve({
      [`GET ${API}/threads/${CHAT}`]: threadWith([ME]),
      [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt") },
      [`POST ${API}/users/ME/conversations/${CHAT}/messages`]: OK_SEND,
    });
    await send(CHAT, [], { ...live, files: [file("a.txt")] });
    expect(fake!.requests.filter(item => item.path.endsWith("/invite") || item.path.endsWith("fetchShortProfile"))).toHaveLength(0);
    expect(messages(CHAT)).toHaveLength(1);
  });

  test("a name with a space, a quote, a hash, a percent sign, and unicode is encoded in the path", async () => {
    const name = "it's #1 100% ünï.txt";
    serve({ [ONEDRIVE(name)]: { status: 201, json: driveItem(name) }, [`POST ${API}/users/ME/conversations/${NOTES}/messages`]: OK_SEND });
    await send("me", [], { ...live, files: [file(name)] });
    const put = fake!.requests.find(item => item.method === "PUT")!;
    expect(put.rawPath).toBe(`${SITE}/_api/v2.0/drive/root:/Microsoft%20Teams%20Chat%20Files/it's%20%231%20100%25%20%C3%BCn%C3%AF.txt:/content`);
    expect(put.path).toBe(`${SITE}/_api/v2.0/drive/root:/Microsoft Teams Chat Files/${name}:/content`);
  });
});

describe("channel upload", () => {
  test("addStub, StartUpload, and FinishUpload carry the captured addresses and headers, then the message follows", async () => {
    const { routes, calls } = channelUploadRoutes({ name: "note.txt" });
    serve(routes);
    await send(CHAN, ["see file"], { ...live, files: [file("note.txt", "hello")] });
    const [stub, start, finish] = fake!.requests.filter(item => item.method === "POST" && item.path.startsWith("/sites/"));
    expect(stub!.host).toBe(TEAM_HOST);
    expect(stub!.path).toBe("/sites/fakesite/_api/web/GetFolderByServerRelativeUrl(@v)/Files/addStub(@fn)");
    expect(stub!.query).toEqual({ "@v": `'${FOLDER}'`, "@fn": "'note.txt'" });
    expect(stub!.headers).toMatchObject({ scenariotype: "AUO", scenario: "ShareUploadFile", application: "Teams_Web", prefer: "manualRedirect" });
    expect(stub!.headers["x-requestdigest"]).toBeUndefined();
    expect(stub!.body).toBeUndefined();

    expect(start!.path).toMatch(/^\/sites\/fakesite\/_api\/web\/GetFileById\('00000009-aaaa-bbbb-cccc-000000000000'\)\/StartUpload\(uploadId=guid'[0-9a-f-]{36}'\)$/);
    expect(start!.query).toEqual({ "@folder": `'${FOLDER}'`, "@file": "'note.txt'" });
    expect(start!.headers).toMatchObject({ "content-type": "application/octet-stream", scenario: "ShareUploadFile" });
    expect(new TextDecoder().decode(start!.body as Uint8Array)).toBe("hello");

    const uploadId = /guid'([^']+)'/.exec(start!.path)![1];
    expect(finish!.path).toBe(`/sites/fakesite/_api/web/GetFileById('00000009-aaaa-bbbb-cccc-000000000000')/FinishUpload(uploadId=guid'${uploadId}',fileOffset=5)`);
    expect(finish!.headers.prefer).toBe("pacToken=N");
    expect(finish!.body).toBeUndefined();
    expect(calls).toHaveLength(2);
    expect(messages(CHAN)).toHaveLength(1);
  });

  test("a file above the chunk size goes in StartUpload and ContinueUpload pieces with the offset so far", async () => {
    const { routes, calls } = channelUploadRoutes({ name: "big.bin" });
    serve(routes);
    const size = CHUNK_BYTES * 2 + 5;
    await send(CHAN, [], { ...live, files: [file("big.bin", new Uint8Array(size))] });
    expect(calls.map(call => call.replace(/uploadId=guid'[^']+'/, "uploadId=guid'x'"))).toEqual([
      "StartUpload(uploadId=guid'x')",
      `ContinueUpload(uploadId=guid'x',fileOffset=${CHUNK_BYTES})`,
      `ContinueUpload(uploadId=guid'x',fileOffset=${CHUNK_BYTES * 2})`,
      `FinishUpload(uploadId=guid'x',fileOffset=${size})`,
    ]);
    const pieces = fake!.requests.filter(item => /StartUpload|ContinueUpload/.test(item.path)).map(item => (item.body as Uint8Array).length);
    expect(pieces).toEqual([CHUNK_BYTES, CHUNK_BYTES, 5]);
  });

  test("a name with a space, a quote, a hash, a percent sign, and unicode doubles the quote and encodes the rest", async () => {
    const name = "it's #1 100% ünï.txt";
    const { routes } = channelUploadRoutes({ name });
    serve(routes);
    await send(CHAN, [], { ...live, files: [file(name)] });
    const stub = fake!.requests.find(item => item.path.endsWith("addStub(@fn)"))!;
    expect(stub.search).toContain("@fn=%27it%27%27s%20%231%20100%25%20%C3%BCn%C3%AF.txt%27");
    expect(stub.query["@fn"]).toBe(`'it''s #1 100% ünï.txt'`);
    const start = fake!.requests.find(item => item.path.includes("StartUpload"))!;
    expect(start.query["@file"]).toBe(`'it''s #1 100% ünï.txt'`);
  });

  test("the message carries the team-site entry of capture 11, with no preview and no share link", async () => {
    const { routes } = channelUploadRoutes({ name: "note.txt" });
    serve(routes);
    await send(CHAN, ["see file"], { ...live, files: [file("note.txt")] });
    const body: any = messages(CHAN)[0]!.body;
    const url = `https://${TEAM_HOST}${FOLDER}/note.txt`;
    const guid = "00000009-aaaa-bbbb-cccc-000000000000";
    expect(body.amsreferences).toEqual([]);
    expect(body.postType).toBe("Standard");
    expect(JSON.parse(body.properties.files)).toEqual([{
      itemid: guid, fileName: "note.txt", fileType: "txt",
      fileInfo: { itemId: null, fileUrl: url, siteUrl: `https://${TEAM_HOST}/sites/fakesite`, serverRelativeUrl: "", shareUrl: null, shareId: null },
      fileChicletState: { serviceName: "teams", state: "active" },
      "@type": "http://schema.skype.com/File", version: 2, id: guid, baseUrl: `https://${TEAM_HOST}/sites/fakesite`, objectUrl: url, openUrl: null, type: "txt", title: "note.txt", state: "active",
      chicletBreadcrumbs: null, providerData: "", botFileProperties: {}, isUploadError: null, progressComplete: null, permissionScope: null,
      filePreview: { previewUrl: "", previewHeight: 0, previewWidth: 0 },
      sharepointIds: null, publication: null, site: null,
    }]);
    expect(JSON.parse(body.properties.files)[0]).toEqual(channelFileEntry({ name: "note.txt", uniqueId: guid, serverRelativeUrl: `${FOLDER}/note.txt`, host: TEAM_HOST, siteUrl: `https://${TEAM_HOST}/sites/fakesite` }));
    expect(fake!.requests.some(item => item.path.endsWith("/invite"))).toBe(false);
  });

  test("a thread reply carries the files, and a file that exists in the folder stops the send before any upload", async () => {
    const { routes } = channelUploadRoutes({ name: "note.txt" });
    serve({ ...routes, [`POST ${API}/users/ME/conversations/${CHAN};messageid=1791990424088/messages`]: OK_SEND });
    await send(CHAN, ["reply"], { ...live, files: [file("note.txt")], threadTs: "1791990424088" });
    const reply: any = fake!.requests.find(item => item.path.endsWith(";messageid=1791990424088/messages"))!.body;
    expect(reply.replyChainId).toBe("1791990424088");
    expect(JSON.parse(reply.properties.files)).toHaveLength(1);

    const records = readdirSync(ledgerDir());
    const taken = channelUploadRoutes({ name: "note.txt" });
    serve({ ...taken.routes, [FILE_READ]: { status: 200, json: { Exists: true, Name: "note.txt" } } });
    await expect(send(CHAN, [], { ...live, files: [file("note.txt")] })).rejects.toThrow("note.txt already exists in /sites/fakesite/Shared Documents/Original Name");
    expect(fake!.requests.filter(item => item.method === "POST")).toHaveLength(0);
    expect(readdirSync(ledgerDir())).toEqual(records);
  });
});

describe("preview", () => {
  test("a chat preview reads members and profiles, uploads nothing, invites nobody, and keeps no ledger record", async () => {
    serve({
      [`GET ${API}/threads/${CHAT}`]: threadWith([ME, ...PEOPLE.slice(0, 2), GONE]),
      [PROFILES]: goodProfiles,
      [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt") },
      [`POST ${API}/users/ME/conversations/${CHAT}/messages`]: OK_SEND,
    });
    await send(CHAT, ["see"], { ...plain, json: true, files: [file("a.txt")] });
    expect(writes()).toHaveLength(0);
    expect(existsSync(ledgerDir())).toBe(false);
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.files).toEqual([{ file: "a.txt", bytes: 5, destination: `OneDrive ${PERSONAL_HOST}${SITE}/Documents/Microsoft Teams Chat Files/a.txt` }]);
    expect(preview.sharing).toEqual({ access: "link", role: "edit" });
    expect(preview.invite).toEqual({
      default_link: true,
      roles: ["write"],
      recipients: [{ name: "Ann", object_id: PEOPLE[0]!.slice(8) }, { name: "Bob", object_id: PEOPLE[1]!.slice(8) }],
      skipped: [{ mri: GONE, reason: "account disabled or deleted" }],
    });
    expect(JSON.parse(preview.body.properties.files)[0]).toMatchObject({ itemid: "<file-1-guid>", fileName: "a.txt", fileInfo: { shareUrl: null, shareId: null } });
  });

  test("the text preview names each file, its size, its destination, and the people", async () => {
    serve({ [`GET ${API}/threads/${CHAT}`]: threadWith([ME, PEOPLE[0]!]), [PROFILES]: { status: 200, json: { value: [profile(PEOPLE[0]!, "Ann")] } } });
    await send(CHAT, ["see"], { ...plain, files: [file("a.txt")] });
    expect(logs[0]).toContain("[1 file: a.txt 5 bytes to OneDrive fake-my.sharepoint.com/personal/fake_example_test/Documents/Microsoft Teams Chat Files/a.txt; share with the tenant's default sharing link for 1 member: Ann (nobody is added by name; --access=people for that); no upload in a preview]");
  });

  test("a channel preview shows the resolved folder, reads it, and sends and uploads nothing", async () => {
    const { routes } = channelUploadRoutes({ name: "note.txt" });
    serve(routes);
    await send("Fake Team/Renamed Channel", [], { ...plain, json: true, files: [file("note.txt")] });
    expect(writes()).toHaveLength(0);
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.channel_folder).toBe(FOLDER);
    expect(preview.files[0].destination).toBe(`https://${TEAM_HOST}/sites/fakesite/Shared Documents/Original Name/note.txt`);
    expect(fake!.requests.map(item => item.method)).not.toContain("PUT");
    expect(fake!.requests.some(item => /StartUpload|addStub/.test(item.path))).toBe(false);
  });

  test("a channel preview says when a file name is taken", async () => {
    const { routes } = channelUploadRoutes({ name: "note.txt" });
    serve({ ...routes, [FILE_READ]: { status: 200, json: { Exists: true, Name: "note.txt" } } });
    await send(CHAN, [], { ...plain, json: true, files: [file("note.txt")] });
    expect(JSON.parse(logs.join("\n")).files[0].conflict).toContain("A send is refused");
  });
});

describe("a failed upload", () => {
  test("sends no message, drops the ledger record, and names the files that are already uploaded", async () => {
    serve({
      [`GET ${API}/threads/${CHAT}`]: threadWith([ME, PEOPLE[0]!]),
      [PROFILES]: { status: 200, json: { value: [profile(PEOPLE[0]!, "Ann")] } },
      [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt", 1) },
      [ONEDRIVE("b.txt")]: { status: 500 },
      [INVITE]: { status: 200, json: {} },
      [`POST ${API}/users/ME/conversations/${CHAT}/messages`]: OK_SEND,
    });
    const error = await send(CHAT, ["x"], { ...live, files: [file("a.txt"), file("b.txt")] }).then(() => new Error("no error"), error => error as Error);
    expect(error.message).toContain("No message was sent.");
    expect(error.message).toContain(`https://${PERSONAL_HOST}${SITE}/Documents/Microsoft%20Teams%20Chat%20Files/a.txt`);
    expect(error.message).toContain("Teams%20Chat%20Files/b.txt");
    expect(error.message).toContain("Teams Chat Files/b.txt (may exist)");
    expect(messages(CHAT)).toHaveLength(0);
    expect(fake!.requests.some(item => item.path.endsWith("/invite"))).toBe(false);
    expect(existsSync(ledgerDir()) ? readdirSync(ledgerDir()) : []).toEqual([]);
  });

  test("a failed invite sends no message and names the uploaded files", async () => {
    serve({
      [`GET ${API}/threads/${CHAT}`]: threadWith([ME, PEOPLE[0]!]),
      [PROFILES]: { status: 200, json: { value: [profile(PEOPLE[0]!, "Ann")] } },
      [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt", 1) },
      [INVITE]: { status: 403 },
      [`POST ${API}/users/ME/conversations/${CHAT}/messages`]: OK_SEND,
    });
    const error = await send(CHAT, ["x"], { ...live, files: [file("a.txt")] }).then(() => new Error("no error"), error => error as Error);
    expect(error.message).toContain("The invite for a.txt failed");
    expect(error.message).toContain("No message was sent.");
    expect(error.message).toContain("Teams%20Chat%20Files/a.txt");
    expect(messages(CHAT)).toHaveLength(0);
  });

  test("a failed channel upload sends no message and names the file, which may exist", async () => {
    const { routes } = channelUploadRoutes({ name: "note.txt", failOn: "FinishUpload" });
    serve(routes);
    const error = await send(CHAN, [], { ...live, files: [file("note.txt")] }).then(() => new Error("no error"), error => error as Error);
    expect(error.message).toContain("No message was sent.");
    expect(error.message).toContain(`/sites/fakesite/Shared Documents/Original Name/note.txt (may exist)`);
    expect(messages(CHAN)).toHaveLength(0);
  });

  test("a retry after the failure is not blocked by the send-once record", async () => {
    serve({ [ONEDRIVE("a.txt")]: [{ status: 500 }, { status: 201, json: driveItem("a.txt") }], [`POST ${API}/users/ME/conversations/${NOTES}/messages`]: OK_SEND });
    await expect(send("me", ["x"], { ...live, files: [file("a.txt")] })).rejects.toThrow("No message was sent.");
    await send("me", ["x"], { ...live, files: [file("a.txt")] });
    expect(messages(NOTES)).toHaveLength(1);
  });
});

describe("the send-once key", () => {
  const base = { conversationId: NOTES, content: "", mentionMris: [], contentType: "Text" };

  test("includes each file digest, in order, and leaves other keys unchanged", () => {
    const none = ledgerKey(base);
    const one = ledgerKey({ ...base, files: ["aa"] });
    expect(one).not.toBe(none);
    expect(ledgerKey({ ...base, files: ["bb"] })).not.toBe(one);
    expect(ledgerKey({ ...base, files: ["aa", "bb"] })).not.toBe(ledgerKey({ ...base, files: ["bb", "aa"] }));
    expect(ledgerKey({ ...base, files: [] })).toBe(none);
  });

  test("the same file sent twice is refused, and the same name with new bytes is not", async () => {
    serve({ [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt") }, [`POST ${API}/users/ME/conversations/${NOTES}/messages`]: OK_SEND });
    await send("me", ["x"], { ...live, files: [file("a.txt", "one")] });
    await expect(send("me", ["x"], { ...live, files: [file("a.txt", "one")] })).rejects.toThrow("Identical message already sent");
    await send("me", ["x"], { ...live, files: [file("a.txt", "two")] });
    expect(messages(NOTES)).toHaveLength(2);
  });
});

describe("flags", () => {
  test("--file with --at uploads and invites now, then posts a ScheduledDraft that carries the files", async () => {
    serve({
      [`GET ${API}/threads/${CHAT}`]: threadWith([ME, PEOPLE[0]!]),
      [PROFILES]: { status: 200, json: { value: [profile(PEOPLE[0]!, "Ann")] } },
      [ONEDRIVE("a.txt")]: { status: 201, json: driveItem("a.txt", 1) },
      [INVITE]: { status: 200, json: { value: [] } },
      [`POST ${API}/users/ME/drafts`]: { status: 201, json: { OriginalArrivalTime: 1792421450913 } },
    });
    await send(CHAT, ["later"], { ...live, at: "2099-01-02T15:30:00Z", files: [file("a.txt")] });
    const order = writes().map(item => `${item.method} ${item.path.replace(/^.*\/(\w+)$/, "$1")}`);
    expect(order).toEqual(["PUT content", "POST invite", "POST drafts"]);
    const draft: any = fake!.requests.find(item => item.path.endsWith("/users/ME/drafts"))!.body;
    expect(draft).toMatchObject({ draftType: "ScheduledDraft", innerThreadId: CHAT, draftDetails: { sendAt: String(Date.parse("2099-01-02T15:30:00Z")) } });
    expect(JSON.parse(draft.message.properties.files)[0]).toMatchObject({ fileName: "a.txt", itemid: "00000001-aaaa-bbbb-cccc-000000000000" });
    expect(messages(CHAT)).toHaveLength(0);
  });

  test("--image with --at is still refused, and more than 10 files are refused", async () => {
    serve({});
    const many = Array.from({ length: 11 }, (_, index) => file(`f${index}.txt`));
    await expect(send("me", ["x"], { ...plain, files: many })).rejects.toThrow("10 files at most");
  });

  test("--file is a valid send flag, repeats, and takes a value", async () => {
    const run = await runCli(["send", "me", "x", "--file"]);
    expect(run.stderr).toContain("--file requires a value.");
    const unknown = await runCli(["react", "me", "1", "like", "--file", "a.txt"]);
    expect(unknown.stderr).toContain("Flag `--file` is not valid with `teams-cli react`.");
  });
});
