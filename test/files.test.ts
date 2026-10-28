import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { loadProfile, saveProfile } from "../src/config.ts";
import { checkAudience, classifyToken, KINDS, requireToken, resetTokenCache } from "../src/tokens.ts";
import { scopeFor } from "../src/refresh.ts";
import { parseSharepointHost, personalSitePath, personalSiteUrl, sharepointHost, shareId } from "../src/sharepoint.ts";
import { fileFromLink, filesDownload, filesOf, safeFileName } from "../src/commands/files.ts";
import { authStatus } from "../src/commands/auth.ts";
import { requestHeaders } from "../src/transport.ts";
import { resetLocalState, seedProfile } from "./fake-teams.ts";
import { TID, fakeJwt } from "./jwt.ts";

const SP_AUD = "00000003-0000-0ff1-ce00-000000000000";
const HOST = "fake-my.sharepoint.com";
/** The team-site host of the same fake tenant. */
const TEAM_HOST = "fake.sharepoint.com";
const CONV = "48:notes";
const REF = `${CONV}#1700000000001`;
const FILE_URL = `https://${HOST}/personal/fake_example_test/Documents/Microsoft%20Teams%20Chat%20Files/a%20b%23c.txt`;

let original: typeof fetch;
let out: string;
const calls: Array<{ url: URL; init: RequestInit }> = [];
let dir: string;

let stored = "";
function spToken() { return fakeJwt({ aud: SP_AUD }); }

async function seed(withSharepoint = true, host: string | null = HOST) {
  await seedProfile();
  const { name, profile } = await loadProfile();
  const exp = Math.floor(Date.now() / 1000) + 3600;
  stored = spToken();
  await saveProfile(name, {
    ...profile!,
    ...(host ? { sharepointHost: host } : {}),
    tokens: { ...profile!.tokens, ...(withSharepoint ? { sharepoint: { token: stored, exp } } : {}) },
  });
}

function message(files: unknown, content = "") {
  return { id: "1700000000001", content, properties: { files: typeof files === "string" ? files : JSON.stringify(files) } };
}

/** The tag that the web client stores for an inline image. */
function imageTag(id: string, src = `https://us-api.asm.skype.com/v1/objects/${id}/views/imgo`) {
  return `<p><img src="${src}" itemtype="http://schema.skype.com/AMSImage" itemscope="png" width="178" height="192" alt="image" id="x_${id}" itemid="${id}"></p>`;
}

/** The smallest JPEG header that parseImageHeader reads: SOI, then a SOF0 frame of 192 by 178. */
const JPEG = Buffer.from([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0xc0, 0x00, 0xb2, 0x03]);

/** A fake for the Teams host and one SharePoint host. */
function fakeNet(options: { files?: unknown; bytes?: Buffer; downloadHost?: string; contentLength?: string; content?: string } = {}) {
  const files = options.files ?? [{ fileName: "a b#c.txt", fileInfo: { fileUrl: FILE_URL } }];
  globalThis.fetch = (async (input: any, init: RequestInit = {}) => {
    const url = new URL(String(input));
    calls.push({ url, init });
    if (url.host === "us-api.asm.skype.com" && /^\/v1\/objects\/[\w-]+\/views\/imgo$/.test(url.pathname)) return new Response(JPEG);
    if (url.host === "teams.cloud.microsoft") {
      if (url.pathname.endsWith("/messages/1700000000001")) return Response.json(message(files, options.content));
      if (url.pathname.endsWith("/users/ME/properties")) return Response.json({ personalFileSite: JSON.stringify({ personalSiteUrl: `https://${HOST}/personal/fake_example_test/`, availability: 1 }) });
    }
    if ((url.host === HOST || url.host === TEAM_HOST) && url.pathname.startsWith("/_api/v2.0/shares/")) {
      return Response.json({ driveItem: { "@content.downloadUrl": `https://${options.downloadHost ?? url.host}/personal/fake_example_test/_layouts/15/download.aspx?UniqueId=fake` } });
    }
    if (url.host === TEAM_HOST && url.pathname.includes("/_api/web/GetList(")) {
      if (url.searchParams.get("$skiptoken")) return Response.json({ value: [{ Id: 3, Title: "c", "OData__x0031_": "x" }] });
      return Response.json({ value: [{ Id: 1, Title: "a", OData__UIVersionString: "1.0" }, { Id: 2, Title: "b" }], "odata.nextLink": `https://${TEAM_HOST}/sites/engineering/_api/web/GetList('/sites/engineering/Lists/Env Tracker')/items?$skiptoken=Paged%3dTRUE&$top=5000` });
    }
    // Another organization's SharePoint refuses this tenant's token.
    if (url.host === "other.sharepoint.com") return new Response("{}", { status: 401 });
    if (url.pathname.endsWith("/download.aspx")) {
      return new Response(options.bytes ?? Buffer.from("hello"), { headers: options.contentLength ? { "content-length": options.contentLength } : {} });
    }
    return new Response("{}", { status: 404 });
  }) as typeof fetch;
}

beforeEach(async () => {
  original = globalThis.fetch;
  calls.length = 0;
  await resetLocalState();
  dir = mkdtempSync(join(tmpdir(), "teams-cli-files-"));
  out = "";
});
afterEach(() => { globalThis.fetch = original; resetTokenCache(); delete process.env.TEAMS_SHAREPOINT_HOST; });

describe("sharepoint token kind", () => {
  test("the SharePoint audience is the sharepoint kind, and auth status lists it", async () => {
    expect(classifyToken(spToken()).kind).toBe("sharepoint");
    expect(KINDS).toContain("sharepoint");
    expect(() => checkAudience("chat", { aud: SP_AUD })).toThrow("for the sharepoint service");
    await seed(false);
    fakeNet();
    const log = console.log;
    console.log = (...args: unknown[]) => { out += `${args.join(" ")}\n`; };
    try { await authStatus(); } finally { console.log = log; }
    expect(out).toContain("sharepoint: missing");
  });

  test("the scope names the personal site host", () => {
    const config = { clientId: "c", tenantId: TID, refreshToken: "r" };
    expect(scopeFor("sharepoint", config, HOST)).toBe(`https://${HOST}/.default openid profile offline_access`);
    expect(() => scopeFor("sharepoint", config)).toThrow("host is unknown");
    expect(scopeFor("chat", config)).toContain("ic3.teams.office.com");
    expect(scopeFor("sharepoint", { ...config, scopes: { sharepoint: "custom" } }, HOST)).toBe("custom");
  });

  test("an expired token renews with the sharepoint scope and the stored refresh token", async () => {
    await seed(false);
    const { name, profile } = await loadProfile();
    await saveProfile(name, { ...profile!, refresh: { clientId: "fake-client", tenantId: TID, refreshToken: "fake-refresh", origin: "https://teams.cloud.microsoft" } });
    let form: URLSearchParams | undefined;
    globalThis.fetch = (async (input: any, init: RequestInit) => {
      expect(new URL(String(input)).host).toBe("login.microsoftonline.com");
      form = new URLSearchParams(String(init.body));
      return Response.json({ access_token: spToken(), refresh_token: "fake-rotated" });
    }) as typeof fetch;
    const token = await requireToken("sharepoint");
    expect(token.kind).toBe("sharepoint");
    expect(form?.get("scope")).toBe(`https://${HOST}/.default openid profile offline_access`);
    expect(form?.get("grant_type")).toBe("refresh_token");
    expect((await loadProfile()).profile?.tokens.sharepoint).toBeDefined();
  });
});

describe("site host", () => {
  test("only a sharepoint.com host is accepted", () => {
    expect(parseSharepointHost("Fake-My.SharePoint.com")).toBe(HOST);
    expect(parseSharepointHost(`https://${HOST}/personal/x_y_z`)).toBe(HOST);
    expect(parseSharepointHost("fake-my.sharepoint.com.evil.test")).toBeUndefined();
    expect(parseSharepointHost("evil.test")).toBeUndefined();
    expect(parseSharepointHost("")).toBeUndefined();
    expect(parseSharepointHost(undefined)).toBeUndefined();
  });

  test("the personal site path replaces . and @ with _", () => {
    expect(personalSitePath("Tester.One@example.test")).toBe("/personal/tester_one_example_test");
    expect(personalSiteUrl(HOST, "a@b.test")).toBe(`https://${HOST}/personal/a_b_test`);
  });

  test("the host is found once from the chat properties and then saved", async () => {
    await seed(false, null);
    fakeNet();
    expect(await sharepointHost()).toBe(HOST);
    expect((await loadProfile()).profile?.sharepointHost).toBe(HOST);
    const before = calls.length;
    expect(await sharepointHost()).toBe(HOST);
    expect(calls.length).toBe(before);
  });

  test("a bad TEAMS_SHAREPOINT_HOST is refused", async () => {
    await seed();
    process.env.TEAMS_SHAREPOINT_HOST = "evil.test";
    await expect(sharepointHost()).rejects.toThrow("TEAMS_SHAREPOINT_HOST must be");
  });

  test("SharePoint requests carry no Teams client headers", () => {
    const headers = requestHeaders("sharepoint", "fake-token", { region: "amer", hasBody: false });
    expect(Object.keys(headers).sort()).toEqual(["accept", "application", "authorization", "user-agent"]);
  });
});

describe("files property", () => {
  test("a share id is u! and the unpadded base64url of the URL", () => {
    // This URL has bytes that standard base64 writes as + and /.
    const odd = "https://h/a?b>c~~~?>";
    expect(shareId(odd)).toBe(`u!${Buffer.from(odd).toString("base64url")}`);
    expect(shareId(odd)).not.toMatch(/[=+/]/);
    const id = shareId(FILE_URL);
    expect(id.startsWith("u!")).toBe(true);
    expect(id).not.toMatch(/[=+/]/);
    expect(Buffer.from(id.slice(2), "base64url").toString()).toBe(FILE_URL);
  });

  test("files are read from the JSON string and names are made safe", () => {
    expect(filesOf(message([{ fileName: "x.txt", fileInfo: { fileUrl: "https://h/x" } }, { title: "../../y", objectUrl: "https://h/y" }, { fileName: "no url" }]))).toEqual([
      { name: "x.txt", url: "https://h/x" },
      { name: ".._.._y", url: "https://h/y" },
    ]);
    expect(filesOf({ properties: {} })).toEqual([]);
    expect(() => filesOf(message("{not json"))).toThrow("cannot be read");
    expect(safeFileName("..")).toBe("file");
    expect(safeFileName("a/b\\c")).toBe("a_b_c");
  });
});

describe("files download", () => {
  test("it saves the file with mode 0600 and reads through the share and download URLs with the sharepoint token", async () => {
    await seed();
    fakeNet({ bytes: Buffer.from("fake bytes") });
    const log = console.log;
    console.log = (...args: unknown[]) => { out += `${args.join(" ")}\n`; };
    try { await filesDownload(REF, { out: dir }); } finally { console.log = log; }
    const path = join(dir, "a b#c.txt");
    expect(readFileSync(path, "utf8")).toBe("fake bytes");
    expect(statSync(path).mode & 0o777).toBe(0o600);
    expect(out).toContain(`saved ${path} (10 bytes)`);
    const share = calls.find(call => call.url.pathname.startsWith("/_api/v2.0/shares/"))!;
    expect(share.url.host).toBe(HOST);
    expect(share.url.pathname).toBe(`/_api/v2.0/shares/${shareId(FILE_URL)}/`);
    const headers = new Headers(share.init.headers);
    expect(headers.get("authorization")).toBe(`Bearer ${stored}`);
    expect(headers.get("clientinfo")).toBeNull();
    expect(calls.find(call => call.url.pathname.endsWith("/download.aspx"))?.url.searchParams.get("UniqueId")).toBe("fake");
  });

  test("it saves an inline image from the image host with the chat token, named from its bytes", async () => {
    await seed();
    fakeNet({ files: [], content: `<p>This is the schedule:</p>${imageTag("0-cus-d5-abc123")}` });
    const log = console.log;
    console.log = (...args: unknown[]) => { out += `${args.join(" ")}\n`; };
    try { await filesDownload(REF, { out: dir }); } finally { console.log = log; }
    // The tag says png, but the bytes are a JPEG, so the name follows the bytes.
    const path = join(dir, "1700000000001-image-1.jpg");
    expect(readFileSync(path).equals(JPEG)).toBe(true);
    expect(statSync(path).mode & 0o777).toBe(0o600);
    expect(out).toContain(`saved ${path} (${JPEG.length} bytes)`);
    const view = calls.find(call => call.url.host === "us-api.asm.skype.com")!;
    expect(view.url.pathname).toBe("/v1/objects/0-cus-d5-abc123/views/imgo");
    expect(new Headers(view.init.headers).get("authorization")).toBe(`Bearer ${(await requireToken("chat")).token}`);
    // An image-only message needs no SharePoint.
    expect(calls.some(call => call.url.host.endsWith(".sharepoint.com"))).toBe(false);
  });

  test("it saves the files and the inline images of one message together", async () => {
    await seed();
    fakeNet({ bytes: Buffer.from("fake bytes"), content: `${imageTag("0-cus-d5-one")}${imageTag("0-cus-d5-two")}${imageTag("0-cus-d5-one")}` });
    const log = console.log;
    console.log = () => {};
    try { await filesDownload(REF, { out: dir }); } finally { console.log = log; }
    expect(readFileSync(join(dir, "a b#c.txt"), "utf8")).toBe("fake bytes");
    expect(statSync(join(dir, "1700000000001-image-1.jpg")).size).toBe(JPEG.length);
    expect(statSync(join(dir, "1700000000001-image-2.jpg")).size).toBe(JPEG.length);
    // A repeated image is saved once.
    expect(calls.filter(call => call.url.host === "us-api.asm.skype.com")).toHaveLength(2);
  });

  test("an image address in the message content is never followed", async () => {
    await seed();
    fakeNet({ files: [], content: imageTag("0-cus-d5-abc123", "https://evil.test/steal") + `<img src="https://evil.test/x" itemtype="http://schema.skype.com/AMSImage" itemid="../../x">` });
    const log = console.log;
    console.log = () => {};
    try { await filesDownload(REF, { out: dir }); } finally { console.log = log; }
    expect(calls.some(call => call.url.host === "evil.test")).toBe(false);
    expect(calls.filter(call => call.url.host === "us-api.asm.skype.com").map(call => call.url.pathname)).toEqual(["/v1/objects/0-cus-d5-abc123/views/imgo"]);
  });

  test("a message with no files and no inline images fails, and an existing image name stops the download", async () => {
    await seed();
    fakeNet({ files: [], content: "<p>just text</p>" });
    await expect(filesDownload(REF, { out: dir })).rejects.toThrow("That message has no files and no inline images.");
    fakeNet({ files: [], content: imageTag("0-cus-d5-abc123") });
    writeFileSync(join(dir, "1700000000001-image-1.jpg"), "keep");
    await expect(filesDownload(REF, { out: dir })).rejects.toThrow("already exists. Nothing was downloaded.");
    expect(readFileSync(join(dir, "1700000000001-image-1.jpg"), "utf8")).toBe("keep");
  });

  test("it refuses a file that is not on SharePoint, and sends the token nowhere", async () => {
    await seed();
    fakeNet({ files: [{ fileName: "z.txt", fileInfo: { fileUrl: "https://evil.test/sites/s/z.txt" } }] });
    await expect(filesDownload(REF, { out: dir })).rejects.toThrow("reads only files on SharePoint");
    expect(calls.some(call => call.url.host === "evil.test")).toBe(false);
    for (const lookalike of ["https://sharepoint.com.evil.test/z.txt", "http://fake.sharepoint.com/z.txt"]) {
      fakeNet({ files: [{ fileName: "z.txt", fileInfo: { fileUrl: lookalike } }] });
      await expect(filesDownload(REF, { out: dir })).rejects.toThrow("reads only files on SharePoint");
    }
  });

  test("another organization's SharePoint is tried, and its refusal says so", async () => {
    await seed();
    fakeNet({ files: [{ fileName: "z.txt", fileInfo: { fileUrl: "https://other.sharepoint.com/sites/s/z.txt" } }] });
    await expect(filesDownload(REF, { out: dir })).rejects.toThrow("belongs to another organization");
    expect(calls.some(call => call.url.host === "other.sharepoint.com")).toBe(true);
  });

  test("it downloads a team-site file from that site's host with the same token", async () => {
    await seed();
    fakeNet({ files: [{ fileName: "plan.pdf", fileInfo: { fileUrl: `https://${TEAM_HOST}/sites/engineering/Shared%20Documents/plan.pdf` } }] });
    await filesDownload(REF, { out: dir });
    expect(readFileSync(join(dir, "plan.pdf"), "utf8")).toBe("hello");
    const share = calls.find(call => call.url.pathname.startsWith("/_api/v2.0/shares/"))!;
    const download = calls.find(call => call.url.pathname.endsWith("/download.aspx"))!;
    expect(share.url.host).toBe(TEAM_HOST);
    expect(download.url.host).toBe(TEAM_HOST);
    expect(new Headers(download.init.headers).get("authorization")).toStartWith("Bearer ");
  });

  test("a SharePoint link from search --files downloads that one file, with no message read", async () => {
    await seed();
    fakeNet();
    const link = `https://${TEAM_HOST}/sites/Operations/Shared Documents/Study Guide.docx?d=wf307fd2c66044588904e586f0e3f0772`;
    await filesDownload(link, { out: dir });
    expect(readFileSync(join(dir, "Study Guide.docx"), "utf8")).toBe("hello");
    expect(calls.some(call => call.url.host === "teams.cloud.microsoft")).toBe(false);
    const share = calls.find(call => call.url.pathname.startsWith("/_api/v2.0/shares/"))!;
    expect(share.url.host).toBe(TEAM_HOST);
    // The ?d= query is not part of the file's address.
    expect(share.url.pathname).toBe(`/_api/v2.0/shares/${shareId(`https://${TEAM_HOST}/sites/Operations/Shared%20Documents/Study%20Guide.docx`)}/`);
  });

  test("a link names the file from its last path segment, decoded and made safe", () => {
    expect(fileFromLink("https://x.sharepoint.com/sites/a/Shared%20Documents/round-robin.js")).toEqual({ name: "round-robin.js", url: "https://x.sharepoint.com/sites/a/Shared%20Documents/round-robin.js" });
    expect(fileFromLink("https://x.sharepoint.com/sites/a/b%2Fc.txt#frag").name).toBe("b_c.txt");
    expect(() => fileFromLink("https://")).toThrow("Not a link");
  });

  test("a link off SharePoint is refused before any request", async () => {
    await seed();
    fakeNet();
    await expect(filesDownload("https://evil.example.com/x.pdf", { out: dir })).rejects.toThrow("reads only files on SharePoint");
    expect(calls).toHaveLength(0);
  });

  test("it downloads a file from another person's OneDrive on the personal-site host", async () => {
    await seed();
    fakeNet({ files: [{ fileName: "notes.txt", fileInfo: { fileUrl: `https://${HOST}/personal/someone_else_test/Documents/notes.txt` } }] });
    await filesDownload(REF, { out: dir });
    expect(readFileSync(join(dir, "notes.txt"), "utf8")).toBe("hello");
  });

  test("a SharePoint list is saved as its rows, across pages, without SharePoint's own fields", async () => {
    await seed();
    fakeNet({ files: [{ fileName: "Env Tracker", fileInfo: { fileUrl: `https://${TEAM_HOST}/sites/engineering/Lists/Env%20Tracker` } }] });
    await filesDownload(REF, { out: dir });
    const rows = JSON.parse(readFileSync(join(dir, "Env Tracker.json"), "utf8"));
    expect(rows).toEqual([{ Id: 1, Title: "a" }, { Id: 2, Title: "b" }, { Id: 3, Title: "c" }]);
    const first = calls.find(call => call.url.pathname.includes("/_api/web/GetList("))!;
    expect(decodeURIComponent(first.url.pathname)).toBe("/sites/engineering/_api/web/GetList('/sites/engineering/Lists/Env Tracker')/items");
    expect(calls.some(call => call.url.pathname.startsWith("/_api/v2.0/shares/"))).toBe(false);
  });

  test("it refuses a download address on another host", async () => {
    await seed();
    fakeNet({ downloadHost: "evil.test" });
    await expect(filesDownload(REF, { out: dir })).rejects.toThrow("Refused a paging link");
  });

  test("it never overwrites a file", async () => {
    await seed();
    fakeNet();
    writeFileSync(join(dir, "a b#c.txt"), "mine");
    await expect(filesDownload(REF, { out: dir })).rejects.toThrow("already exists");
    expect(readFileSync(join(dir, "a b#c.txt"), "utf8")).toBe("mine");
    expect(calls.some(call => call.url.host === HOST)).toBe(false);
  });

  test("it refuses a file over 250 MB by its content length", async () => {
    await seed();
    fakeNet({ contentLength: String(250 * 1024 * 1024 + 1) });
    await expect(filesDownload(REF, { out: dir })).rejects.toThrow("larger than");
    expect(() => statSync(join(dir, "a b#c.txt"))).toThrow();
  });

  test("a message with no files, a bad ref, and a duplicate name fail plainly", async () => {
    await seed();
    fakeNet({ files: [] });
    await expect(filesDownload(REF, { out: dir })).rejects.toThrow("no files");
    await expect(filesDownload("nonsense", { out: dir })).rejects.toThrow("Give a message ref");
    fakeNet({ files: [{ fileName: "d.txt", fileInfo: { fileUrl: FILE_URL } }, { fileName: "d.txt", fileInfo: { fileUrl: FILE_URL } }] });
    await expect(filesDownload(REF, { out: dir })).rejects.toThrow("two files named d.txt");
  });
});
