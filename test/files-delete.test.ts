import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { loadProfile, saveProfile } from "../src/config.ts";
import { resetTokenCache } from "../src/tokens.ts";
import { filesDelete, serverRelativeAddress } from "../src/commands/files.ts";
import { deleteMessages } from "../src/commands/delete.ts";
import { setPaceForTests } from "../src/pace.ts";
import { resetLocalState, seedProfile } from "./fake-teams.ts";
import { fakeJwt } from "./jwt.ts";

const SP_AUD = "00000003-0000-0ff1-ce00-000000000000";
const HOST = "fake.sharepoint.com";
const SITE = `https://${HOST}/sites/fake-team`;
const REF = "48:notes#1700000000001";
const FOLDER = "/sites/fake-team/Shared Documents/General";
const NAME = "q's #1 100% café.png";
const FILE_URL = `https://${HOST}${FOLDER.replace(/ /g, "%20")}/${encodeURIComponent(NAME).replace(/'/g, "%27")}`;

type Call = { method: string; url: URL };
let original: typeof fetch;
let calls: Call[] = [];
let logs: string[] = [];
const originalLog = console.log;

async function seed() {
  await seedProfile();
  const { name, profile } = await loadProfile();
  const exp = Math.floor(Date.now() / 1000) + 3600;
  await saveProfile(name, { ...profile!, sharepointHost: "fake-my.sharepoint.com", tokens: { ...profile!.tokens, sharepoint: { token: fakeJwt({ aud: SP_AUD }), exp } } });
}

function entry(url = FILE_URL, siteUrl: string | undefined = SITE) {
  return { fileName: NAME, fileInfo: { fileUrl: url, ...(siteUrl ? { siteUrl } : {}) } };
}

function fakeNet(options: { message?: unknown; email?: string; gone?: boolean; recycleStatus?: number } = {}) {
  const message = options.message ?? { id: "1700000000001", content: "<p>x</p>", properties: { files: JSON.stringify([entry()]) } };
  globalThis.fetch = (async (input: any, init: RequestInit = {}) => {
    const url = new URL(String(input));
    const method = init.method ?? "GET";
    calls.push({ method, url });
    if (url.host === "teams.cloud.microsoft" && url.pathname.endsWith("/messages/1700000000001")) return Response.json(message);
    if (url.host === HOST) {
      if (options.gone && !url.pathname.endsWith("/recycle()")) return new Response("{}", { status: 404 });
      if (url.pathname.endsWith("/recycle()")) return options.recycleStatus ? new Response("{}", { status: options.recycleStatus }) : Response.json({ value: "fake-bin-id" });
      if (url.pathname.endsWith("/Author")) return Response.json({ Title: "Test User", Email: options.email ?? "tester@example.test" });
      return Response.json({ Name: NAME, Length: "2048", TimeCreated: "2026-10-19T00:00:00Z" });
    }
    return new Response("{}", { status: 404 });
  }) as typeof fetch;
}

const posts = () => calls.filter(call => call.method === "POST");

beforeEach(async () => {
  original = globalThis.fetch;
  calls = [];
  logs = [];
  await resetLocalState();
  await seed();
  setPaceForTests({});
  console.log = (...args: unknown[]) => { logs.push(args.map(String).join(" ")); };
});
afterEach(() => {
  console.log = originalLog;
  globalThis.fetch = original;
  resetTokenCache();
  process.exitCode = 0;
  setPaceForTests({});
});

describe("serverRelativeAddress", () => {
  test("doubles ' and encodes space, #, %, ? and unicode, and keeps /", () => {
    expect(serverRelativeAddress("/sites/fake-team", `${FOLDER}/${NAME}`)).toBe(
      "/sites/fake-team/_api/web/GetFileByServerRelativePath(decodedurl='/sites/fake-team/Shared%20Documents/General/q''s%20%231%20100%25%20caf%C3%A9.png')",
    );
    expect(serverRelativeAddress("/personal/a_b", "/personal/a_b/x?y.txt")).toContain("x%3Fy.txt')");
  });
});

describe("files delete preview", () => {
  test("lists the file, its size and its uploader, and sends no POST", async () => {
    fakeNet();
    await filesDelete(REF, { allowWrite: false });
    expect(posts()).toHaveLength(0);
    expect(logs[0]).toBe("Dry run (effect: file): would move 1 file(s) to the SharePoint recycle bin:");
    expect(logs[1]).toBe(`  ${NAME}  2048 bytes, uploaded by Test User, recycle`);
    expect(logs.join("\n")).toContain("restored from the site's recycle bin");
    expect(logs.join("\n")).toContain("The Teams message is not changed");
    const read = calls.find(call => call.url.host === HOST && !call.url.pathname.endsWith("/Author"))!;
    expect(read.url.searchParams.get("$select")).toBe("Name,Length,TimeCreated");
  });

  test("the JSON preview names the effect and the count", async () => {
    fakeNet();
    await filesDelete(REF, { allowWrite: false, json: true });
    const out = JSON.parse(logs.join("\n"));
    expect(out).toMatchObject({ dry_run: true, effect: "file", count: 1 });
    expect(out.files[0]).toMatchObject({ size: 2048, uploader: "Test User", status: "recycle" });
    expect(posts()).toHaveLength(0);
  });

  test("another person's file is shown as refused", async () => {
    fakeNet({ email: "other@example.test" });
    await filesDelete(REF, { allowWrite: false });
    expect(logs.join("\n")).toContain("refused: uploaded by Test User (other@example.test), not by you");
    expect(posts()).toHaveLength(0);
  });
});

describe("files delete write", () => {
  test("POSTs recycle() on the quoted path, with no body, and prints the bin item id", async () => {
    fakeNet();
    await filesDelete(REF, { allowWrite: true, noWait: true });
    expect(posts()).toHaveLength(1);
    const post = posts()[0]!;
    expect(post.url.host).toBe(HOST);
    expect(post.url.pathname).toBe(`${serverRelativeAddress("/sites/fake-team", `${FOLDER}/${NAME}`)}/recycle()`);
    expect(post.url.pathname).toContain("q''s%20%231%20100%25%20caf%C3%A9.png");
    expect(logs[0]).toBe(`recycled ${NAME} (recycle bin item fake-bin-id)`);
    expect(logs.join("\n")).toContain("Use `delete` for the message");
  });

  test("another person's file stops the run before any POST", async () => {
    fakeNet({ email: "other@example.test" });
    await expect(filesDelete(REF, { allowWrite: true })).rejects.toThrow("Nothing was recycled");
    expect(posts()).toHaveLength(0);
  });

  test("a 404 is reported as already gone, with no POST", async () => {
    fakeNet({ gone: true });
    await filesDelete(REF, { allowWrite: true, json: true });
    expect(posts()).toHaveLength(0);
    expect(JSON.parse(logs.join("\n")).files).toEqual([{ name: NAME, status: "already gone" }]);
    expect(process.exitCode).toBe(0);
  });

  test("a 5xx on the recycle is uncertain, is not retried, and exits 1", async () => {
    fakeNet({ recycleStatus: 503 });
    await filesDelete(REF, { allowWrite: true, json: true });
    expect(posts()).toHaveLength(1);
    expect(JSON.parse(logs.join("\n")).files[0].status).toBe("uncertain");
    expect(process.exitCode).toBe(1);
  });
});

describe("files delete refusals", () => {
  test("a file outside SharePoint is refused before any SharePoint request", async () => {
    fakeNet({ message: { id: "1700000000001", properties: { files: JSON.stringify([entry("https://evil.test/sites/x/a.txt")]) } } });
    await expect(filesDelete(REF, { allowWrite: true })).rejects.toThrow("only on files on SharePoint");
    expect(calls.every(call => call.url.host === "teams.cloud.microsoft")).toBe(true);
  });

  test("a list is refused", async () => {
    fakeNet({ message: { id: "1700000000001", properties: { files: JSON.stringify([entry(`${SITE}/Lists/Env%20Tracker`)]) } } });
    await expect(filesDelete(REF, { allowWrite: true })).rejects.toThrow("is a SharePoint list");
    expect(calls.some(call => call.url.host === HOST)).toBe(false);
  });

  test("a deleted message gives the clear error", async () => {
    fakeNet({ message: { id: "1700000000001", content: "", properties: { deletetime: "1700000500000" } } });
    await expect(filesDelete(REF, { allowWrite: false })).rejects.toThrow("Run `files delete` before `delete`");
  });

  test("a message with no files is refused", async () => {
    fakeNet({ message: { id: "1700000000001", content: "<p>x</p>", properties: {} } });
    await expect(filesDelete(REF, { allowWrite: false })).rejects.toThrow("no files");
  });
});

describe("delete preview warning", () => {
  test("carries the files warning in text and JSON", async () => {
    globalThis.fetch = (async (_input: any) => Response.json({ id: "1700000000001", content: "<p>hi</p>", imdisplayname: "Test User", properties: {} })) as typeof fetch;
    await deleteMessages("me", ["1700000000001"], { json: false, allowWrite: false });
    expect(logs).toContain("Deleting a message leaves its files on SharePoint. Run `files delete <ref>` first.");
    logs = [];
    await deleteMessages("me", ["1700000000001"], { json: true, allowWrite: false });
    expect(JSON.parse(logs.join("\n")).warning).toContain("Run `files delete <ref>` first.");
  });
});
