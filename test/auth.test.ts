import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readFileSync, statSync } from "node:fs";
import { authStatus } from "../src/commands/auth.ts";
import { parsePaste } from "../src/commands/auth.ts";
import { loadProfile } from "../src/config.ts";
import { resetTokenCache } from "../src/tokens.ts";
import { AUD, fakeJwt, OID, TID } from "./jwt.ts";
import { installFakeTeams, seedProfile } from "./fake-teams.ts";
import { runCli } from "./cli.ts";

const chat = fakeJwt({ aud: AUD.chat });
const csa = fakeJwt({ aud: AUD.csa });
const spaces = fakeJwt({ aud: AUD.spaces });
const curl = (token: string, path: string) =>
  `curl 'https://teams.cloud.microsoft/api/${path}/amer/v1/x' \\\n  -H 'authorization: Bearer ${token}' \\\n  -H 'accept: */*'`;

describe("parsePaste", () => {
  test("reads three cURL commands in any order", () => {
    const paste = [curl(spaces, "mt"), curl(chat, "chatsvc"), curl(csa, "csa")].join("\n\n");
    const parsed = parsePaste(paste);
    expect(Object.keys(parsed.tokens).sort()).toEqual(["chat", "csa", "spaces"]);
    expect(parsed.region).toBe("amer");
    expect(parsed.ignored).toBe(0);
  });

  test("reads one cURL command", () => {
    const parsed = parsePaste(curl(chat, "chatsvc"));
    expect(Object.keys(parsed.tokens)).toEqual(["chat"]);
  });

  test("handles angle brackets around URLs and backslash line breaks", () => {
    const paste = `curl '<https://teams.cloud.microsoft/api/csa/amer/api/v3>' \\\n -H "Authorization: Bearer ${csa}"`;
    const parsed = parsePaste(paste);
    expect(parsed.tokens.csa?.token).toBe(csa);
    expect(parsed.region).toBe("amer");
  });

  test("an unknown audience is ignored and counted", () => {
    const parsed = parsePaste(curl(fakeJwt({ aud: "https://graph.microsoft.com" }), "chatsvc"));
    expect(parsed.tokens).toEqual({});
    expect(parsed.ignored).toBe(1);
  });

  test("two tokens of one kind keep the one that expires later", () => {
    const early = fakeJwt({ aud: AUD.chat, exp: 1_900_000_000 });
    const late = fakeJwt({ aud: AUD.chat, exp: 1_900_003_600 });
    expect(parsePaste(`${curl(early, "chatsvc")}\n${curl(late, "chatsvc")}`).tokens.chat?.token).toBe(late);
  });

  test("the region comes from the URL path", () => {
    expect(parsePaste(curl(chat, "chatsvc")).region).toBe("amer");
    expect(parsePaste("https://teams.cloud.microsoft/api/mt/emea/beta/x").region).toBe("emea");
  });
});

describe("auth command", () => {
  test("a paste of three cURLs stores all kinds and prints no token", async () => {
    const paste = [curl(chat, "chatsvc"), curl(csa, "csa"), curl(spaces, "mt")].join("\n");
    const result = await runCli(["auth"], { input: paste });
    expect(result.code).toBe(0);
    expect(result.stdout).toContain("chat: ok, expires in");
    expect(result.stdout).toContain("spaces: ok, expires in");
    expect(result.stdout).toContain("csa unlocks: conversations");
    for (const token of [chat, csa, spaces]) expect(result.stdout + result.stderr).not.toContain(token);
    const stored = JSON.parse(readFileSync(`${result.home}/config/teams-cli/credentials.json`, "utf8"));
    expect(stored.profiles.default.tokens.chat.token).toBe(chat);
  });

  test("the stored credentials file has mode 0600", async () => {
    const result = await runCli(["auth"], { input: curl(chat, "chatsvc") });
    expect(result.code).toBe(0);
    const path = `${result.home}/config/teams-cli/credentials.json`;
    expect(statSync(path).mode & 0o777).toBe(0o600);
  });

  test("a second paste replaces only the kinds it contains", async () => {
    const home = (await runCli(["auth"], { input: [curl(chat, "chatsvc"), curl(csa, "csa")].join("\n") })).home;
    const second = fakeJwt({ aud: AUD.chat, exp: Math.floor(Date.now() / 1000) + 7200 });
    const result = await runCli(["auth"], { input: curl(second, "chatsvc"), home });
    expect(result.code).toBe(0);
    const stored = JSON.parse(readFileSync(`${home}/config/teams-cli/credentials.json`, "utf8"));
    expect(stored.profiles.default.tokens.chat.token).toBe(second);
    expect(stored.profiles.default.tokens.csa.token).toBe(csa);
  });

  test("a paste with no bearer token fails", async () => {
    const result = await runCli(["auth"], { input: "curl https://example.test" });
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("No authorization token found in the pasted text.");
  });

  test("a paste with only an unknown audience fails and says so", async () => {
    const result = await runCli(["auth"], { input: curl(fakeJwt({ aud: "https://graph.microsoft.com" }), "chatsvc") });
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("Found 1 token(s), but none for a Teams service. Audience(s): https://graph.microsoft.com.");
    expect(result.stderr).toContain("https://substrate.office.com");
  });

  test("logout removes the profile and a second logout reports none", async () => {
    const home = (await runCli(["auth"], { input: curl(chat, "chatsvc") })).home;
    const first = await runCli(["logout"], { home });
    expect(first.stdout.trim()).toBe("Profile removed.");
    const second = await runCli(["logout"], { home });
    expect(second.stdout.trim()).toBe("No matching stored profile.");
  });
});

describe("auth status", () => {
  let fake: ReturnType<typeof installFakeTeams> | undefined;
  beforeEach(() => resetTokenCache());
  afterEach(() => { fake?.uninstall(); fake = undefined; });

  test("prints each kind and expects online ok", async () => {
    await seedProfile();
    fake = installFakeTeams({ "GET /api/chatsvc/amer/v1/users/ME/properties": { status: 200, json: {} } });
    const lines: string[] = [];
    const original = console.log;
    console.log = (...args: unknown[]) => { lines.push(args.map(String).join(" ")); };
    try {
      await authStatus();
    } finally {
      console.log = original;
    }
    expect(lines).toContain(`tenant: ${TID}`);
    expect(lines.some(line => line.startsWith("chat: ok, expires in"))).toBe(true);
    expect(lines).toContain("online: ok");
    expect(fake.requests.length).toBe(1);
    expect(lines.join("\n")).not.toContain("eyJ");
    expect((await loadProfile()).profile?.userId).toBe(OID);
  });

  test("a stored profile without a chat token reports the failure", async () => {
    await seedProfile(["csa"]);
    fake = installFakeTeams({});
    const lines: string[] = [];
    const original = console.log;
    console.log = (...args: unknown[]) => { lines.push(args.map(String).join(" ")); };
    try {
      await authStatus();
    } finally {
      console.log = original;
    }
    expect(lines).toContain("chat: missing");
    expect(lines.some(line => line.startsWith("online: No chat token."))).toBe(true);
    expect(process.exitCode).toBe(1);
    process.exitCode = 0;
  });
});
