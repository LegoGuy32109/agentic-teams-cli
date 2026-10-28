import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { mkdtempSync, writeFileSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { classifyToken, formatRemaining, identity, redact, requireToken, resetTokenCache, resolveToken } from "../src/tokens.ts";
import { configDir, credentialsPath, saveProfile } from "../src/config.ts";
import { AUD, OID, TID, fakeJwt } from "./jwt.ts";
import { seedProfile } from "./fake-teams.ts";

const ENV_KEYS = ["TEAMS_CHAT_TOKEN", "TEAMS_CSA_TOKEN", "TEAMS_SPACES_TOKEN", "TEAMS_TOKEN_COMMAND", "FAKE_TOKEN_JSON_PATH"];

beforeEach(() => {
  for (const key of ENV_KEYS) delete process.env[key];
  resetTokenCache();
});

afterEach(() => {
  for (const key of ENV_KEYS) delete process.env[key];
  resetTokenCache();
});

describe("classifyToken", () => {
  test("classifies each kind by audience", () => {
    expect(classifyToken(fakeJwt({ aud: AUD.chat })).kind).toBe("chat");
    expect(classifyToken(fakeJwt({ aud: AUD.csa })).kind).toBe("csa");
    expect(classifyToken(fakeJwt({ aud: AUD.spaces })).kind).toBe("spaces");
  });

  test("an unknown audience has no kind", () => {
    expect(classifyToken(fakeJwt({ aud: "https://graph.microsoft.com" })).kind).toBeUndefined();
  });

  test("a value that is not a JWT has no claims", () => {
    expect(classifyToken("not-a-jwt")).toEqual({});
  });
});

describe("token sources", () => {
  test("environment wins over the stored profile", async () => {
    await seedProfile(["chat"]);
    const envToken = fakeJwt({ aud: AUD.chat, exp: Math.floor(Date.now() / 1000) + 900 });
    process.env.TEAMS_CHAT_TOKEN = envToken;
    const resolved = await resolveToken("chat");
    expect(resolved?.source).toBe("environment");
    expect(resolved?.token).toBe(envToken);
  });

  test("the command wins over the stored profile and loses to the environment", async () => {
    await seedProfile(["chat", "csa"]);
    const commandToken = fakeJwt({ aud: AUD.chat });
    const file = join(mkdtempSync(join(tmpdir(), "teams-cli-tok-")), "tokens.json");
    writeFileSync(file, JSON.stringify({ chat: commandToken }));
    process.env.FAKE_TOKEN_JSON_PATH = file;
    process.env.TEAMS_TOKEN_COMMAND = `${process.execPath} ${join(import.meta.dir, "fixtures", "token-command.ts")}`;
    expect((await resolveToken("chat"))?.source).toBe("command");
    expect((await resolveToken("csa"))?.source).toBe("store");
    process.env.TEAMS_CHAT_TOKEN = fakeJwt({ aud: AUD.chat });
    expect((await resolveToken("chat"))?.source).toBe("environment");
  });

  test("a failing token command reports the exit status and no output", async () => {
    process.env.TEAMS_TOKEN_COMMAND = `${process.execPath} -e process.exit(3)`;
    await expect(resolveToken("chat")).rejects.toThrow("TEAMS_TOKEN_COMMAND exited with status 3.");
  });

  test("a token command that prints something other than JSON fails without echoing it", async () => {
    const file = join(mkdtempSync(join(tmpdir(), "teams-cli-tok-")), "bad.json");
    writeFileSync(file, "eyJhbGciOiJub25lIn0.eyJhdWQiOiJ4In0.sig");
    process.env.FAKE_TOKEN_JSON_PATH = file;
    process.env.TEAMS_TOKEN_COMMAND = `${process.execPath} ${join(import.meta.dir, "fixtures", "token-command.ts")}`;
    const error = (await resolveToken("chat").then(() => undefined, e => e as Error))!;
    expect(error.message).toBe("TEAMS_TOKEN_COMMAND must print JSON with any of the keys chat, csa, spaces.");
    expect(error.message).not.toContain("eyJ");
  });

  test("a non-JWT environment token is rejected", async () => {
    process.env.TEAMS_CSA_TOKEN = "plain-text";
    await expect(resolveToken("csa")).rejects.toThrow("TEAMS_CSA_TOKEN is not a JWT.");
  });
});

describe("expiry", () => {
  test("a token that expires in under 60 seconds counts as expired", async () => {
    await seedProfile(["chat"], Math.floor(Date.now() / 1000) + 30);
    await expect(requireToken("chat")).rejects.toThrow("The chat token expires in under a minute. Supply a new one: teams-cli auth");
  });

  test("a fresh token passes", async () => {
    await seedProfile(["chat"]);
    expect((await requireToken("chat")).kind).toBe("chat");
  });

  test("a missing kind says how to fix it", async () => {
    await seedProfile(["chat"]);
    await expect(requireToken("csa")).rejects.toThrow("No csa token. Run: teams-cli auth");
  });

  test("remaining time formats as hours and minutes", () => {
    expect(formatRemaining(21 * 3600_000 + 3 * 60_000)).toBe("21h 3m");
    expect(formatRemaining(42 * 60_000)).toBe("42m");
  });
});

describe("identity and redaction", () => {
  test("identity comes from the stored profile", async () => {
    await seedProfile(["chat"]);
    const who = await identity();
    expect(who.mri).toBe(`8:orgid:${OID}`);
    expect(who.tenantId).toBe(TID);
    expect(who.region).toBe("amer");
  });

  test("identity falls back to token claims when no profile exists", async () => {
    process.env.TEAMS_CHAT_TOKEN = fakeJwt({ aud: AUD.chat });
    expect((await identity()).oid).toBe(OID);
  });

  test("redact removes tokens and bearer values", () => {
    const token = fakeJwt({ aud: AUD.chat });
    expect(redact(`failed with ${token}`)).toBe("failed with <token>");
    expect(redact("authorization: Bearer abcdefghijklmnop.qrstuvwxyz-0123")).toBe("authorization: Bearer <token>");
  });
});

describe("credentials file", () => {
  test("the credentials file is written with mode 0600", async () => {
    await saveProfile("default", {
      tokens: {}, userId: OID, tenantId: TID, upn: "", name: "", region: "amer", updatedAt: "",
    });
    expect(credentialsPath().startsWith(configDir())).toBe(true);
    expect(statSync(credentialsPath()).mode & 0o777).toBe(0o600);
  });
});

describe("audience checks for supplied tokens", () => {
  afterEach(() => {
    delete process.env.TEAMS_CHAT_TOKEN;
    delete process.env.TEAMS_CSA_TOKEN;
    resetTokenCache();
  });

  test("an environment token with another service's audience fails before any request", async () => {
    process.env.TEAMS_CHAT_TOKEN = fakeJwt({ aud: AUD.csa, exp: Math.floor(Date.now() / 1000) + 3600 });
    await expect(resolveToken("chat")).rejects.toThrow(
      "The token given as chat is for the csa service. Supply the right token: teams-cli auth",
    );
  });

  test("a matching audience is accepted", async () => {
    process.env.TEAMS_CSA_TOKEN = fakeJwt({ aud: AUD.csa, exp: Math.floor(Date.now() / 1000) + 3600 });
    expect((await resolveToken("csa"))?.source).toBe("environment");
  });

  test("an unknown audience passes, since the service decides", async () => {
    process.env.TEAMS_CHAT_TOKEN = fakeJwt({ aud: "https://example.test", exp: Math.floor(Date.now() / 1000) + 3600 });
    expect((await resolveToken("chat"))?.source).toBe("environment");
  });
});
