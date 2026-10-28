import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { mkdir, readFile, rm, stat, symlink, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { authLogin, authRenew, browserStatusLine, importCaptured, renewFromBrowser, setBrowserTestHooks, SIGN_IN_AGAIN, SignInRequiredError } from "../src/browser-auth.ts";
import { browserProfileDir, prepareProfileDir } from "../src/browser.ts";
import { authStatus } from "../src/commands/auth.ts";
import { credentialsPath, loadProfile } from "../src/config.ts";
import { refreshToken, setupRefresh } from "../src/refresh.ts";
import { resetTokenCache } from "../src/tokens.ts";
import { installFakeTeams, resetLocalState, seedProfile } from "./fake-teams.ts";
import { runCli } from "./cli.ts";
import { body, startFakeCdp, tokenBody, tokenExchange, type FakeCdpOptions, type FakeEvent } from "./fake-cdp.ts";
import { AUD, fakeJwt, OID, TID } from "./jwt.ts";

const FORM = { client_id: "fake-client", grant_type: "refresh_token", refresh_token: "fake-old", scope: "x" };
const LOGIN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=fake";
const TIMING = { pollMs: 10, settleMs: 20, graceMs: 120, scanAfterMs: 150, nudgeMs: 60, renewTimeoutMs: 700, loginTimeoutMs: 700 };

let fake: ReturnType<typeof startFakeCdp> | undefined;
let launches: Array<{ executable: string; userDataDir: string; headless: boolean }>;
let closes: number;
let originalFetch: typeof fetch;

/** Starts a fake browser and points the launcher at it. */
function useBrowser(options: FakeCdpOptions, timing: Partial<typeof TIMING> = {}) {
  fake = startFakeCdp(options);
  setBrowserTestHooks({
    executable: "/fake/chromium",
    timing: { ...TIMING, ...timing },
    launcher: async launch => { launches.push(launch); return { wsUrl: fake!.wsUrl, close: async () => { closes++; } }; },
  });
  return fake;
}

/** Events and bodies for token responses, by audience. */
function scenario(kinds: Array<[string, string]>, extra: Partial<FakeCdpOptions> = {}, claims: Record<string, unknown> = {}): FakeCdpOptions {
  const bodies: FakeCdpOptions["bodies"] = {};
  const events: FakeEvent[] = [];
  kinds.forEach(([aud, refresh], index) => {
    bodies[`r${index}`] = tokenBody(aud, { refresh_token: refresh }, claims);
    events.push(...tokenExchange(`r${index}`, TID, FORM));
  });
  return { bodies, onNavigate: () => events, ...extra };
}

const CORE_THREE: Array<[string, string]> = [[AUD.chat, "fake-refresh-1"], [AUD.csa, "fake-refresh-2"], [AUD.spaces, "fake-refresh-3"]];

async function makeBrowserProfile() {
  const dir = browserProfileDir("default");
  await prepareProfileDir(dir);
  await writeFile(join(dir, "Local State"), "{}");
  return dir;
}

async function captureOutput<T>(work: () => Promise<T>): Promise<{ out: string[]; err: string[]; result?: T; error?: Error }> {
  const out: string[] = [], err: string[] = [];
  const log = console.log, error = console.error;
  console.log = (...args: unknown[]) => { out.push(args.map(String).join(" ")); };
  console.error = (...args: unknown[]) => { err.push(args.map(String).join(" ")); };
  try { return { out, err, result: await work() }; }
  catch (caught) { return { out, err, error: caught as Error }; }
  finally { console.log = log; console.error = error; process.exitCode = 0; }
}

beforeEach(async () => {
  originalFetch = globalThis.fetch;
  launches = [];
  closes = 0;
  await resetLocalState();
});
afterEach(() => {
  fake?.stop();
  fake = undefined;
  setBrowserTestHooks();
  globalThis.fetch = originalFetch;
  resetTokenCache();
  process.exitCode = 0;
});

describe("renew", () => {
  beforeEach(async () => {
    await seedProfile(["chat"], Math.floor(Date.now() / 1000) - 60);
    await makeBrowserProfile();
  });

  test("a headless run captures the tokens, stores the newest refresh token and closes the browser", async () => {
    const browser = useBrowser(scenario([...CORE_THREE, ["https://substrate.office.com", "fake-refresh-4"]]));
    const imported = await renewFromBrowser();
    const { profile } = await loadProfile();
    expect(Object.keys(profile!.tokens).sort()).toEqual(["chat", "csa", "spaces", "substrate"]);
    expect(profile!.refresh).toEqual({ clientId: "fake-client", tenantId: TID, refreshToken: "fake-refresh-4", origin: "https://teams.cloud.microsoft" });
    expect(profile!.browser?.lastMethod).toBe("renew");
    expect(imported.refreshCaptured).toBe(true);
    expect(imported.notCaptured).toEqual([]);
    expect(launches).toEqual([{ executable: "/fake/chromium", userDataDir: browserProfileDir("default"), headless: true }]);
    expect(browser.commands.find(command => command.method === "Page.navigate")?.params.url).toBe("https://teams.cloud.microsoft/");
    expect(browser.closedByCommand).toBe(true);
    expect(closes).toBe(1);
    expect((await stat(credentialsPath())).mode & 0o777).toBe(0o600);
    expect((await stat(browserProfileDir("default"))).mode & 0o777).toBe(0o700);
  });

  test("the headless user agent is replaced with the ordinary one", async () => {
    const browser = useBrowser(scenario(CORE_THREE));
    await renewFromBrowser();
    const override = browser.commands.find(command => command.method === "Network.setUserAgentOverride");
    expect(override?.params.userAgent).toContain("Chrome/154.0.0.0");
    expect(override?.params.userAgent).not.toContain("Headless");
    expect(override?.sessionId).toBe("S1");
  });

  test("a missing search token is nudged once and reported, not treated as a failure", async () => {
    const browser = useBrowser(scenario(CORE_THREE));
    const imported = await renewFromBrowser();
    expect(imported.notCaptured).toEqual(["substrate"]);
    expect(browser.commands.filter(command => command.method === "Runtime.evaluate" && String(command.params.expression).includes("AUTOSUGGEST_INPUT"))).toHaveLength(1);
  });

  test("a page that stays on the Microsoft sign-in page fails with the login command before the time limit", async () => {
    useBrowser({ onNavigate: () => [{ method: "Page.frameNavigated", params: { frame: { id: "F1", url: LOGIN_URL } } }] }, { renewTimeoutMs: 2_000 });
    const started = Date.now();
    const error = await renewFromBrowser().then(() => undefined, e => e as Error);
    expect(error).toBeInstanceOf(SignInRequiredError);
    expect(error?.message).toBe("Microsoft needs you to sign in again. Run: teams-cli auth login");
    expect(Date.now() - started).toBeLessThan(1_500);
    expect(closes).toBe(1);
  });

  test("a page on the sign-in page when time runs out is a sign-in failure, not a timeout", async () => {
    useBrowser({ onNavigate: () => [{ method: "Page.frameNavigated", params: { frame: { id: "F1", url: LOGIN_URL } } }] }, { graceMs: 5_000, renewTimeoutMs: 300 });
    await expect(renewFromBrowser()).rejects.toThrow(SIGN_IN_AGAIN);
  });

  test("a quick pass through the login host does not count as a sign-in page", async () => {
    useBrowser(scenario(CORE_THREE, {
      onNavigate: () => [
        { method: "Page.frameNavigated", params: { frame: { id: "F1", url: LOGIN_URL } } },
        { method: "Page.frameNavigated", params: { frame: { id: "F1", url: "https://teams.cloud.microsoft/v2/" } }, delayMs: 20 },
        ...scenario(CORE_THREE).onNavigate!(""),
      ],
    }));
    expect((await renewFromBrowser()).notCaptured).toEqual(["substrate"]);
  });

  test("a sub-frame on the login host is not a sign-in page", async () => {
    useBrowser(scenario(CORE_THREE, {
      onNavigate: () => [
        { method: "Page.frameNavigated", params: { frame: { id: "F2", parentId: "F1", url: LOGIN_URL } } },
        ...scenario(CORE_THREE).onNavigate!(""),
      ],
    }));
    await renewFromBrowser();
  });

  test("a token request from an iframe or worker is read through its own session", async () => {
    const child = scenario(CORE_THREE);
    const browser = useBrowser({
      ...child,
      onNavigate: () => [
        { method: "Target.attachedToTarget", params: { sessionId: "S2", targetInfo: { type: "iframe" } } },
        ...child.onNavigate!("").map(event => ({ ...event, sessionId: "S2" })),
      ],
    });
    expect((await renewFromBrowser()).notCaptured).toEqual(["substrate"]);
    expect(browser.commands.some(command => command.method === "Network.enable" && command.sessionId === "S2")).toBe(true);
    expect(browser.commands.filter(command => command.method === "Network.getResponseBody").every(command => command.sessionId === "S2")).toBe(true);
  });

  test("events from a session that was never attached are ignored", async () => {
    const stray = scenario(CORE_THREE);
    useBrowser({ ...stray, onNavigate: () => stray.onNavigate!("").map(event => ({ ...event, sessionId: "S9" })) });
    await expect(renewFromBrowser()).rejects.toThrow("Teams issued no tokens");
  });

  test("a run that captures nothing and is not on a sign-in page says so", async () => {
    useBrowser({});
    await expect(renewFromBrowser()).rejects.toThrow("Teams issued no tokens within 1 seconds.");
  });

  test("tokens from the page's own storage are used when the page made no token request", async () => {
    useBrowser({ storage: [fakeJwt({ aud: AUD.chat }), fakeJwt({ aud: AUD.csa }), fakeJwt({ aud: AUD.spaces }), "not a jwt", fakeJwt({ aud: "https://graph.microsoft.com" })] });
    const imported = await renewFromBrowser();
    expect(imported.notCaptured).toEqual(["substrate"]);
    expect(imported.refreshCaptured).toBe(false);
  });

  test("a closed browser window ends the run with a plain message", async () => {
    const browser = useBrowser({ onNavigate: () => [] });
    setTimeout(() => browser.dropConnections(), 80);
    await expect(renewFromBrowser()).rejects.toThrow("The browser closed before Teams issued tokens.");
  });

  test("a token for another account is refused and the stored profile is unchanged", async () => {
    useBrowser(scenario(CORE_THREE, {}, { oid: "another-user" }));
    const before = JSON.stringify((await loadProfile()).profile);
    await expect(renewFromBrowser()).rejects.toThrow("another account");
    expect(JSON.stringify((await loadProfile()).profile)).toBe(before);
  });

  test("a profile that a running browser holds is not launched", async () => {
    useBrowser(scenario(CORE_THREE));
    await symlink(`host-${process.pid}`, join(browserProfileDir("default"), "SingletonLock"));
    await expect(renewFromBrowser()).rejects.toThrow("open in a running browser");
    expect(launches).toEqual([]);
  });

  test("a launch failure is an error, and it does not leave a lock", async () => {
    setBrowserTestHooks({ executable: "/fake/chromium", launcher: async () => { throw new Error("Could not start the browser: /fake/chromium"); } });
    await expect(authRenew()).rejects.toThrow("Could not start the browser");
    await expect(stat(`${credentialsPath()}.refresh.lock`)).rejects.toThrow();
  });

  test("renew without a browser profile asks for a login and starts no browser", async () => {
    await resetLocalState();
    await seedProfile(["chat"]);
    useBrowser(scenario(CORE_THREE));
    await expect(authRenew()).rejects.toThrow("No browser profile. Run: teams-cli auth login");
    expect(launches).toEqual([]);
  });

  test("auth renew prints one line per kind and never a token", async () => {
    useBrowser(scenario([...CORE_THREE, ["https://substrate.office.com", "fake-refresh-4"]]));
    const { out, err } = await captureOutput(() => authRenew());
    expect(out).toContain("refresh: captured, good for about 24 hours");
    expect(out.filter(line => /^(chat|csa|spaces|substrate): captured, expires in/.test(line))).toHaveLength(4);
    expect(out.join("\n") + err.join("\n")).not.toMatch(/eyJ|fake-refresh/);
    expect(process.exitCode ?? 0).toBe(0);
  });

  test("auth renew exits 1 and names the core kinds that never appeared", async () => {
    useBrowser(scenario([[AUD.chat, "fake-refresh-1"]]));
    let code: number | string | null | undefined;
    const { out, err } = await captureOutput(async () => { await authRenew(); code = process.exitCode; });
    expect(code).toBe(1);
    expect(out).toContain("csa: not captured");
    expect(err.join("\n")).toContain("Teams did not issue these tokens: csa, spaces.");
  });
});

describe("login", () => {
  test("a visible run creates the profile when none exists and prints no token", async () => {
    useBrowser(scenario([...CORE_THREE, ["https://substrate.office.com", "fake-refresh-4"]], {}, { upn: "josh@example.test" }));
    const { out, err } = await captureOutput(() => authLogin({ browser: "/opt/chromium" }));
    expect(launches).toEqual([{ executable: "/fake/chromium", userDataDir: browserProfileDir("default"), headless: false }]);
    const { profile } = await loadProfile();
    expect(profile).toMatchObject({ userId: OID, tenantId: TID, upn: "josh@example.test", region: "amer", browser: { lastMethod: "login" } });
    expect(profile!.refresh?.refreshToken).toBe("fake-refresh-4");
    expect(out.join("\n") + err.join("\n")).not.toMatch(/eyJ|fake-refresh/);
    expect(out).toContain("Profile: default");
    expect(err.join("\n")).toContain(`Browser profile: ${browserProfileDir("default")}`);
    expect((await stat(browserProfileDir("default"))).isDirectory()).toBe(true);
  });

  test("a visible run waits for the person and does not apply the sign-in page rule", async () => {
    // The person is on the sign-in page for longer than the grace period, then the tokens arrive.
    const tokens = scenario(CORE_THREE).onNavigate!("");
    useBrowser({
      ...scenario(CORE_THREE),
      onNavigate: () => [
        { method: "Page.frameNavigated", params: { frame: { id: "F1", url: LOGIN_URL } } },
        ...tokens.map((event, index) => index === 0 ? { ...event, delayMs: 300 } : event),
      ],
    });
    await captureOutput(() => authLogin());
    expect((await loadProfile()).profile?.tokens.chat).toBeDefined();
  });

  test("the login import waits for a renewal that holds the lock", async () => {
    useBrowser(scenario(CORE_THREE));
    const lock = `${credentialsPath()}.refresh.lock`;
    await mkdir(dirname(lock), { recursive: true });
    await writeFile(lock, "");
    const running = captureOutput(() => authLogin());
    await Bun.sleep(250);
    expect((await loadProfile()).profile).toBeUndefined();
    await rm(lock);
    await running;
    expect((await loadProfile()).profile?.tokens.chat).toBeDefined();
  });

  test("a login that times out says so and saves nothing", async () => {
    useBrowser({});
    const { error } = await captureOutput(() => authLogin());
    expect(error?.message).toContain("Sign in, then run the command again.");
    expect((await loadProfile()).profile).toBeUndefined();
    expect(closes).toBe(1);
  });
});

describe("import", () => {
  test("a capture with no token or no identity is refused", async () => {
    await expect(importCaptured({ tokens: {} }, "login")).rejects.toThrow("no Teams tokens");
    await expect(importCaptured({ tokens: { chat: { token: fakeJwt(), exp: 1 } } }, "login")).rejects.toThrow("no Teams tokens");
  });

  test("a refresh token for another tenant is not stored", async () => {
    await seedProfile(["chat"]);
    await importCaptured({
      tokens: { chat: { token: fakeJwt(), exp: Math.floor(Date.now() / 1000) + 3600 } },
      claims: { oid: OID, tid: TID },
      refresh: { clientId: "c", tenantId: "other-tenant", refreshToken: "fake-x", origin: "https://teams.cloud.microsoft" },
    }, "renew");
    expect((await loadProfile()).profile?.refresh).toBeUndefined();
  });
});

describe("status line", () => {
  test("it says none, then present with the last method", async () => {
    await seedProfile(["chat"]);
    expect(await browserStatusLine("default", (await loadProfile()).profile)).toBe("browser profile: none. Run: teams-cli auth login");
    await makeBrowserProfile();
    expect(await browserStatusLine("default", (await loadProfile()).profile)).toBe("browser profile: present, no token import yet");
    useBrowser(scenario(CORE_THREE));
    await renewFromBrowser();
    expect(await browserStatusLine("default", (await loadProfile()).profile)).toMatch(/^browser profile: present, last renewal 20\d\d-/);
  });
});

describe("auth status", () => {
  test("it prints the browser profile line after the token lines", async () => {
    await seedProfile();
    await makeBrowserProfile();
    const fakeTeams = installFakeTeams({ "GET /api/chatsvc/amer/v1/users/ME/properties": { status: 200, json: {} } });
    try {
      const { out } = await captureOutput(() => authStatus());
      expect(out).toContain("browser profile: present, no token import yet");
      expect(out.indexOf("browser profile: present, no token import yet")).toBeGreaterThan(out.findIndex(line => line.startsWith("spaces:")));
    } finally { fakeTeams.uninstall(); }
  });
});

describe("automatic renewal when Microsoft refuses the refresh token", () => {
  const refused = () => Response.json({ error: "invalid_grant", error_description: "fake-refresh-secret" }, { status: 400 });

  beforeEach(async () => {
    await seedProfile(["chat"], Math.floor(Date.now() / 1000) - 60);
    await setupRefresh(JSON.stringify({ clientId: "fake-client", tenantId: TID, refreshToken: "fake-expired-secret", origin: "https://teams.cloud.microsoft" }));
  });

  function fakeMicrosoft(replies: Array<(form: URLSearchParams) => Response>) {
    const forms: URLSearchParams[] = [];
    globalThis.fetch = (async (input: any, init: RequestInit) => {
      expect(new URL(String(input)).host).toBe("login.microsoftonline.com");
      const form = new URLSearchParams(String(init.body));
      forms.push(form);
      return replies[Math.min(forms.length - 1, replies.length - 1)]!(form);
    }) as typeof fetch;
    return forms;
  }

  test("one browser renewal, then the original renewal runs again with the fresh refresh token", async () => {
    await makeBrowserProfile();
    useBrowser(scenario(CORE_THREE));
    const forms = fakeMicrosoft([refused, () => Response.json({ access_token: fakeJwt({ aud: AUD.chat }), refresh_token: "fake-rotated-after" })]);
    await refreshToken("chat", true);
    expect(launches).toHaveLength(1);
    expect(launches[0]!.headless).toBe(true);
    expect(forms.map(form => form.get("refresh_token"))).toEqual(["fake-expired-secret", "fake-refresh-3"]);
    expect((await loadProfile()).profile?.refresh?.refreshToken).toBe("fake-rotated-after");
  });

  test("when the browser already supplied a valid token, no second request is needed", async () => {
    await makeBrowserProfile();
    useBrowser(scenario(CORE_THREE));
    const forms = fakeMicrosoft([refused]);
    await refreshToken("chat");
    expect(forms).toHaveLength(1);
    expect((await loadProfile()).profile!.tokens.chat!.exp * 1000).toBeGreaterThan(Date.now());
  });

  test("a second refusal after the browser renewal stops without another browser", async () => {
    await makeBrowserProfile();
    useBrowser(scenario(CORE_THREE));
    fakeMicrosoft([refused]);
    const error = await refreshToken("chat", true).then(() => undefined, e => e as Error);
    expect(error?.message).toContain("Run: teams-cli auth login");
    expect(error?.message).not.toContain("fake-refresh-secret");
    expect(launches).toHaveLength(1);
  });

  test("with no browser profile the error keeps its meaning and points to auth login", async () => {
    useBrowser(scenario(CORE_THREE));
    fakeMicrosoft([refused]);
    const error = await refreshToken("chat").then(() => undefined, e => e as Error);
    expect(error?.message).toBe("Microsoft requires a new sign-in. Run: teams-cli auth login");
    expect(launches).toEqual([]);
  });

  test("a browser that needs a person gives the sign-in message", async () => {
    await makeBrowserProfile();
    useBrowser({ onNavigate: () => [{ method: "Page.frameNavigated", params: { frame: { id: "F1", url: LOGIN_URL } } }] });
    fakeMicrosoft([refused]);
    await expect(refreshToken("chat")).rejects.toThrow(SIGN_IN_AGAIN);
  });

  test("a browser that fails for another reason is named in the error", async () => {
    await makeBrowserProfile();
    setBrowserTestHooks({ executable: "/fake/chromium", launcher: async () => { throw new Error("Could not start the browser: /fake/chromium"); } });
    fakeMicrosoft([refused]);
    const error = await refreshToken("chat").then(() => undefined, e => e as Error);
    expect(error?.message).toContain("The browser renewal failed: Could not start the browser");
    expect(error?.message).toContain("Run: teams-cli auth login");
  });

  test("a different Microsoft error does not start a browser", async () => {
    await makeBrowserProfile();
    useBrowser(scenario(CORE_THREE));
    fakeMicrosoft([() => Response.json({ error: "invalid_scope" }, { status: 400 })]);
    await expect(refreshToken("chat")).rejects.toThrow("refused chat token renewal (400)");
    expect(launches).toEqual([]);
  });

  test("two renewals at once start one browser", async () => {
    await makeBrowserProfile();
    useBrowser(scenario(CORE_THREE));
    fakeMicrosoft([refused]);
    await Promise.all([refreshToken("chat"), refreshToken("chat")]);
    expect(launches).toHaveLength(1);
  });

  test("a renewal waits for the lock that a running browser renewal holds", async () => {
    await makeBrowserProfile();
    useBrowser(scenario(CORE_THREE, { onNavigate: () => scenario(CORE_THREE).onNavigate!("").map((event, index) => index === 0 ? { ...event, delayMs: 200 } : event) }));
    fakeMicrosoft([refused]);
    const running = captureOutput(() => authRenew());
    await Bun.sleep(50);
    const started = Date.now();
    await refreshToken("chat");
    await running;
    expect(Date.now() - started).toBeGreaterThan(100);
    expect(launches).toHaveLength(1);
  });
});

describe("command line", () => {
  test("auth login with a browser that does not exist fails plainly", async () => {
    const result = await runCli(["auth", "login", "--browser", "/nonexistent/browser"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("Could not start the browser: /nonexistent/browser");
  });

  test("auth renew with no browser profile asks for a login", async () => {
    const result = await runCli(["auth", "renew"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("No browser profile. Run: teams-cli auth login");
  });

  test("--browser is only valid for auth login and auth renew", async () => {
    const wrong = await runCli(["auth", "status", "--browser", "x"]);
    expect(wrong.stderr).toContain("Flag `--browser` is not valid");
    const other = await runCli(["history", "me", "--browser", "x"]);
    expect(other.stderr).toContain("Flag `--browser` is not valid");
    const extra = await runCli(["auth", "login", "extra"]);
    expect(extra.stderr).toContain("Usage: teams-cli auth login [--browser PATH]");
    const missing = await runCli(["auth", "login", "--browser"]);
    expect(missing.stderr).toContain("--browser requires a value.");
  });

  test("a pasted token import keeps the browser profile record", async () => {
    const chat = fakeJwt({ aud: AUD.chat });
    const paste = `curl 'https://teams.cloud.microsoft/api/chatsvc/amer/v1/x' -H 'authorization: Bearer ${chat}'`;
    const first = await runCli(["auth"], { input: paste });
    const path = `${first.home}/config/teams-cli/credentials.json`;
    const stored = JSON.parse(await readFile(path, "utf8"));
    stored.profiles.default.browser = { lastAt: "2026-10-16T00:00:00.000Z", lastMethod: "login" };
    await writeFile(path, JSON.stringify(stored));
    await runCli(["auth"], { input: paste, home: first.home });
    expect(JSON.parse(await readFile(path, "utf8")).profiles.default.browser).toEqual({ lastAt: "2026-10-16T00:00:00.000Z", lastMethod: "login" });
  });

  test("the help text lists both commands", async () => {
    const result = await runCli(["--help"]);
    expect(result.stdout).toContain("auth login [--browser PATH]");
    expect(result.stdout).toContain("auth renew [--browser PATH]");
  });
});
