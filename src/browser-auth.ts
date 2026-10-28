import { CdpClient, type CdpEvent } from "./cdp.ts";
import { SignInWatch, TokenCapture, type Captured } from "./capture.ts";
import { browserProfileDir, browserProfileExists, findBrowser, launchBrowser, prepareProfileDir, profileInUse, type Launcher } from "./browser.ts";
import { loadProfile, saveProfile, type Profile, type TokenKind } from "./config.ts";
import { locked } from "./lock.ts";
import { describeKind, KINDS, redact } from "./tokens.ts";

/**
 * Sign-in through a dedicated browser profile. `login` opens a visible browser for a person.
 * `renew` runs the same profile headless and lets Teams sign in silently from its session cookies.
 * Tokens come from the page's own requests to the Microsoft token endpoint.
 */

const TEAMS_URL = "https://teams.cloud.microsoft/";
const CORE: TokenKind[] = ["chat", "csa", "spaces"];
/** The browser never asks SharePoint for a token. That one renews from the refresh token when a command needs it. */
const BROWSER_KINDS = KINDS.filter(kind => kind !== "sharepoint");

export const SIGN_IN_AGAIN = "Microsoft needs you to sign in again. Run: teams-cli auth login";

export class SignInRequiredError extends Error {
  constructor() { super(SIGN_IN_AGAIN); }
}

export type BrowserMode = "login" | "renew";
export type BrowserOptions = { browser?: string };

type Timing = {
  pollMs: number; settleMs: number; graceMs: number; scanAfterMs: number; nudgeMs: number;
  renewTimeoutMs: number; loginTimeoutMs: number;
};
const DEFAULT_TIMING: Timing = {
  pollMs: 250, settleMs: 2_000, graceMs: 8_000, scanAfterMs: 25_000, nudgeMs: 12_000,
  renewTimeoutMs: 90_000, loginTimeoutMs: 600_000,
};

type Hooks = { launcher?: Launcher; timing?: Partial<Timing>; url?: string; executable?: string };
let hooks: Hooks = {};
/** Tests replace the browser launcher, the timing, the start page, and the browser lookup. */
export function setBrowserTestHooks(next: Hooks = {}): void { hooks = next; }

export type RunResult = { captured: Captured; method: BrowserMode };

/** Finds access tokens that the Teams page keeps in its own local storage. Runs inside the page. */
const STORAGE_SCAN = `(() => {
  const out = [];
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const raw = localStorage.getItem(localStorage.key(i));
      if (!raw) continue;
      if (raw.startsWith("eyJ")) { out.push(raw); continue; }
      try { const value = JSON.parse(raw); if (value && typeof value.secret === "string") out.push(value.secret); } catch {}
    }
  } catch {}
  return out;
})()`;

/** Focuses the Teams search box, which makes the page ask for a search token. The selectors are a guess. */
const SEARCH_NUDGE = `(() => {
  const box = document.querySelector('[data-tid="AUTOSUGGEST_INPUT"], input[type="search"], input[placeholder*="Search" i]');
  if (!box) return false;
  box.focus(); box.click();
  return true;
})()`;

async function firstPage(cdp: CdpClient): Promise<string> {
  const deadline = Date.now() + 5_000;
  for (;;) {
    const { targetInfos } = await cdp.send<{ targetInfos: Array<{ targetId: string; type: string }> }>("Target.getTargets");
    const page = targetInfos?.find(info => info.type === "page");
    if (page) return page.targetId;
    if (Date.now() >= deadline) break;
    await Bun.sleep(100);
  }
  return (await cdp.send<{ targetId: string }>("Target.createTarget", { url: "about:blank" })).targetId;
}

async function drive(cdp: CdpClient, mode: BrowserMode, timing: Timing): Promise<Captured> {
  const { sessionId } = await cdp.send<{ sessionId: string }>("Target.attachToTarget", { targetId: await firstPage(cdp), flatten: true });
  const sessions = new Set([sessionId]);
  const capture = new TokenCapture({
    getResponseBody: (requestId, session) => cdp.send("Network.getResponseBody", { requestId }, session ?? sessionId),
    getPostData: async (requestId, session) => (await cdp.send("Network.getRequestPostData", { requestId }, session ?? sessionId)).postData,
  });
  const watch = new SignInWatch(timing.graceMs);
  const enable = async (session: string) => { await cdp.send("Network.enable", { maxPostDataSize: 65_536 }, session); };
  cdp.on((event: CdpEvent) => {
    if (event.method === "Target.attachedToTarget" && typeof event.params?.sessionId === "string") {
      // An iframe or a worker can make the token request. Watch it too.
      sessions.add(event.params.sessionId);
      void enable(event.params.sessionId).catch(() => {});
      return;
    }
    if (!event.sessionId || !sessions.has(event.sessionId)) return;
    capture.onEvent(event);
    if (event.sessionId === sessionId && event.method === "Page.frameNavigated" && !event.params?.frame?.parentId) {
      watch.observe(String(event.params.frame.url ?? ""), Date.now());
    }
  });
  await enable(sessionId);
  await cdp.send("Page.enable", {}, sessionId);
  await cdp.send("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: false, flatten: true }, sessionId).catch(() => {});
  if (mode === "renew") {
    // Headless Chrome names itself HeadlessChrome. Send the ordinary name that a visible browser sends.
    const { userAgent } = await cdp.send<{ userAgent?: string }>("Browser.getVersion");
    if (userAgent?.includes("HeadlessChrome")) {
      await cdp.send("Network.setUserAgentOverride", { userAgent: userAgent.replace("HeadlessChrome", "Chrome") }, sessionId);
    }
  }
  await cdp.send("Page.navigate", { url: hooks.url ?? TEAMS_URL }, sessionId);

  const scan = async () => {
    try {
      const result = await cdp.send<{ result?: { value?: unknown } }>("Runtime.evaluate", { expression: STORAGE_SCAN, returnByValue: true }, sessionId);
      const values = result.result?.value;
      if (Array.isArray(values)) for (const value of values) if (typeof value === "string") capture.addAccessToken(value);
    } catch { /* The page may be mid-navigation. The next check tries again. */ }
  };

  const limit = mode === "login" ? timing.loginTimeoutMs : timing.renewTimeoutMs;
  const started = Date.now();
  const deadline = started + limit;
  let scanned = false;
  const run = async () => {
    for (;;) {
      await Bun.sleep(timing.pollMs);
      await capture.settled();
      if (mode === "renew" && watch.needsSignIn(Date.now())) throw new SignInRequiredError();
      if (!capture.missing(CORE).length) break;
      if (!scanned && Date.now() - started >= timing.scanAfterMs) { scanned = true; await scan(); }
      if (Date.now() >= deadline) break;
    }
    if (!capture.missing(CORE).length) {
      await Bun.sleep(timing.settleMs);
      await capture.settled();
      if (capture.missing(["substrate"]).length) {
        await cdp.send("Runtime.evaluate", { expression: SEARCH_NUDGE }, sessionId).catch(() => {});
        const until = Math.min(Date.now() + timing.nudgeMs, deadline);
        while (Date.now() < until && capture.missing(["substrate"]).length) {
          await Bun.sleep(timing.pollMs);
          await capture.settled();
        }
      }
    }
    await capture.settled();
    const captured = capture.captured;
    if (!KINDS.some(kind => captured.tokens[kind])) {
      if (mode === "renew" && watch.onSignInPage) throw new SignInRequiredError();
      throw new Error(`Teams issued no tokens within ${Math.round(limit / 1000)} seconds.${mode === "login" ? " Sign in, then run the command again." : ""}`);
    }
    return captured;
  };
  const closedEarly = cdp.closed.then((): never => { throw new Error("The browser closed before Teams issued tokens."); });
  return Promise.race([run(), closedEarly]);
}

/** Runs the browser once and returns what the page requested. It takes no lock. A caller holds one. */
export async function runBrowser(mode: BrowserMode, options: BrowserOptions = {}): Promise<RunResult> {
  const { name } = await loadProfile();
  const dir = browserProfileDir(name);
  const executable = hooks.executable ?? findBrowser(options.browser);
  await prepareProfileDir(dir);
  if (await profileInUse(dir)) throw new Error("The browser profile is open in a running browser. Close that browser and try again.");
  const timing = { ...DEFAULT_TIMING, ...hooks.timing };
  const launched = await (hooks.launcher ?? launchBrowser)({ executable, userDataDir: dir, headless: mode === "renew" });
  let cdp: CdpClient | undefined;
  try {
    cdp = await CdpClient.connect(launched.wsUrl);
    return { captured: await drive(cdp, mode, timing), method: mode };
  } finally {
    // A clean close lets the browser write the session cookies that the next renewal needs.
    if (cdp) { await cdp.send("Browser.close", {}, undefined, 3_000).catch(() => {}); cdp.close(); }
    await launched.close().catch(() => {});
  }
}

export type Imported = { profileName: string; profile: Profile; notCaptured: TokenKind[]; refreshCaptured: boolean };

/** Saves what the browser captured into the profile. Another account is refused. It takes no lock. */
export async function importCaptured(captured: Captured, method: BrowserMode): Promise<Imported> {
  const { name, profile } = await loadProfile();
  const claims = captured.claims;
  if (!claims?.oid || !claims.tid || !KINDS.some(kind => captured.tokens[kind])) throw new Error("The browser produced no Teams tokens.");
  if (profile && (profile.tenantId !== claims.tid || profile.userId !== claims.oid)) {
    throw new Error("The browser signed in as another account than this profile. Run teams-cli logout first, or use --profile NAME.");
  }
  const refresh = captured.refresh && captured.refresh.tenantId === claims.tid ? captured.refresh : undefined;
  const next: Profile = {
    refresh: refresh ?? profile?.refresh,
    browser: { lastAt: new Date().toISOString(), lastMethod: method },
    sharepointHost: profile?.sharepointHost,
    tokens: { ...(profile?.tokens ?? {}), ...captured.tokens },
    userId: claims.oid,
    tenantId: claims.tid,
    upn: claims.upn ?? profile?.upn ?? "",
    name: claims.name ?? profile?.name ?? "",
    region: captured.region ?? profile?.region ?? "amer",
    updatedAt: new Date().toISOString(),
  };
  await saveProfile(name, next);
  return {
    profileName: name,
    profile: next,
    notCaptured: BROWSER_KINDS.filter(kind => !captured.tokens[kind]),
    refreshCaptured: !!refresh,
  };
}

/** Renews through the headless profile. It prints nothing, so a command with `--json` stays clean. The caller holds the lock. */
export async function renewFromBrowser(options: BrowserOptions = {}): Promise<Imported> {
  const { name } = await loadProfile();
  if (!(await browserProfileExists(name))) throw new Error("No browser profile. Run: teams-cli auth login");
  const { captured } = await runBrowser("renew", options);
  return importCaptured(captured, "renew");
}

/** One automatic attempt for `refreshToken()`. It never throws. */
export async function tryBrowserRenew(): Promise<{ ok: true } | { ok: false; reason?: string; signInRequired?: boolean }> {
  const { name } = await loadProfile();
  if (!(await browserProfileExists(name))) return { ok: false };
  try {
    await renewFromBrowser();
    return { ok: true };
  } catch (error) {
    if (error instanceof SignInRequiredError) return { ok: false, signInRequired: true };
    return { ok: false, reason: redact(error instanceof Error ? error.message : String(error)) };
  }
}

function report(imported: Imported, captured: Captured, now = Date.now()): void {
  for (const kind of BROWSER_KINDS) {
    const got = captured.tokens[kind];
    if (got) console.log(describeKind(kind, got.exp, now).replace(": ok, expires", ": captured, expires"));
    else if (imported.profile.tokens[kind]) console.log(`${describeKind(kind, imported.profile.tokens[kind]!.exp, now)} (not captured this time)`);
    else console.log(`${kind}: not captured`);
  }
  console.log(imported.refreshCaptured ? "refresh: captured, good for about 24 hours" : "refresh: not captured");
  console.log(`Profile: ${imported.profileName}`);
}

function finish(imported: Imported, captured: Captured): void {
  report(imported, captured);
  const missing = CORE.filter(kind => !captured.tokens[kind]);
  if (missing.length) {
    console.error(`Teams did not issue these tokens: ${missing.join(", ")}.`);
    process.exitCode = 1;
  }
}

/** `auth login`: a visible browser, a person signs in, then the tokens are saved and the browser closes. */
export async function authLogin(options: BrowserOptions = {}): Promise<void> {
  const { name } = await loadProfile();
  console.error(`Opening a browser. Sign in to Teams there. Browser profile: ${browserProfileDir(name)}`);
  const { captured } = await runBrowser("login", options);
  const imported = await locked(() => importCaptured(captured, "login"));
  finish(imported, captured);
}

/** `auth renew`: the same profile, headless. Needs a prior `auth login`. */
export async function authRenew(options: BrowserOptions = {}): Promise<void> {
  const result = await locked(async () => {
    const { name } = await loadProfile();
    if (!(await browserProfileExists(name))) throw new Error("No browser profile. Run: teams-cli auth login");
    const { captured } = await runBrowser("renew", options);
    return { imported: await importCaptured(captured, "renew"), captured };
  });
  finish(result.imported, result.captured);
}

/** The status lines for `auth status`. */
export async function browserStatusLine(name: string, profile: Profile | undefined): Promise<string> {
  if (!(await browserProfileExists(name))) return "browser profile: none. Run: teams-cli auth login";
  const meta = profile?.browser;
  return meta ? `browser profile: present, last ${meta.lastMethod === "login" ? "sign-in" : "renewal"} ${meta.lastAt}` : "browser profile: present, no token import yet";
}
