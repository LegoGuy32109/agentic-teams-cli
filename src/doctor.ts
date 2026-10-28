import { existsSync } from "node:fs";
import { browserProfileExists, findBrowser } from "./browser.ts";
import { loadProfile, type TokenKind } from "./config.ts";
import { parseSharepointHost } from "./sharepoint.ts";
import { describeKind, externalTokenSource, KINDS, resolveToken, UNLOCKS } from "./tokens.ts";
import { teamsHost } from "./transport.ts";
import { VERSION } from "./version.ts";

/**
 * `auth doctor`: one check per line, each with a status and a fix. It reads local state and makes one
 * unauthenticated request for the clock. It writes nothing and never starts a browser.
 */

export type CheckStatus = "ok" | "warn" | "fail";
export type Check = { name: string; status: CheckStatus; detail: string; fix: string };

const CORE: TokenKind[] = ["chat", "csa", "spaces"];
/** Microsoft refuses a token when the clock is off by more than 5 minutes. */
const SKEW_FAIL_MS = 5 * 60_000;
const SKEW_WARN_MS = 60_000;
const CLOCK_TIMEOUT_MS = 5_000;

export type DoctorDeps = {
  platform: NodeJS.Platform;
  arch: string;
  find: (flag: string | undefined) => string;
  exists: (path: string) => boolean;
  now: () => number;
};

const defaultDeps: DoctorDeps = {
  platform: process.platform,
  arch: process.arch,
  find: flag => findBrowser(flag),
  exists: existsSync,
  now: () => Date.now(),
};

const check = (name: string, status: CheckStatus, detail: string, fix = ""): Check => ({ name, status, detail, fix });

/** Reads the `Date` header of one unauthenticated request. Returns undefined when offline. */
async function clockCheck(deps: DoctorDeps): Promise<Check> {
  const name = "clock skew";
  let response: Response;
  const sent = deps.now();
  try {
    response = await fetch(`${teamsHost()}/`, { method: "HEAD", redirect: "manual", signal: AbortSignal.timeout(CLOCK_TIMEOUT_MS) });
  } catch {
    return check(name, "warn", "skipped, the Teams host cannot be reached (offline?)", "Connect to the network, then run auth doctor again.");
  }
  const header = response.headers.get("date");
  const server = header ? Date.parse(header) : NaN;
  if (Number.isNaN(server)) return check(name, "warn", "skipped, the server sent no Date header");
  // The reply takes about half the round trip to arrive, so measure against the midpoint.
  const received = deps.now();
  const skew = Math.round((sent + received) / 2 - server);
  const seconds = Math.abs(skew) / 1000;
  const detail = `${seconds < 1 ? "under 1" : Math.round(seconds)} second(s) ${skew > 0 ? "ahead of" : "behind"} Microsoft`;
  if (Math.abs(skew) > SKEW_FAIL_MS) return check(name, "fail", detail, "Set the system clock from the network. Microsoft refuses tokens when the clock is off by more than 5 minutes.");
  if (Math.abs(skew) > SKEW_WARN_MS) return check(name, "warn", detail, "Set the system clock from the network.");
  return check(name, "ok", detail);
}

/** Runs every check. It does not throw. */
export async function runDoctor(options: { browser?: string } = {}, overrides: Partial<DoctorDeps> = {}): Promise<Check[]> {
  const deps = { ...defaultDeps, ...overrides };
  const checks: Check[] = [];
  const external = externalTokenSource();
  const { name: profileName, profile } = await loadProfile();

  checks.push(check("version", "ok", `teams-cli ${VERSION}`));
  checks.push(check("platform", "ok", `${deps.platform} ${deps.arch}`));

  let browserPath: string | undefined;
  let browserError = "";
  try {
    browserPath = deps.find(options.browser);
    if (!deps.exists(browserPath)) { browserError = `The browser path does not exist: ${browserPath}`; browserPath = undefined; }
  } catch (error) {
    browserError = error instanceof Error ? error.message : String(error);
  }
  checks.push(browserPath
    ? check("browser", "ok", browserPath)
    : check("browser", external ? "warn" : "fail", browserError, "Install Chrome, Edge, Brave, or Chromium, or pass --browser PATH, or set TEAMS_CLI_BROWSER."));

  const browserProfile = await browserProfileExists(profileName);
  checks.push(browserProfile
    ? check("browser profile", "ok", `present for profile ${profileName}`)
    : check("browser profile", "warn", `none for profile ${profileName}`, "Run: teams-cli auth login"));

  checks.push(profile
    ? check("stored profile", "ok", `${profileName}${profile.upn ? `, ${profile.upn}` : ""}`)
    : check("stored profile", external ? "warn" : "fail", `none named ${profileName}`, "Run: teams-cli auth login"));

  const renewable = !!profile?.refresh || browserProfile;
  const now = deps.now();
  for (const kind of KINDS) {
    let resolved;
    try { resolved = await resolveToken(kind); } catch (error) {
      checks.push(check(`${kind} token`, "fail", error instanceof Error ? error.message.replace(/eyJ[\w.-]*/g, "<token>") : "unreadable", "Supply a valid token, or run: teams-cli auth login"));
      continue;
    }
    const line = describeKind(kind, resolved?.exp, now);
    const detail = `${line.slice(kind.length + 2).replace(/^ok, /, "")}${resolved ? ` (source: ${resolved.source})` : ""}. Unlocks: ${UNLOCKS[kind]}`;
    const usable = !!resolved && resolved.exp * 1000 - now >= 60_000;
    if (usable) { checks.push(check(`${kind} token`, "ok", detail)); continue; }
    if (CORE.includes(kind)) {
      checks.push(renewable
        ? check(`${kind} token`, "warn", `${detail}. It renews when a command needs it.`, "Run: teams-cli auth renew")
        : check(`${kind} token`, "fail", detail, "Run: teams-cli auth login"));
    } else {
      checks.push(renewable
        ? check(`${kind} token`, "ok", `${detail}. It renews or is captured when a command needs it.`)
        : check(`${kind} token`, "warn", detail, "Run: teams-cli auth login"));
    }
  }

  const host = parseSharepointHost(process.env.TEAMS_SHAREPOINT_HOST) ?? parseSharepointHost(profile?.sharepointHost);
  checks.push(host
    ? check("sharepoint host", "ok", host)
    : check("sharepoint host", "warn", "not known yet. The first files command finds it.", "Set TEAMS_SHAREPOINT_HOST, for example contoso-my.sharepoint.com."));

  const region = process.env.TEAMS_REGION || profile?.region;
  if (!region) checks.push(check("region", "warn", "not known, no stored profile", "Run: teams-cli auth login"));
  else if (region !== "amer") checks.push(check("region", "warn", `${region}. Image upload is verified only for amer.`, "Use --file instead of --image, or capture the hosts for this region."));
  else checks.push(check("region", "ok", region));

  checks.push(await clockCheck(deps));
  return checks;
}

/** One line per check: status, name, detail, and the fix when there is one. */
export function formatCheck(item: Check): string {
  return `${item.status.padEnd(4)}  ${item.name}: ${item.detail}${item.fix ? ` Fix: ${item.fix}` : ""}`;
}

export async function authDoctor(options: { json?: boolean; browser?: string } = {}, overrides: Partial<DoctorDeps> = {}): Promise<void> {
  const checks = await runDoctor(options, overrides);
  if (options.json) console.log(JSON.stringify({ checks }, null, 2));
  else for (const item of checks) console.log(formatCheck(item));
  if (checks.some(item => item.status === "fail")) process.exitCode = 1;
}
