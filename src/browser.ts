import { existsSync } from "node:fs";
import { mkdir, open, readFile, readlink, rm, stat } from "node:fs/promises";
import { join, posix, win32 } from "node:path";
import { configDir, restrictMode } from "./config.ts";

/** Finding, preparing and launching the dedicated Chromium-family browser. */

/** Names looked up on PATH, in order. */
const CANDIDATES = ["chromium", "google-chrome", "brave", "microsoft-edge"];

/** Install paths for systems where browsers are not on PATH. Edge ships with Windows, so Windows always has one. */
export function installedBrowserPaths(platform: NodeJS.Platform = process.platform, env: Record<string, string | undefined> = process.env): string[] {
  if (platform === "darwin") {
    return ["Google Chrome", "Chromium", "Brave Browser", "Microsoft Edge"].map(app => `/Applications/${app}.app/Contents/MacOS/${app}`);
  }
  if (platform === "win32") {
    const roots = [env.PROGRAMFILES, env["PROGRAMFILES(X86)"], env.LOCALAPPDATA].filter((root): root is string => !!root);
    const apps = ["Google\\Chrome\\Application\\chrome.exe", "BraveSoftware\\Brave-Browser\\Application\\brave.exe", "Microsoft\\Edge\\Application\\msedge.exe"];
    return apps.flatMap(app => roots.map(root => `${root}\\${app}`));
  }
  return [];
}

/**
 * Picks the browser: `--browser`, then TEAMS_CLI_BROWSER, then chromium, google-chrome, brave, microsoft-edge on PATH,
 * then the usual install paths on macOS and Windows. A path that was named on purpose must exist. The first two sources
 * never fall through to a guess.
 */
export function findBrowser(
  flag: string | undefined,
  env: string | undefined = process.env.TEAMS_CLI_BROWSER,
  which: (name: string) => string | null = name => Bun.which(name),
  installed: string[] = installedBrowserPaths(),
  exists: (path: string) => boolean = existsSync,
): string {
  for (const [source, value] of [["--browser", flag], ["TEAMS_CLI_BROWSER", env]] as const) {
    if (!value) continue;
    const found = /[\\/]/.test(value) ? value : which(value);
    if (!found) throw new Error(`The browser named by ${source} was not found: ${value}`);
    return found;
  }
  for (const name of CANDIDATES) {
    const found = which(name);
    if (found) return found;
  }
  const path = installed.find(exists);
  if (path) return path;
  throw new Error("No browser found. Install Chrome, Edge, Brave, or Chromium, or pass --browser PATH, or set TEAMS_CLI_BROWSER.");
}

/** A profile name becomes a directory name, so it must stay inside the browser-profile folder. */
export function browserProfileDir(profileName: string): string {
  if (!/^[\w][\w.-]*$/.test(profileName)) throw new Error("A profile name may use letters, digits, dots, dashes, and underscores.");
  return join(configDir(), "browser-profile", profileName);
}

/** A profile exists once the browser has started in it. Chromium writes `Local State` on its first run. */
export async function browserProfileExists(profileName: string): Promise<boolean> {
  try {
    return (await stat(join(browserProfileDir(profileName), "Local State"))).isFile();
  } catch {
    return false;
  }
}

/**
 * The profile folder holds session cookies. Create it, and its parent, readable by this user only.
 * On Windows there are no modes to set. The folder sits under the user's `%APPDATA%`, which is per-user by default.
 */
export async function prepareProfileDir(dir: string, platform: NodeJS.Platform = process.platform): Promise<void> {
  const parent = join(dir, "..");
  await mkdir(dir, { recursive: true, mode: 0o700 });
  await restrictMode(parent, 0o700, platform);
  await restrictMode(dir, 0o700, platform);
}

/** Opens a file for writing and closes it at once. It fails while another process holds the file. */
async function openForWrite(path: string): Promise<void> {
  await (await open(path, "r+")).close();
}

/**
 * Is a browser running in this profile?
 *
 * On Linux and macOS, Chromium links `SingletonLock` to `host-pid` while a browser runs, so the check reads the link and
 * tests the process. Windows has no such link. There, Chromium keeps `<profile>/lockfile` open with no sharing while it
 * runs. The profile counts as in use when that file exists and cannot be opened for writing. A file that opens is a
 * leftover, so the profile is free. The check never removes the file. It is not verified on a real Windows machine.
 * If the probe fails for any reason other than a missing file, the profile counts as in use.
 */
export async function profileInUse(
  dir: string,
  platform: NodeJS.Platform = process.platform,
  probe: (path: string) => Promise<void> = openForWrite,
): Promise<boolean> {
  if (platform === "win32") {
    try {
      await probe(win32.join(dir, "lockfile"));
      return false;
    } catch (error) {
      return (error as NodeJS.ErrnoException).code !== "ENOENT";
    }
  }
  let target: string;
  try { target = await readlink(join(dir, "SingletonLock")); } catch { return false; }
  const pid = Number(target.match(/-(\d+)$/)?.[1]);
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try { process.kill(pid, 0); return true; } catch (error) { return (error as NodeJS.ErrnoException).code === "EPERM"; }
}

export type LaunchOptions = { executable: string; userDataDir: string; headless: boolean; startUrl?: string; startTimeoutMs?: number; exitWaitMs?: number; platform?: NodeJS.Platform };
export type LaunchedBrowser = { wsUrl: string; close(): Promise<void> };
export type Launcher = (options: LaunchOptions) => Promise<LaunchedBrowser>;

/** Reads the port and the browser socket path that Chromium writes when remote debugging starts. */
export function parseDevToolsActivePort(text: string): string | undefined {
  // Windows writes CRLF line endings.
  const [port, path] = text.split(/\r?\n/);
  if (!port || !/^\d+$/.test(port.trim()) || !path?.trim().startsWith("/devtools/browser/")) return undefined;
  return `ws://127.0.0.1:${port.trim()}${path.trim()}`;
}

/** The file where Chromium writes its debugging port. A Windows folder gets a Windows path. */
export function devToolsPortFile(userDataDir: string, platform: NodeJS.Platform = process.platform): string {
  return (platform === "win32" ? win32 : posix).join(userDataDir, "DevToolsActivePort");
}

/**
 * The cookie-store flag for a platform. It gives the visible and the headless run the same cookie store, with no
 * keyring or Keychain prompt in a timer. Each flag goes only to its own platform. Windows needs none.
 */
export function cookieStoreArgs(platform: NodeJS.Platform): string[] {
  if (platform === "linux") return ["--password-store=basic"];
  if (platform === "darwin") return ["--use-mock-keychain"];
  return [];
}

/** The browser command line. The visible and the headless run share every flag except the headless ones. */
export function browserArgs(options: LaunchOptions, platform: NodeJS.Platform = options.platform ?? process.platform): string[] {
  return [
    `--user-data-dir=${options.userDataDir}`,
    "--remote-debugging-port=0",
    "--no-first-run",
    "--no-default-browser-check",
    ...cookieStoreArgs(platform),
    ...(options.headless ? ["--headless=new", "--disable-gpu"] : []),
    options.startUrl ?? "about:blank",
  ];
}

/** Starts the browser with its own profile and an OS-chosen debugging port. Never attaches to a running browser. */
export const launchBrowser: Launcher = async options => {
  const portFile = devToolsPortFile(options.userDataDir, options.platform ?? process.platform);
  await rm(portFile, { force: true });
  const args = browserArgs(options);
  let proc: ReturnType<typeof Bun.spawn>;
  try {
    proc = Bun.spawn([options.executable, ...args], { stdin: "ignore", stdout: "ignore", stderr: "ignore" });
  } catch {
    throw new Error(`Could not start the browser: ${options.executable}`);
  }
  let exited = false;
  void proc.exited.then(() => { exited = true; });
  const stop = async () => {
    if (exited) return;
    const wait = (ms: number) => Promise.race([proc.exited, Bun.sleep(ms)]);
    await wait(options.exitWaitMs ?? 8_000);
    if (!exited) { proc.kill("SIGTERM"); await wait(options.exitWaitMs ?? 3_000); }
    if (!exited) { proc.kill("SIGKILL"); await proc.exited; }
  };
  const deadline = Date.now() + (options.startTimeoutMs ?? 20_000);
  for (;;) {
    const wsUrl = parseDevToolsActivePort(await readFile(portFile, "utf8").catch(() => ""));
    if (wsUrl) return { wsUrl, close: stop };
    if (exited) throw new Error(`The browser exited before it opened its debugging port (exit code ${proc.exitCode}).`);
    if (Date.now() >= deadline) {
      proc.kill("SIGKILL");
      throw new Error("The browser did not open its debugging port in time.");
    }
    await Bun.sleep(100);
  }
};
