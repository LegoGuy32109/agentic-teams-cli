import { afterEach, beforeEach, expect, test } from "bun:test";
import { chmod, mkdir, mkdtemp, readFile, stat, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { browserProfileDir, browserProfileExists, findBrowser, installedBrowserPaths, launchBrowser, parseDevToolsActivePort, prepareProfileDir, profileInUse } from "../src/browser.ts";
import { configDir } from "../src/config.ts";
import { resetLocalState } from "./fake-teams.ts";

beforeEach(resetLocalState);
let dir: string;
beforeEach(async () => { dir = await mkdtemp(join(tmpdir(), "teams-cli-browser-")); });
afterEach(() => { process.exitCode = 0; });

const which = (found: Record<string, string>) => (name: string) => found[name] ?? null;

test("the browser comes from --browser, then the environment, then chromium, google-chrome, brave", () => {
  const all = which({ chromium: "/bin/chromium", "google-chrome": "/bin/chrome", brave: "/bin/brave", env: "/bin/from-env-name" });
  expect(findBrowser("/opt/flag/browser", "/opt/env/browser", all)).toBe("/opt/flag/browser");
  expect(findBrowser(undefined, "/opt/env/browser", all)).toBe("/opt/env/browser");
  expect(findBrowser(undefined, "env", all)).toBe("/bin/from-env-name");
  expect(findBrowser(undefined, undefined, all)).toBe("/bin/chromium");
  expect(findBrowser(undefined, undefined, which({ "google-chrome": "/bin/chrome", brave: "/bin/brave" }))).toBe("/bin/chrome");
  expect(findBrowser(undefined, undefined, which({ brave: "/bin/brave" }))).toBe("/bin/brave");
});

test("a browser that was named on purpose but cannot be found is an error, not a fallback", () => {
  const all = which({ chromium: "/bin/chromium" });
  expect(() => findBrowser("nope", undefined, all)).toThrow("--browser was not found");
  expect(() => findBrowser(undefined, "nope", all)).toThrow("TEAMS_CLI_BROWSER was not found");
  expect(() => findBrowser(undefined, undefined, which({}))).toThrow("No browser found");
});

test("on macOS and Windows a browser that is not on PATH is found at its install path", () => {
  const mac = installedBrowserPaths("darwin", {});
  expect(mac[0]).toBe("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome");
  expect(findBrowser(undefined, undefined, which({}), mac, path => path.includes("Brave"))).toBe("/Applications/Brave Browser.app/Contents/MacOS/Brave Browser");
  const windows = installedBrowserPaths("win32", { PROGRAMFILES: "C:\\Program Files", "PROGRAMFILES(X86)": "C:\\Program Files (x86)" });
  const edge = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
  expect(windows).toContain(edge);
  expect(findBrowser(undefined, undefined, which({}), windows, path => path === edge)).toBe(edge);
  expect(installedBrowserPaths("linux", {})).toEqual([]);
});

test("a Windows path given with --browser is used as a path", () => {
  expect(findBrowser("C:\\Tools\\chrome.exe", undefined, which({}))).toBe("C:\\Tools\\chrome.exe");
});

test("the profile folder name cannot leave the browser-profile folder", () => {
  expect(browserProfileDir("default")).toBe(join(configDir(), "browser-profile", "default"));
  for (const bad of ["../x", "a/b", "", ".."]) expect(() => browserProfileDir(bad)).toThrow("A profile name may use");
});

test("the profile folder and its parent are created with mode 0700", async () => {
  const target = join(dir, "browser-profile", "default");
  await mkdir(target, { recursive: true, mode: 0o755 });
  await chmod(join(dir, "browser-profile"), 0o755);
  await chmod(target, 0o755);
  await prepareProfileDir(target);
  expect((await stat(target)).mode & 0o777).toBe(0o700);
  expect((await stat(join(dir, "browser-profile"))).mode & 0o777).toBe(0o700);
});

test("a browser profile exists once Chromium has written Local State", async () => {
  expect(await browserProfileExists("default")).toBe(false);
  await prepareProfileDir(browserProfileDir("default"));
  expect(await browserProfileExists("default")).toBe(false);
  await writeFile(join(browserProfileDir("default"), "Local State"), "{}");
  expect(await browserProfileExists("default")).toBe(true);
});

test("a profile is in use while its SingletonLock names a live process", async () => {
  expect(await profileInUse(dir)).toBe(false);
  await symlink(`host-${process.pid}`, join(dir, "SingletonLock"));
  expect(await profileInUse(dir)).toBe(true);
});

test("a stale SingletonLock does not block the profile", async () => {
  await symlink("host-2147483000", join(dir, "SingletonLock"));
  expect(await profileInUse(dir)).toBe(false);
});

test("the debugging port file gives the browser socket address", () => {
  expect(parseDevToolsActivePort("40123\n/devtools/browser/abc-def\n")).toBe("ws://127.0.0.1:40123/devtools/browser/abc-def");
  for (const bad of ["", "abc\n/devtools/browser/x", "4000\n/other", "4000"]) expect(parseDevToolsActivePort(bad)).toBeUndefined();
});

async function fakeBrowser(script: string): Promise<string> {
  const path = join(dir, "fake-browser.sh");
  await writeFile(path, `#!/bin/sh\n${script}\n`);
  await chmod(path, 0o755);
  return path;
}

test("the launcher passes the dedicated profile, port 0 and headless flags, then returns the socket address", async () => {
  const userDataDir = join(dir, "profile");
  await mkdir(userDataDir);
  await writeFile(join(userDataDir, "DevToolsActivePort"), "1\n/devtools/browser/stale\n");
  const args = join(dir, "args.txt");
  const executable = await fakeBrowser(`printf '%s\\n' "$@" > '${args}'\nsleep 0.3\nprintf '41234\\n/devtools/browser/fresh\\n' > "$(echo "$1" | cut -d= -f2)/DevToolsActivePort"\nexec sleep 30`);
  const browser = await launchBrowser({ executable, userDataDir, headless: true, startUrl: "about:blank", exitWaitMs: 100, platform: "linux" });
  expect(browser.wsUrl).toBe("ws://127.0.0.1:41234/devtools/browser/fresh");
  const given = (await readFile(args, "utf8")).trim().split("\n");
  expect(given).toContain(`--user-data-dir=${userDataDir}`);
  expect(given).toContain("--remote-debugging-port=0");
  expect(given).toContain("--headless=new");
  expect(given).toContain("--password-store=basic");
  expect(given.at(-1)).toBe("about:blank");
  expect(given.some(arg => arg.startsWith("--user-agent"))).toBe(false);
  await browser.close();
});

test("a visible launch has no headless flag", async () => {
  const userDataDir = join(dir, "profile");
  await mkdir(userDataDir);
  const args = join(dir, "args.txt");
  const executable = await fakeBrowser(`printf '%s\\n' "$@" > '${args}'\nprintf '41234\\n/devtools/browser/x\\n' > "$(echo "$1" | cut -d= -f2)/DevToolsActivePort"\nexec sleep 30`);
  const browser = await launchBrowser({ executable, userDataDir, headless: false, exitWaitMs: 100 });
  expect((await readFile(args, "utf8"))).not.toContain("headless");
  await browser.close();
});

test("close ends a browser that ignores the debugging close", async () => {
  const userDataDir = join(dir, "profile");
  await mkdir(userDataDir);
  const executable = await fakeBrowser(`printf '41234\\n/devtools/browser/x\\n' > "$(echo "$1" | cut -d= -f2)/DevToolsActivePort"\nexec sleep 30`);
  const browser = await launchBrowser({ executable, userDataDir, headless: true, exitWaitMs: 100 });
  const started = Date.now();
  await browser.close();
  expect(Date.now() - started).toBeLessThan(2_000);
});

test("a browser that exits before it opens a port is reported", async () => {
  const userDataDir = join(dir, "profile");
  await mkdir(userDataDir);
  const executable = await fakeBrowser("exit 3");
  await expect(launchBrowser({ executable, userDataDir, headless: true })).rejects.toThrow("exited before it opened its debugging port (exit code 3)");
});

test("a browser that never opens a port times out and is stopped", async () => {
  const userDataDir = join(dir, "profile");
  await mkdir(userDataDir);
  const executable = await fakeBrowser("exec sleep 30");
  await expect(launchBrowser({ executable, userDataDir, headless: true, startTimeoutMs: 300 })).rejects.toThrow("did not open its debugging port in time");
});

test("a missing executable is reported without a stack", async () => {
  await expect(launchBrowser({ executable: join(dir, "missing"), userDataDir: dir, headless: true })).rejects.toThrow("Could not start the browser");
});
