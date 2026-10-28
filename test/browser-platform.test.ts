import { beforeEach, expect, test } from "bun:test";
import { chmod, mkdir, mkdtemp, readFile, stat, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { browserArgs, cookieStoreArgs, devToolsPortFile, launchBrowser, parseDevToolsActivePort, prepareProfileDir, profileInUse } from "../src/browser.ts";
import { restrictMode } from "../src/config.ts";

/** Every platform branch of the dedicated browser, with the platform passed in. */

let dir: string;
beforeEach(async () => { dir = await mkdtemp(join(tmpdir(), "teams-cli-platform-")); });

const base = { executable: "/x/browser", userDataDir: "/p", startUrl: "about:blank" };

test("each cookie-store flag goes only to its own platform, for the visible and the headless run", () => {
  expect(cookieStoreArgs("linux")).toEqual(["--password-store=basic"]);
  expect(cookieStoreArgs("darwin")).toEqual(["--use-mock-keychain"]);
  expect(cookieStoreArgs("win32")).toEqual([]);
  for (const headless of [false, true]) {
    const linux = browserArgs({ ...base, headless }, "linux");
    const mac = browserArgs({ ...base, headless }, "darwin");
    const windows = browserArgs({ ...base, headless }, "win32");
    expect(linux).toContain("--password-store=basic");
    expect(linux).not.toContain("--use-mock-keychain");
    expect(mac).toContain("--use-mock-keychain");
    expect(mac).not.toContain("--password-store=basic");
    expect(windows).not.toContain("--use-mock-keychain");
    expect(windows).not.toContain("--password-store=basic");
    for (const args of [linux, mac, windows]) {
      expect(args).toContain("--remote-debugging-port=0");
      expect(args.includes("--headless=new")).toBe(headless);
      expect(args.at(-1)).toBe("about:blank");
    }
  }
});

test("the launcher reads the platform from its options and passes the macOS flag", async () => {
  const userDataDir = join(dir, "profile");
  await mkdir(userDataDir);
  const args = join(dir, "args.txt");
  const script = join(dir, "fake-browser.sh");
  await writeFile(script, `#!/bin/sh\nprintf '%s\\n' "$@" > '${args}'\nprintf '41234\\n/devtools/browser/x\\n' > "$(echo "$1" | cut -d= -f2)/DevToolsActivePort"\nexec sleep 30\n`);
  await chmod(script, 0o755);
  const browser = await launchBrowser({ executable: script, userDataDir, headless: true, exitWaitMs: 100, platform: "darwin" });
  const given = (await readFile(args, "utf8")).trim().split("\n");
  expect(given).toContain("--use-mock-keychain");
  expect(given).not.toContain("--password-store=basic");
  await browser.close();
});

test("on Windows the mode steps do nothing, and on Linux they set the mode", async () => {
  const target = join(dir, "browser-profile", "default");
  await mkdir(target, { recursive: true, mode: 0o755 });
  await chmod(join(dir, "browser-profile"), 0o755);
  await chmod(target, 0o755);
  await prepareProfileDir(target, "win32");
  expect((await stat(target)).mode & 0o777).toBe(0o755);
  // chmod on a missing path throws. On win32 it is skipped without an error.
  await restrictMode(join(dir, "missing"), 0o600, "win32");
  await expect(restrictMode(join(dir, "missing"), 0o600, "linux")).rejects.toThrow();
  await prepareProfileDir(target, "linux");
  expect((await stat(target)).mode & 0o777).toBe(0o700);
});

test("on Windows a profile is in use when its lockfile exists and cannot be opened", async () => {
  expect(await profileInUse(dir, "win32")).toBe(false);
  await writeFile(join(dir, "lockfile"), "");
  expect(await profileInUse(dir, "win32")).toBe(false);
  const probed: string[] = [];
  const refuse = (code: string) => async (path: string) => { probed.push(path); throw Object.assign(new Error(code), { code }); };
  expect(await profileInUse("C:\\Users\\me\\profile", "win32", refuse("EBUSY"))).toBe(true);
  expect(probed[0]).toBe("C:\\Users\\me\\profile\\lockfile");
  expect(await profileInUse(dir, "win32", refuse("EPERM"))).toBe(true);
  expect(await profileInUse(dir, "win32", refuse("EACCES"))).toBe(true);
  expect(await profileInUse(dir, "win32", refuse("ENOENT"))).toBe(false);
});

test("on Windows the SingletonLock link is not read, and on macOS it is", async () => {
  await symlink(`host-${process.pid}`, join(dir, "SingletonLock"));
  expect(await profileInUse(dir, "win32")).toBe(false);
  expect(await profileInUse(dir, "darwin")).toBe(true);
  expect(await profileInUse(dir, "linux")).toBe(true);
});

test("the debugging port file works with Windows paths and CRLF line endings", () => {
  const folder = "C:\\Users\\me\\AppData\\Roaming\\teams-cli\\browser-profile\\default";
  expect(devToolsPortFile(folder, "win32")).toBe(`${folder}\\DevToolsActivePort`);
  expect(devToolsPortFile("/home/me/profile", "linux")).toBe("/home/me/profile/DevToolsActivePort");
  expect(parseDevToolsActivePort("40123\r\n/devtools/browser/abc-def\r\n")).toBe("ws://127.0.0.1:40123/devtools/browser/abc-def");
});
