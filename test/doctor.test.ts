import { afterEach, beforeEach, expect, test } from "bun:test";
import { existsSync } from "node:fs";
import { writeFile } from "node:fs/promises";
import { join } from "node:path";
import { browserProfileDir, prepareProfileDir } from "../src/browser.ts";
import { configDir, loadProfile, saveProfile } from "../src/config.ts";
import { authDoctor, runDoctor, type Check } from "../src/doctor.ts";
import { VERSION } from "../src/version.ts";
import { installFakeTeams, resetLocalState, seedProfile } from "./fake-teams.ts";
import { runCli } from "./cli.ts";

beforeEach(resetLocalState);
let fake: ReturnType<typeof installFakeTeams> | undefined;
afterEach(() => {
  fake?.uninstall();
  fake = undefined;
  process.exitCode = 0;
  delete process.env.TEAMS_REGION;
});

const deps = { find: () => "/usr/bin/chromium", exists: () => true, platform: "linux" as const, arch: "x64" };
const byName = (checks: Check[], name: string) => checks.find(item => item.name === name)!;
const serverClock = (offsetMs = 0) => ({ "HEAD /": { status: 200, headers: { date: new Date(Date.now() + offsetMs).toUTCString() } } });

async function capture(work: () => Promise<void>): Promise<string[]> {
  const out: string[] = [];
  const log = console.log;
  console.log = (...args: unknown[]) => { out.push(args.map(String).join(" ")); };
  try { await work(); } finally { console.log = log; }
  return out;
}

async function healthyMachine() {
  await seedProfile(["chat", "csa", "spaces", "sharepoint"]);
  const dir = browserProfileDir("default");
  await prepareProfileDir(dir);
  await writeFile(join(dir, "Local State"), "{}");
}

test("a healthy machine passes every check", async () => {
  await healthyMachine();
  fake = installFakeTeams(serverClock());
  const checks = await runDoctor({}, deps);
  expect(checks.map(item => item.name)).toEqual([
    "version", "platform", "browser", "browser profile", "stored profile",
    "chat token", "csa token", "spaces token", "substrate token", "sharepoint token",
    "sharepoint host", "region", "clock skew",
  ]);
  expect(byName(checks, "version").detail).toBe(`teams-cli ${VERSION}`);
  expect(byName(checks, "platform").detail).toBe("linux x64");
  expect(byName(checks, "browser").detail).toBe("/usr/bin/chromium");
  expect(byName(checks, "chat token").detail).toContain("expires in");
  expect(checks.filter(item => item.status === "fail")).toEqual([]);
  expect(checks.filter(item => item.status === "warn").map(item => item.name)).toEqual([]);
  expect(fake.requests.map(item => `${item.method} ${item.path}`)).toEqual(["HEAD /"]);
  expect(fake.requests[0]!.headers.authorization).toBeUndefined();
});

test("token values never appear in the output", async () => {
  const tokens = await seedProfile(["chat", "csa", "spaces"]);
  fake = installFakeTeams(serverClock());
  const lines = await capture(() => authDoctor({}, deps));
  const json = await capture(() => authDoctor({ json: true }, deps));
  const text = [...lines, ...json].join("\n");
  for (const stored of Object.values(tokens)) expect(text).not.toContain(stored.token);
  expect(text).not.toContain("eyJ");
});

test("a machine with nothing set up fails, names the fix, and exits 1", async () => {
  fake = installFakeTeams(serverClock());
  const lines = await capture(() => authDoctor({}, { ...deps, find: () => { throw new Error("No browser found. Install Chrome."); } }));
  expect(process.exitCode).toBe(1);
  expect(lines.length).toBe(13);
  expect(lines.find(line => line.startsWith("fail  browser:"))).toContain("Fix: Install Chrome, Edge, Brave, or Chromium");
  expect(lines.find(line => line.startsWith("fail  stored profile:"))).toContain("Fix: Run: teams-cli auth login");
  expect(lines.find(line => line.startsWith("fail  chat token:"))).toContain("missing");
  expect(lines.find(line => line.startsWith("warn  browser profile:"))).toContain("auth login");
});

test("an expired core token warns when a renewal is possible, and fails when it is not", async () => {
  await healthyMachine();
  const past = Math.floor(Date.now() / 1000) - 3600;
  await seedProfile(["chat", "csa", "spaces"], past);
  const dir = browserProfileDir("default");
  await writeFile(join(dir, "Local State"), "{}");
  fake = installFakeTeams(serverClock());
  const renewable = await runDoctor({}, deps);
  expect(byName(renewable, "chat token").status).toBe("warn");
  expect(byName(renewable, "chat token").fix).toBe("Run: teams-cli auth renew");
  await resetLocalState();
  await seedProfile(["chat", "csa", "spaces"], past);
  const stuck = await runDoctor({}, deps);
  expect(byName(stuck, "chat token").status).toBe("fail");
  expect(byName(stuck, "chat token").detail).toContain("expired at");
});

test("the SharePoint host and the region come from the profile or the environment", async () => {
  await seedProfile(["chat"]);
  fake = installFakeTeams(serverClock());
  expect(byName(await runDoctor({}, deps), "sharepoint host").status).toBe("warn");
  const { name, profile } = await loadProfile();
  await saveProfile(name, { ...profile!, sharepointHost: "contoso-my.sharepoint.com", region: "emea" });
  const checks = await runDoctor({}, deps);
  expect(byName(checks, "sharepoint host")).toMatchObject({ status: "ok", detail: "contoso-my.sharepoint.com" });
  expect(byName(checks, "region").status).toBe("warn");
  process.env.TEAMS_REGION = "amer";
  expect(byName(await runDoctor({}, deps), "region")).toMatchObject({ status: "ok", detail: "amer" });
});

test("clock skew: ok when close, warn at 2 minutes, fail at 10 minutes", async () => {
  await healthyMachine();
  fake = installFakeTeams(serverClock(0));
  expect(byName(await runDoctor({}, deps), "clock skew").status).toBe("ok");
  fake.uninstall();
  fake = installFakeTeams(serverClock(-120_000));
  const warn = byName(await runDoctor({}, deps), "clock skew");
  expect(warn.status).toBe("warn");
  expect(warn.detail).toContain("ahead of Microsoft");
  fake.uninstall();
  fake = installFakeTeams(serverClock(600_000));
  const fail = byName(await runDoctor({}, deps), "clock skew");
  expect(fail).toMatchObject({ status: "fail" });
  expect(fail.detail).toContain("behind Microsoft");
  expect(fail.fix).toContain("clock");
});

test("offline skips the clock check without a failure", async () => {
  await healthyMachine();
  fake = installFakeTeams({ "HEAD /": "network" });
  const clock = byName(await runDoctor({}, deps), "clock skew");
  expect(clock.status).toBe("warn");
  expect(clock.detail).toContain("skipped");
  await capture(() => authDoctor({}, deps));
  expect(process.exitCode ?? 0).toBe(0);
});

test("a reply with no Date header skips the clock check", async () => {
  await healthyMachine();
  fake = installFakeTeams({ "HEAD /": { status: 200 } });
  expect(byName(await runDoctor({}, deps), "clock skew").detail).toContain("no Date header");
});

test("--json prints {checks:[{name,status,detail,fix}]}", async () => {
  await healthyMachine();
  fake = installFakeTeams(serverClock());
  const lines = await capture(() => authDoctor({ json: true }, deps));
  const parsed = JSON.parse(lines.join("\n"));
  expect(Object.keys(parsed)).toEqual(["checks"]);
  for (const item of parsed.checks) {
    expect(Object.keys(item).sort()).toEqual(["detail", "fix", "name", "status"]);
    expect(["ok", "warn", "fail"]).toContain(item.status);
  }
});

test("a named browser path that does not exist fails", async () => {
  await healthyMachine();
  fake = installFakeTeams(serverClock());
  const checks = await runDoctor({ browser: "/nope/chrome" }, { ...deps, find: flag => flag!, exists: () => false });
  expect(byName(checks, "browser")).toMatchObject({ status: "fail" });
  expect(byName(checks, "browser").detail).toContain("/nope/chrome");
});

test("the doctor writes nothing and starts no browser", async () => {
  fake = installFakeTeams(serverClock());
  await runDoctor({}, deps);
  expect(existsSync(configDir())).toBe(false);
});

test("the CLI runs auth doctor with --json against an unreachable host", async () => {
  const result = await runCli(["auth", "doctor", "--json"], { env: { TEAMS_HOST: "http://127.0.0.1:1", TEAMS_CLI_BROWSER: process.execPath } });
  const parsed = JSON.parse(result.stdout);
  expect(parsed.checks.find((item: Check) => item.name === "clock skew").status).toBe("warn");
  expect(parsed.checks.find((item: Check) => item.name === "stored profile").status).toBe("fail");
  expect(result.code).toBe(1);
});
