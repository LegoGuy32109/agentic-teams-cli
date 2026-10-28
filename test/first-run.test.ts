import { afterEach, beforeEach, expect, test } from "bun:test";
import { configureFirstRun, isYes, setFirstRunHooks, SIGN_IN_QUESTION } from "../src/first-run.ts";
import { identity, requireToken } from "../src/tokens.ts";
import { seedProfile, resetLocalState } from "./fake-teams.ts";
import { fakeJwt } from "./jwt.ts";
import { runCli } from "./cli.ts";

beforeEach(resetLocalState);
afterEach(() => {
  configureFirstRun({ enabled: false });
  setFirstRunHooks();
  delete process.env.TEAMS_CHAT_TOKEN;
  delete process.env.TEAMS_TOKEN_COMMAND;
});

/** Hooks for a person at a terminal. `login` seeds a profile, as a real sign-in would save one. */
function terminal(answer: string, browser?: { seen?: Array<string | undefined> }) {
  const asked: string[] = [];
  let logins = 0;
  setFirstRunHooks({
    isTerminal: () => true,
    ask: async question => { asked.push(question); return answer; },
    login: async options => { logins++; browser?.seen?.push(options.browser); await seedProfile(); },
  });
  return { asked, logins: () => logins };
}

test("a person at a terminal is asked, and a yes signs in and continues the command", async () => {
  configureFirstRun({ enabled: true, browser: "/opt/chrome" });
  const seen: Array<string | undefined> = [];
  const hooks = terminal("", { seen });
  const who = await identity();
  expect(who.upn).toBe("tester@example.test");
  expect(hooks.asked).toEqual([SIGN_IN_QUESTION]);
  expect(SIGN_IN_QUESTION).toBe("No Teams sign-in on this machine. Sign in now in a browser window? [Y/n] ");
  expect(hooks.logins()).toBe(1);
  expect(seen).toEqual(["/opt/chrome"]);
  // The tokens are there now, and nobody is asked again.
  expect((await requireToken("chat")).source).toBe("store");
  expect(hooks.asked.length).toBe(1);
});

test("an explicit yes also signs in", async () => {
  configureFirstRun({ enabled: true });
  const hooks = terminal("yes");
  await requireToken("chat");
  expect(hooks.logins()).toBe(1);
});

test("a no ends the command with the login command and starts no browser", async () => {
  configureFirstRun({ enabled: true });
  const hooks = terminal("n");
  await expect(identity()).rejects.toThrow("Run: teams-cli auth login");
  await expect(requireToken("chat")).rejects.toThrow("Run: teams-cli auth login");
  expect(hooks.logins()).toBe(0);
});

test("parallel callers share one prompt", async () => {
  configureFirstRun({ enabled: true });
  const hooks = terminal("y");
  await Promise.all([identity(), identity(), requireToken("csa")]);
  expect(hooks.asked.length).toBe(1);
  expect(hooks.logins()).toBe(1);
});

test("without a terminal there is no prompt and the ordinary error stays", async () => {
  configureFirstRun({ enabled: true });
  let asked = 0;
  setFirstRunHooks({ isTerminal: () => false, ask: async () => { asked++; return "y"; }, login: async () => { asked++; } });
  await expect(identity()).rejects.toThrow("Teams identity is unknown. Run: teams-cli auth login");
  await expect(requireToken("chat")).rejects.toThrow("No chat token. Run: teams-cli auth login");
  expect(asked).toBe(0);
});

test("with --json (offer off) there is no prompt even at a terminal", async () => {
  configureFirstRun({ enabled: false });
  const hooks = terminal("y");
  await expect(identity()).rejects.toThrow("Teams identity is unknown");
  expect(hooks.asked).toEqual([]);
});

test("a token from the environment or a command means no prompt", async () => {
  configureFirstRun({ enabled: true });
  const hooks = terminal("y");
  process.env.TEAMS_CHAT_TOKEN = fakeJwt();
  await expect(requireToken("csa")).rejects.toThrow("No csa token. Run: teams-cli auth login");
  delete process.env.TEAMS_CHAT_TOKEN;
  process.env.TEAMS_TOKEN_COMMAND = "echo {}";
  await expect(requireToken("csa")).rejects.toThrow("No csa token");
  expect(hooks.asked).toEqual([]);
});

test("a stored profile means no prompt", async () => {
  configureFirstRun({ enabled: true });
  await seedProfile(["chat"]);
  const hooks = terminal("y");
  await expect(requireToken("csa")).rejects.toThrow("No csa token. Run: teams-cli auth login");
  expect(hooks.asked).toEqual([]);
});

test("answers: empty, y and yes are yes, anything else is no", () => {
  for (const yes of ["", "\n", "y", "Y", " yes ", "YES"]) expect(isYes(yes)).toBe(true);
  for (const no of ["n", "no", "nope", "maybe"]) expect(isYes(no)).toBe(false);
});

test("the CLI in a pipe prints the ordinary error and never waits for input", async () => {
  const piped = await runCli(["history", "me"], { input: "y\n" });
  expect(piped.code).toBe(1);
  expect(piped.stderr).toContain("Teams identity is unknown. Run: teams-cli auth login");
  expect(piped.stderr).not.toContain("Sign in now");
  const json = await runCli(["history", "me", "--json"]);
  expect(json.code).toBe(1);
  expect(json.stderr + json.stdout).not.toContain("Sign in now");
});
