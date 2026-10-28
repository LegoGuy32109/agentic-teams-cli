import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { allowedFlags, FLAG_SPECS } from "../src/flags.ts";
import { commandHelp, fullHelp, helpJson, topicHelp } from "../src/help.ts";
import { COMMANDS, TOPICS } from "../src/spec.ts";
import { MAX_FILES, MAX_UPLOAD_BYTES, mb } from "../src/limits.ts";
import { withTable } from "../scripts/docs.ts";
import { runCli } from "./cli.ts";

// These tests keep the help text, the README, and the parser from drifting apart. All three come from src/spec.ts.
// `--help` and `--version` work alone, before any command, so no command lists them.
const ALIASES = new Set(["help", "version"]);
const flagsIn = (text: string) => new Set([...text.matchAll(/--([a-z][a-z-]*)/g)].map(match => match[1]!).filter(name => !ALIASES.has(name)));

describe("help stays in step with the CLI", () => {
  test("a usage line names only flags that its command accepts", () => {
    for (const spec of COMMANDS) {
      const allowed = new Set<string>(allowedFlags(spec.key));
      for (const entry of spec.entries) {
        for (const name of flagsIn(entry.usage)) expect({ key: spec.key, flag: name, accepted: allowed.has(name) }).toEqual({ key: spec.key, flag: name, accepted: true });
      }
    }
  });

  test("every flag a command accepts is in its help", () => {
    for (const spec of COMMANDS) {
      const documented = flagsIn(spec.entries.map(entry => `${entry.usage} ${entry.text}`).join(" "));
      for (const name of allowedFlags(spec.key).filter(name => name !== "profile")) {
        expect({ key: spec.key, flag: name, documented: documented.has(name) }).toEqual({ key: spec.key, flag: name, documented: true });
      }
    }
  });

  test("every flag the parser knows appears in the full help", () => {
    const help = fullHelp();
    for (const name of Object.keys(FLAG_SPECS)) expect(help).toContain(`--${name}`);
  });

  test("every command that index.ts dispatches has help", () => {
    const source = readFileSync(join(import.meta.dir, "..", "index.ts"), "utf8");
    const dispatched = [...source.matchAll(/^\s+case "([a-z-]+)":/gm)].map(match => match[1]!);
    expect(dispatched.length).toBeGreaterThan(20);
    for (const name of dispatched) expect({ name, help: commandHelp([name]) !== undefined }).toEqual({ name, help: true });
  });

  test("every example passes the real parser", async () => {
    const examples = COMMANDS.flatMap(spec => (spec.examples ?? []).map(example => ({ key: spec.key, args: example.args })));
    expect(examples.length).toBeGreaterThan(10);
    const results = await Promise.all(examples.map(async example => {
      const result = await runCli(example.args, { env: { TEAMS_CLI_CHECK_ARGS: "1" } });
      return { ...example, code: result.code, error: result.stderr.trim() };
    }));
    for (const result of results) expect(result).toEqual({ ...result, code: 0, error: "" });
  });

  test("limits in the help come from the code", () => {
    expect(fullHelp()).toContain(`Attach up to ${MAX_FILES} files (${mb(MAX_UPLOAD_BYTES)} each)`);
  });

  test("the README command table is current. Run: bun run docs", () => {
    const readme = readFileSync(join(import.meta.dir, "..", "README.md"), "utf8");
    expect(withTable(readme) === readme).toBe(true);
  });

  test("help --json lists every command with its flags", async () => {
    const result = await runCli(["help", "--json"]);
    expect(result.code).toBe(0);
    const parsed = JSON.parse(result.stdout);
    expect(parsed.commands.map((item: { command: string }) => item.command)).toEqual(COMMANDS.map(spec => spec.key));
    const send = parsed.commands.find((item: { command: string }) => item.command === "send");
    expect(send.flags).toContainEqual({ name: "--thread-ts", takes_value: true });
    expect(send.flags).toContainEqual({ name: "--profile", takes_value: true });
    expect((await runCli(["react", "--help", "--json"])).stdout).toBe(`${JSON.stringify(helpJson(["react"]), null, 2)}\n`);
  });
});

describe("help for one command", () => {
  test("<command> --help prints that command, its references, and runs nothing", async () => {
    const result = await runCli(["send", "--help"]);
    expect(result.code).toBe(0);
    expect(result.stdout).toContain("send <conversation> [text...] --file PATH");
    expect(result.stdout).toContain("<conversation>");
    expect(result.stdout).not.toContain("files delete");
  });

  test("a group prints each subcommand, and a subcommand prints only itself", async () => {
    const group = await runCli(["files", "--help"]);
    expect(group.stdout).toContain("files download");
    expect(group.stdout).toContain("files delete");
    const one = await runCli(["files", "delete", "--help"]);
    expect(one.stdout).toContain("files delete");
    expect(one.stdout).not.toContain("files download");
    expect(one.stdout).toContain("<conversation-id>#<message-id>, as history");
  });

  test("help <command> is the same as <command> --help", async () => {
    expect((await runCli(["help", "react"])).stdout).toBe((await runCli(["react", "--help"])).stdout);
  });

  test("help for an unknown command fails", async () => {
    const result = await runCli(["help", "nope"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("No help for `nope`.");
  });

  test("a usage error points at the command's own help", async () => {
    const result = await runCli(["react", "me"]);
    expect(result.stderr).toContain("Run `teams-cli react --help` for this command");
  });

  test("with no profile, a command points at auth login", async () => {
    const result = await runCli(["conversations"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("Run: teams-cli auth login");
  });
});

describe("help writing", () => {
  const GATE = "Could the recipient write a meaningful reply directly from this message, without first asking what you want?";
  const topicText = () => TOPICS.find(topic => topic.key === "writing")!.sections.flatMap(section => section.points).join("\n");

  test("prints the quality-gate sentence word for word", async () => {
    const result = await runCli(["help", "writing"]);
    expect(result.code).toBe(0);
    expect(result.stdout.replace(/\s+/g, " ")).toContain(GATE);
    expect(result.stdout).toBe(`${topicHelp("writing")}\n`);
    expect(result.stdout.split("\n").every(line => line.length <= 120)).toBe(true);
  });

  test("the full help points to it in one line", () => {
    expect(fullHelp()).toContain("Writing a message someone must act on? Run: teams-cli help writing");
  });

  test("every command and flag that the topic names exists", () => {
    const text = topicText();
    const keys = COMMANDS.map(spec => spec.key);
    const named = [...text.matchAll(/teams-cli ([a-z][a-z-]*)(?: ([a-z][a-z-]*))?/g)];
    expect(named.length).toBeGreaterThan(0);
    for (const match of named) {
      const known = keys.includes(match[1]!) || keys.includes(`${match[1]} ${match[2]}`) || keys.some(key => key.startsWith(`${match[1]} `));
      expect({ command: match[0], known }).toEqual({ command: match[0], known: true });
    }
    const flags = flagsIn(text);
    expect(flags.size).toBeGreaterThan(4);
    const send = new Set<string>([...allowedFlags("send"), ...allowedFlags("card send")]);
    for (const name of flags) {
      expect({ flag: name, parsed: name in FLAG_SPECS, onSend: send.has(name) }).toEqual({ flag: name, parsed: true, onSend: true });
    }
  });

  test("--json has the topic, the title, and sections with points", async () => {
    const result = await runCli(["help", "writing", "--json"]);
    expect(result.code).toBe(0);
    const parsed = JSON.parse(result.stdout);
    expect(Object.keys(parsed)).toEqual(["topic", "title", "sections"]);
    expect(parsed.topic).toBe("writing");
    expect(typeof parsed.title).toBe("string");
    expect(parsed.sections.length).toBeGreaterThan(3);
    for (const section of parsed.sections) {
      expect(Object.keys(section)).toEqual(["heading", "points"]);
      expect(section.points.length).toBeGreaterThan(0);
    }
  });

  test("help --json lists the topic names, and an unknown topic still fails", async () => {
    const parsed = JSON.parse((await runCli(["help", "--json"])).stdout);
    expect(parsed.topics.map((item: { topic: string }) => item.topic)).toEqual(TOPICS.map(topic => topic.key));
    const result = await runCli(["help", "nope"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("No help for `nope`.");
  });
});
