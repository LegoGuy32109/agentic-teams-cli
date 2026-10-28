import { describe, expect, test } from "bun:test";
import { runCli } from "./cli.ts";
import pkg from "../package.json";

const COMMANDS = ["auth", "auth status", "logout", "config path", "conversations", "users find", "history", "send", "ledger", "version", "help"];

describe("CLI shell", () => {
  test("version prints the version", async () => {
    const result = await runCli(["version"]);
    expect(result.code).toBe(0);
    expect(result.stdout.trim()).toBe(pkg.version);
  });

  test("help lists every 1.0 command", async () => {
    const result = await runCli(["--help"]);
    expect(result.code).toBe(0);
    for (const name of COMMANDS.slice(0, 9)) expect(result.stdout).toContain(name);
    expect(result.stdout).toContain("help [command] [--json] | <command> --help");
  });

  test("update has help without loading Teams credentials", async () => {
    const result = await runCli(["update", "--help"]);
    expect(result.code).toBe(0);
    expect(result.stdout).toContain("update [--check] [--force] [--json]");
    expect(result.stdout).toContain("SHA256SUMS");
  });

  test("unknown flag fails with the help pointer", async () => {
    const result = await runCli(["version", "--nope"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("Unknown flag `--nope`.");
    expect(result.stderr).toContain("teams-cli --help");
  });

  test("a flag from another command is rejected", async () => {
    const result = await runCli(["conversations", "--top=3"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("Flag `--top` is not valid with `teams-cli conversations`.");
  });

  test("a value flag without a value fails", async () => {
    const result = await runCli(["history", "me", "--top"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("--top requires a value.");
  });

  test("wrong positional count fails with usage", async () => {
    const result = await runCli(["history"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("Usage: teams-cli history <conversation>");
  });

  test("config path prints a path and no token", async () => {
    const result = await runCli(["config", "path"]);
    expect(result.code).toBe(0);
    expect(result.stdout.trim().endsWith("credentials.json")).toBe(true);
  });

  test("a value on a boolean flag is rejected", async () => {
    const result = await runCli(["conversations", "--json=yes"]);
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("Flag `--json` does not take a value.");
  });
});
