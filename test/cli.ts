import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const indexPath = join(import.meta.dir, "..", "index.ts");

/** Runs the CLI in a subprocess with isolated config and cache directories. */
export async function runCli(args: string[], options: { env?: Record<string, string>; input?: string; home?: string } = {}) {
  const home = options.home ?? mkdtempSync(join(tmpdir(), "teams-cli-run-"));
  const proc = Bun.spawn([process.execPath, indexPath, ...args], {
    stdin: options.input === undefined ? "ignore" : new Blob([options.input]),
    stdout: "pipe",
    stderr: "pipe",
    env: { PATH: process.env.PATH ?? "", XDG_CONFIG_HOME: join(home, "config"), XDG_CACHE_HOME: join(home, "cache"), ...(options.env ?? {}) },
  });
  const [stdout, stderr] = await Promise.all([new Response(proc.stdout).text(), new Response(proc.stderr).text()]);
  return { code: await proc.exited, stdout, stderr, home };
}
