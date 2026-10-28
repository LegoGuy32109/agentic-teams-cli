import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

// Every test runs with no network. A test that needs HTTP installs test/fake-teams.ts.
globalThis.fetch = (async () => {
  throw new Error("Network is blocked in tests. Install the fake.");
}) as unknown as typeof fetch;

// Keep config and cache out of the real home directory.
const root = mkdtempSync(join(tmpdir(), "teams-cli-test-"));
process.env.XDG_CONFIG_HOME = join(root, "config");
process.env.XDG_CACHE_HOME = join(root, "cache");
for (const name of Object.keys(process.env)) {
  if (name.startsWith("TEAMS_")) delete process.env[name];
}
