// Builds one single-file binary per platform into dist/. Run: bun run build:all
// A target that fails does not stop the others. The exit code is 1 if any target failed.
import { mkdirSync, statSync } from "node:fs";
import { join } from "node:path";

const root = join(import.meta.dir, "..");
const dist = join(root, "dist");

export const TARGETS = [
  { target: "bun-linux-x64", file: "teams-cli-linux-x64" },
  { target: "bun-darwin-arm64", file: "teams-cli-darwin-arm64" },
  { target: "bun-darwin-x64", file: "teams-cli-darwin-x64" },
  { target: "bun-windows-x64", file: "teams-cli-windows-x64.exe" },
] as const;

if (import.meta.main) {
  mkdirSync(dist, { recursive: true });
  let failed = 0;
  for (const { target, file } of TARGETS) {
    const outfile = join(dist, file);
    const proc = Bun.spawn(["bun", "build", "--compile", `--target=${target}`, `--outfile=${outfile}`, "index.ts"], {
      cwd: root, stdout: "pipe", stderr: "pipe",
    });
    const [out, err] = await Promise.all([new Response(proc.stdout).text(), new Response(proc.stderr).text()]);
    if (await proc.exited === 0) {
      console.log(`ok    ${file} (${(statSync(outfile).size / 1_048_576).toFixed(1)} MB)`);
    } else {
      failed++;
      console.error(`fail  ${file}\n${(err || out).trim()}`);
    }
  }
  if (failed) process.exitCode = 1;
}
