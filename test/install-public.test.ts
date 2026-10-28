import { afterEach, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { chmod, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DISTRIBUTION } from "../src/version.ts";

const roots: string[] = [];
afterEach(async () => { for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true }); });

async function install(badChecksum: boolean) {
  const root = await mkdtemp(join(tmpdir(), "teams-public-install-")); roots.push(root);
  const bin = join(root, "fake-bin"); await mkdir(bin);
  const payload = "#!/bin/sh\necho fake-release\n";
  await writeFile(join(root, "payload"), payload);
  await writeFile(join(root, "sums"), `${badChecksum ? "0".repeat(64) : createHash("sha256").update(payload).digest("hex")}  teams-cli-linux-x64\n`);
  const curl = join(bin, "curl");
  await writeFile(curl, `#!/bin/sh
set -eu
output=
url=
while [ "$#" -gt 0 ]; do
  case "$1" in
    -o) output=$2; shift 2 ;;
    -w) shift 2 ;;
    https://*) url=$1; shift ;;
    *) shift ;;
  esac
done
case "$url" in
  */releases/latest) printf '%s' 'https://github.com/LegoGuy32109/agentic-teams-cli/releases/tag/v1.0.0' ;;
  */teams-cli-linux-x64) cp "$FAKE_ROOT/payload" "$output" ;;
  */SHA256SUMS) cp "$FAKE_ROOT/sums" "$output" ;;
  *) exit 99 ;;
esac
`);
  await chmod(curl, 0o755);
  const uname = join(bin, "uname");
  await writeFile(uname, '#!/bin/sh\ncase "$1" in -s) echo Linux ;; -m) echo x86_64 ;; esac\n');
  await chmod(uname, 0o755);
  const installed = join(root, ".local/bin"); await mkdir(installed, { recursive: true });
  const target = join(installed, "teams-cli"); await writeFile(target, "old binary");
  const script = join(import.meta.dir, "..", DISTRIBUTION === "source" ? "public/install.sh" : "install.sh");
  const child = Bun.spawn(["sh", script], {
    env: { HOME: root, PATH: `${bin}:/usr/bin:/bin`, FAKE_ROOT: root }, stdout: "pipe", stderr: "pipe",
  });
  const [stdout, stderr] = await Promise.all([new Response(child.stdout).text(), new Response(child.stderr).text()]);
  return { code: await child.exited, stdout, stderr, content: await readFile(target, "utf8"), payload };
}

test.skipIf(process.platform !== "linux")("public installer verifies the pinned release before replacement", async () => {
  const result = await install(false);
  expect(result.code).toBe(0);
  expect(result.content).toBe(result.payload);
  expect(result.stdout).toContain("Updates: teams-cli update");
});

test.skipIf(process.platform !== "linux")("public installer checksum failure preserves the existing installation", async () => {
  const result = await install(true);
  expect(result.code).toBe(1);
  expect(result.content).toBe("old binary");
  expect(result.stderr).toContain("Checksum mismatch");
});
