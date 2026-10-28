import { beforeEach, expect, test } from "bun:test";
import { mkdir, mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { TARGETS } from "../scripts/build-all.ts";

const script = join(import.meta.dir, "..", "install", "install.sh");
let home: string;
let dist: string;

beforeEach(async () => {
  home = await mkdtemp(join(tmpdir(), "teams-cli-install-"));
  dist = join(home, "built");
  await mkdir(dist);
  for (const { file } of TARGETS) await writeFile(join(dist, file), `#!/bin/sh\necho ${file}\n`);
});

async function install(args: string[], env: Record<string, string> = {}) {
  const proc = Bun.spawn(["sh", script, ...args], {
    stdout: "pipe", stderr: "pipe", cwd: home,
    env: { HOME: home, PATH: "/usr/bin:/bin", ...env },
  });
  const [stdout, stderr] = await Promise.all([new Response(proc.stdout).text(), new Response(proc.stderr).text()]);
  return { stdout, stderr, code: await proc.exited };
}

test("build:all lists the four binaries", () => {
  expect(TARGETS.map(item => item.file)).toEqual(["teams-cli-linux-x64", "teams-cli-darwin-arm64", "teams-cli-darwin-x64", "teams-cli-windows-x64.exe"]);
  expect(TARGETS.map(item => item.target)).toEqual(["bun-linux-x64", "bun-darwin-arm64", "bun-darwin-x64", "bun-windows-x64"]);
});

test("the Unix installer copies the binary for the platform to ~/.local/bin and says what is next", async () => {
  for (const platform of ["linux-x64", "darwin-arm64", "darwin-x64"]) {
    const result = await install([dist], { TEAMS_CLI_PLATFORM: platform });
    expect(result.code).toBe(0);
    const target = join(home, ".local", "bin", "teams-cli");
    expect((await stat(target)).mode & 0o111).not.toBe(0);
    expect(await readFile(target, "utf8")).toContain(`teams-cli-${platform}`);
    expect(result.stdout.trimEnd().split("\n").at(-1)).toBe("Next: teams-cli auth login");
  }
});

test("it tells the user how to add the folder to PATH only when it is missing", async () => {
  const missing = await install([dist], { TEAMS_CLI_PLATFORM: "linux-x64" });
  expect(missing.stdout).toContain("is not on your PATH");
  expect(missing.stdout).toContain('export PATH="$HOME/.local/bin:$PATH"');
  const present = await install([dist], { TEAMS_CLI_PLATFORM: "linux-x64", PATH: `${join(home, ".local", "bin")}:/usr/bin:/bin` });
  expect(present.stdout).not.toContain("not on your PATH");
  expect(present.stdout.trimEnd().split("\n").at(-1)).toBe("Next: teams-cli auth login");
});

test("without an argument it uses ./dist", async () => {
  await mkdir(join(home, "dist"));
  await writeFile(join(home, "dist", "teams-cli-linux-x64"), "#!/bin/sh\necho from-dist\n");
  const result = await install([], { TEAMS_CLI_PLATFORM: "linux-x64" });
  expect(result.code).toBe(0);
  expect(await readFile(join(home, ".local", "bin", "teams-cli"), "utf8")).toContain("from-dist");
});

test("a missing binary, a platform with no binary, and a Windows shell fail with a message", async () => {
  const empty = join(home, "empty");
  await mkdir(empty);
  const missing = await install([empty], { TEAMS_CLI_PLATFORM: "linux-x64" });
  expect(missing.code).toBe(1);
  expect(missing.stderr).toContain("teams-cli-linux-x64 was not found");
  expect(missing.stdout).not.toContain("Next:");
  const odd = await install([dist], { TEAMS_CLI_PLATFORM: "freebsd-x64" });
  expect(odd.code).toBe(1);
  expect(odd.stderr).toContain("no binary for freebsd-x64");
});

test("the Windows installer targets %LOCALAPPDATA%, downloads nothing, and ends with the next step", async () => {
  const text = await readFile(join(import.meta.dir, "..", "install", "install.ps1"), "utf8");
  expect(text).toContain("teams-cli-windows-x64.exe");
  expect(text).toContain("$env:LOCALAPPDATA");
  expect(text).toContain("teams-cli\\bin");
  expect(text).toContain("SetEnvironmentVariable('Path'");
  expect(text.trimEnd().split("\n").at(-1)).toBe("Write-Host 'Next: teams-cli auth login'");
  for (const forbidden of ["Invoke-WebRequest", "Invoke-RestMethod", "curl", "wget", "DownloadFile", "RunAs"]) expect(text).not.toContain(forbidden);
  const unix = await readFile(script, "utf8");
  for (const forbidden of ["curl ", "wget ", "sudo "]) expect(unix).not.toContain(forbidden);
});
