import { afterEach, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { compareVersions, expectedChecksum, targetAsset, update, windowsScript } from "../src/commands/update.ts";
import { DISTRIBUTION, REPOSITORY } from "../src/version.ts";
import { runCli } from "./cli.ts";

const roots: string[] = [];
afterEach(async () => { for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true }); });

async function fixture(options: { tag?: string; checksum?: string; missing?: string; status?: number; badUrl?: boolean } = {}) {
  const root = await mkdtemp(join(tmpdir(), "teams-cli-update-")); roots.push(root);
  const executable = join(root, "teams-cli"); await writeFile(executable, "original");
  const binary = "new compiled binary";
  const tag = options.tag ?? "v1.1.0";
  const name = "teams-cli-linux-x64";
  const assets = [name, "SHA256SUMS"].filter(n => n !== options.missing).map(name => ({
    name, browser_download_url: `https://github.com/${REPOSITORY}/releases/download/${tag}/${name}`,
  }));
  if (options.badUrl) assets[0]!.browser_download_url = "https://example.com/binary";
  const calls: string[] = [];
  const fake = (async (input: string | URL | Request) => {
    const url = String(input); calls.push(url);
    if (options.status) return new Response("failure", { status: options.status });
    if (url.endsWith("/latest")) return Response.json({ tag_name: tag, assets });
    if (url.endsWith("/SHA256SUMS")) return new Response(`${options.checksum ?? createHash("sha256").update(binary).digest("hex")}  ${name}\n`);
    return new Response(binary);
  }) as typeof fetch;
  return { root, executable, calls, runtime: { fetch: fake, executable, platform: "linux", arch: "x64", distribution: "github" as const, version: "1.0.0" } };
}
const opts = { check: false, force: false, json: true };

test("check reads only release metadata and never replaces the executable", async () => {
  const f = await fixture();
  await update({ ...opts, check: true }, f.runtime);
  expect(f.calls).toHaveLength(1);
  expect(await readFile(f.executable, "utf8")).toBe("original");
  expect(await readdir(f.root)).toEqual(["teams-cli"]);
});

test("a verified public update replaces only its executable", async () => {
  const f = await fixture();
  await writeFile(join(f.root, "credentials.json"), "unchanged");
  await update(opts, f.runtime);
  expect(await readFile(f.executable, "utf8")).toBe("new compiled binary");
  expect(await readFile(join(f.root, "credentials.json"), "utf8")).toBe("unchanged");
  if (process.platform !== "win32") expect((await stat(f.executable)).mode & 0o777).toBe(0o755);
  expect((await readdir(f.root)).sort()).toEqual(["credentials.json", "teams-cli"]);
});

test.each([
  [{ checksum: "0".repeat(64) }, "checksum"],
  [{ missing: "SHA256SUMS" }, "SHA256SUMS"],
  [{ missing: "teams-cli-linux-x64" }, "teams-cli-linux-x64"],
  [{ status: 404 }, "404"],
  [{ status: 429 }, "429"],
  [{ badUrl: true }, "Unexpected release URL"],
  [{ tag: "v1.2.0-rc.1" }, "Invalid stable release"],
] as const)("an invalid release preserves the installed binary: %j", async (options, error) => {
  const f = await fixture(options);
  await expect(update(opts, f.runtime)).rejects.toThrow(error);
  expect(await readFile(f.executable, "utf8")).toBe("original");
  expect(await readdir(f.root)).toEqual(["teams-cli"]);
});

test("downgrades require force and equal versions download nothing", async () => {
  const f = await fixture({ tag: "v0.9.0" });
  await expect(update(opts, f.runtime)).rejects.toThrow("Refusing downgrade");
  expect(f.calls).toHaveLength(1);
  await update({ ...opts, force: true }, f.runtime);
  expect(await readFile(f.executable, "utf8")).toBe("new compiled binary");
  const same = await fixture({ tag: "v1.0.0" });
  await update(opts, same.runtime);
  expect(same.calls).toHaveLength(1);
});

test("local builds refuse replacement before any network request, including force", async () => {
  const f = await fixture();
  await expect(update({ ...opts, force: true }, { ...f.runtime, distribution: "source" })).rejects.toThrow("managed by work-portal");
  expect(f.calls).toHaveLength(0);
  await expect(update(opts, { ...f.runtime, executable: join(f.root, "bun") })).rejects.toThrow("compiled");
  expect(f.calls).toHaveLength(0);
  const cli = await runCli(["update"]);
  expect(cli.code).toBe(1);
  expect(cli.stderr).toContain(DISTRIBUTION === "source" ? "managed by work-portal" : "compiled");
});

test("platform mapping and checksums reject ambiguous input", () => {
  expect(targetAsset("win32", "x64")).toBe("teams-cli-windows-x64.exe");
  expect(targetAsset("darwin", "arm64")).toBe("teams-cli-darwin-arm64");
  expect(() => targetAsset("linux", "arm64")).toThrow("No release binary");
  expect(compareVersions("v1.10.0", "1.9.9")).toBe(1);
  expect(() => compareVersions("banana", "1.0.0")).toThrow();
  const line = `${"a".repeat(64)} *file.exe`;
  expect(expectedChecksum(line + "\r\n", "file.exe")).toBe("a".repeat(64));
  expect(() => expectedChecksum(line + "\n" + line, "file.exe")).toThrow("exactly one");
});

test("Windows replacement waits for the parent and treats special paths as literals", () => {
  const script = windowsScript("C:\\O'Brien %TEMP% & tools\\teams-cli.exe", "C:\\tools\\new", 123);
  expect(script).toContain("Wait-Process -Id 123");
  expect(script).toContain("'C:\\O''Brien %TEMP% & tools\\teams-cli.exe'");
  expect(script).toContain("Move-Item -LiteralPath $temporary");
  expect(script).toContain("$attempt -lt 30");
});
