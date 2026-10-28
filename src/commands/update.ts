import { createHash, randomUUID } from "node:crypto";
import { chmod, rename, rm, writeFile } from "node:fs/promises";
import { basename, dirname, join } from "node:path";
import { DISTRIBUTION, REPOSITORY, VERSION } from "../version.ts";

type Asset = { name: string; browser_download_url: string };
type Release = { tag_name: string; assets: Asset[]; draft?: boolean; prerelease?: boolean };
type Runtime = {
  fetch: typeof fetch; executable: string; platform: string; arch: string;
  distribution: "source" | "github"; version: string;
};

export function compareVersions(a: string, b: string): number {
  const parse = (s: string) => {
    if (!/^v?(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(s)) throw new Error(`Invalid stable release version: ${s}.`);
    const parts = s.replace(/^v/, "").split(".").map(Number);
    if (parts.some(n => !Number.isSafeInteger(n))) throw new Error("Release version exceeds the supported range.");
    return parts;
  };
  const left = parse(a), right = parse(b);
  for (let i = 0; i < 3; i++) if (left[i] !== right[i]) return left[i]! > right[i]! ? 1 : -1;
  return 0;
}

export function targetAsset(platform: string, arch: string): string {
  const suffix = platform === "win32" ? "windows" : platform;
  if (!["linux-x64", "darwin-arm64", "darwin-x64", "windows-x64"].includes(`${suffix}-${arch}`)) {
    throw new Error(`No release binary for ${platform}/${arch}.`);
  }
  return `teams-cli-${suffix}-${arch}${platform === "win32" ? ".exe" : ""}`;
}

export function expectedChecksum(text: string, name: string): string {
  const matches = text.split(/\r?\n/).map(line => line.match(/^([a-fA-F0-9]{64})\s+\*?(.+)$/))
    .filter(match => match?.[2] === name);
  if (matches.length !== 1) throw new Error(`SHA256SUMS must contain exactly one checksum for ${name}.`);
  return matches[0]![1]!.toLowerCase();
}

function assetUrl(asset: Asset, tag: string): string {
  const url = new URL(asset.browser_download_url);
  if (url.origin !== "https://github.com" || url.username || url.password || url.search || url.hash ||
      url.pathname !== `/${REPOSITORY}/releases/download/${tag}/${asset.name}`) {
    throw new Error(`Unexpected release URL for ${asset.name}.`);
  }
  return url.href;
}

// Quote paths as PowerShell literals. EncodedCommand avoids cmd.exe path expansion.
export function windowsScript(target: string, temporary: string, pid: number): string {
  const literal = (s: string) => `'${s.replaceAll("'", "''")}'`;
  return `$ErrorActionPreference = 'Stop'
Wait-Process -Id ${pid} -ErrorAction SilentlyContinue
$target = ${literal(target)}
$temporary = ${literal(temporary)}
for ($attempt = 0; $attempt -lt 30; $attempt++) {
  try { Move-Item -LiteralPath $temporary -Destination $target -Force; exit 0 }
  catch { Start-Sleep -Milliseconds 500 }
}
exit 1
`;
}

export async function update(opts: { check: boolean; force: boolean; json: boolean }, injected?: Partial<Runtime>) {
  const runtime: Runtime = {
    fetch: globalThis.fetch, executable: process.execPath, platform: process.platform, arch: process.arch,
    distribution: DISTRIBUTION, version: VERSION, ...injected,
  };
  if (!opts.check) {
    if (runtime.distribution !== "github") throw new Error("This build is managed by work-portal. Run: apps/wp build teams-cli. Use update --check to inspect public releases.");
    if (/^bun(?:\.exe)?$/i.test(basename(runtime.executable))) throw new Error("Run update from an installed compiled teams-cli binary, not bun run.");
  }
  const request = async (url: string) => {
    const response = await runtime.fetch(url, {
      headers: { Accept: "application/vnd.github+json", "User-Agent": "teams-cli" }, signal: AbortSignal.timeout(120_000),
    });
    if (!response.ok) throw new Error(`GitHub update request failed (${response.status}). Check the connection and whether a public release exists.`);
    return response;
  };
  const release = await (await request(`https://api.github.com/repos/${REPOSITORY}/releases/latest`)).json() as Release;
  const comparison = compareVersions(release.tag_name, runtime.version);
  if (release.draft || release.prerelease || !Array.isArray(release.assets)) throw new Error("Expected a published stable release with assets.");
  const output = (data: Record<string, unknown>, message: string) => console.log(opts.json || opts.check ? JSON.stringify(data, null, 2) : message);
  const state = { current: runtime.version, latest: release.tag_name, updateAvailable: comparison > 0 };
  if (opts.check) return output(state, "");
  if (comparison === 0) return output({ ...state, updated: false }, `Already up to date (${runtime.version}).`);
  if (comparison < 0 && !opts.force) throw new Error(`Refusing downgrade from ${runtime.version} to ${release.tag_name}. Use --force to allow it.`);
  const name = targetAsset(runtime.platform, runtime.arch);
  const oneAsset = (name: string) => {
    const found = release.assets.filter(asset => asset.name === name);
    if (found.length !== 1) throw new Error(`Release must contain exactly one ${name} asset.`);
    return found[0]!;
  };
  const binaryUrl = assetUrl(oneAsset(name), release.tag_name);
  const sumsUrl = assetUrl(oneAsset("SHA256SUMS"), release.tag_name);
  const [binary, sums] = await Promise.all([
    request(binaryUrl).then(r => r.arrayBuffer()).then(b => new Uint8Array(b)),
    request(sumsUrl).then(r => r.text()),
  ]);
  if (!binary.length) throw new Error("Release binary is empty. Existing binary was not changed.");
  if (createHash("sha256").update(binary).digest("hex") !== expectedChecksum(sums, name)) {
    throw new Error("Downloaded binary checksum did not match SHA256SUMS. Existing binary was not changed.");
  }
  const temporary = join(dirname(runtime.executable), `.teams-cli-update-${randomUUID()}.new`);
  try {
    await writeFile(temporary, binary, { mode: 0o755, flag: "wx" });
    if (runtime.platform === "win32") {
      const encoded = Buffer.from(windowsScript(runtime.executable, temporary, process.pid), "utf16le").toString("base64");
      const child = Bun.spawn(["powershell.exe", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded], {
        stdin: "ignore", stdout: "ignore", stderr: "ignore", windowsHide: true,
      });
      child.unref();
      return output({ ...state, scheduled: true, updated: false }, `Update to ${release.tag_name} is scheduled after teams-cli exits. Run teams-cli version to confirm.`);
    }
    await chmod(temporary, 0o755);
    await rename(temporary, runtime.executable);
    output({ ...state, updated: true }, `Updated teams-cli from ${runtime.version} to ${release.tag_name}.`);
  } catch (error) {
    await rm(temporary, { force: true }).catch(() => {});
    throw new Error(`Verified update could not replace the binary: ${error instanceof Error ? error.message : error}`);
  }
}
