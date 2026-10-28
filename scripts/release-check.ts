import { VERSION, DISTRIBUTION } from "../src/version.ts";
import pkg from "../package.json";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";

const manifest = JSON.parse(await readFile("upstream-manifest.json", "utf8"));
const sha256 = (data: string | Uint8Array) => createHash("sha256").update(data).digest("hex");
if (manifest.schema !== 1 || sha256(JSON.stringify(manifest.files)) !== manifest.fingerprint) throw new Error("Invalid upstream manifest.");
for (const [path, expected] of Object.entries(manifest.files)) {
  if (path.startsWith("/") || path.split(/[\\/]/).includes("..")) throw new Error("Unsafe manifest path.");
  if (sha256(await readFile(path)) !== expected) throw new Error(`Export differs at ${path}. Port changes upstream and export again.`);
}

if (pkg.version !== VERSION) throw new Error("package.json and src/version.ts versions differ.");
if (DISTRIBUTION !== "github") throw new Error("Release requires a public export.");
if (!/^\d+\.\d+\.\d+$/.test(VERSION)) throw new Error("Release requires a stable version.");
if (process.env.GITHUB_REF_TYPE === "tag" && process.env.GITHUB_REF_NAME !== `v${VERSION}`) {
  throw new Error(`Release tag must be v${VERSION}.`);
}
for (const command of [["bun", "run", "typecheck"], ["bun", "test"]]) {
  const child = Bun.spawn(command, { stdout: "inherit", stderr: "inherit" });
  if (await child.exited !== 0) throw new Error(`${command.join(" ")} failed.`);
}
console.log(`Release checks passed for v${VERSION}.`);
