import { createHash } from "node:crypto";
import { readFile, rm } from "node:fs/promises";
import { join } from "node:path";
import { cacheDir, writeFileAtomic } from "./config.ts";

type Scope = { tenantId: string; mri: string; profile: string };
export type AccessDenial = { status: 403; checked_at: string; reason: string };
export const ACCESS_TTL_MS = 60 * 60_000;

function pathFor(scope: Scope, conversationId: string): string {
  const key = createHash("sha256").update(JSON.stringify([scope.tenantId, scope.mri, scope.profile, conversationId])).digest("hex");
  return join(cacheDir(), "access", `${key}.json`);
}

/** Only observed history denials are recorded. Meeting attendance metadata never proves loss of access. */
export async function rememberDenial(scope: Scope, conversationId: string, now = Date.now()): Promise<void> {
  await writeFileAtomic(pathFor(scope, conversationId), JSON.stringify({ status: 403, checked_at: new Date(now).toISOString(),
    reason: "Teams denied history access (403). Check meeting membership and the signed-in tenant." }), 0o600);
}

export async function clearDenial(scope: Scope, conversationId: string): Promise<void> {
  await rm(pathFor(scope, conversationId), { force: true });
}

export async function knownDenial(scope: Scope, conversationId: string, now = Date.now()): Promise<AccessDenial | undefined> {
  try {
    const value = JSON.parse(await readFile(pathFor(scope, conversationId), "utf8"));
    const age = now - Date.parse(value.checked_at);
    if (value.status === 403 && age >= 0 && age < ACCESS_TTL_MS && typeof value.reason === "string") return value;
  } catch { /* No recent observed denial. */ }
  return undefined;
}
