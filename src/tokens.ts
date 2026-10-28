import { loadProfile, type TokenKind } from "./config.ts";

export const KINDS: TokenKind[] = ["chat", "csa", "spaces", "substrate", "sharepoint"];

/** Maps a token audience to its kind. The audience is the only sure way to tell tokens apart. */
export const KIND_BY_AUD: Record<string, TokenKind> = {
  "https://ic3.teams.office.com": "chat",
  "https://chatsvcagg.teams.microsoft.com": "csa",
  "https://api.spaces.skype.com": "spaces",
  "66a88757-258c-4c72-893c-3e8bed4d6899": "substrate",
  "https://substrate.office.com": "substrate",
  "00000003-0000-0ff1-ce00-000000000000": "sharepoint",
};

export const UNLOCKS: Record<TokenKind, string> = {
  chat: "history, send, auth status",
  csa: "conversations, and name resolution for Team/Channel, #name, and @Name",
  spaces: "users find, and @{Name} mentions",
  substrate: "search",
  sharepoint: "files download",
};

const ENV_NAME: Record<TokenKind, string> = {
  chat: "TEAMS_CHAT_TOKEN",
  csa: "TEAMS_CSA_TOKEN",
  spaces: "TEAMS_SPACES_TOKEN",
  substrate: "TEAMS_SUBSTRATE_TOKEN",
  sharepoint: "TEAMS_SHAREPOINT_TOKEN",
};

/** True when an environment variable or TEAMS_TOKEN_COMMAND supplies tokens, so no sign-in is needed. */
export function externalTokenSource(): boolean {
  return !!process.env.TEAMS_TOKEN_COMMAND?.trim() || Object.values(ENV_NAME).some(name => !!process.env[name]);
}

export type Claims = {
  aud?: string | string[];
  exp?: number;
  oid?: string;
  tid?: string;
  upn?: string;
  name?: string;
};

export type ResolvedToken = {
  kind: TokenKind;
  token: string;
  exp: number;
  source: "environment" | "command" | "store";
  claims?: Claims;
};

/** Decodes a JWT payload. It does not verify the signature. Returns undefined for anything else. */
export function decodeJwt(token: string): Claims | undefined {
  const parts = token.split(".");
  if (parts.length !== 3 || !parts[1]) return undefined;
  try {
    const parsed = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"));
    return parsed && typeof parsed === "object" ? parsed as Claims : undefined;
  } catch {
    return undefined;
  }
}

/** Returns the kind for a JWT by its audience, or undefined for an unknown audience. */
export function classifyToken(token: string): { kind?: TokenKind; claims?: Claims } {
  const claims = decodeJwt(token);
  if (!claims) return {};
  const audiences = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
  const kind = audiences.map(aud => aud ? KIND_BY_AUD[aud] : undefined).find(Boolean);
  return { kind, claims };
}

/** Removes anything that looks like a token or a bearer header from text. */
export function redact(text: string): string {
  // The card router URL carries its own key. It never reaches output, so the exact value and any Workflows URL are removed.
  const router = process.env.TEAMS_CARD_ROUTER;
  if (router) text = text.split(router).join("<card-router-url>");
  return text
    .replace(/https?:\/\/[^\s"'<>]*\.(?:powerplatform\.com|logic\.azure\.com)[^\s"'<>]*/gi, "<card-router-url>")
    .replace(/eyJ[A-Za-z0-9_-]*\.[A-Za-z0-9_-]*\.[A-Za-z0-9_.-]*/g, "<token>")
    .replace(/Bearer\s+[A-Za-z0-9._-]{20,}/gi, "Bearer <token>");
}

export function expiredMessage(kind: TokenKind, exp: number): string {
  return `The ${kind} token expired at ${new Date(exp * 1000).toLocaleString()}. Supply a new one: teams-cli auth`;
}

/**
 * The message for a 401. It says the token expired only when its `exp` is in the past. Otherwise
 * the service rejected a token that is still valid by its own clock.
 */
export function authFailureMessage(kind: TokenKind, exp: number, now = Date.now()): string {
  if (exp > 0 && exp * 1000 < now) return expiredMessage(kind, exp);
  return `Teams rejected the ${kind} token (401). It may be revoked, signed out, or for another service. Supply a new one: teams-cli auth`;
}

/** Fails when a token supplied for one kind has the audience of another kind. Unknown audiences pass. */
export function checkAudience(kind: TokenKind, claims: Claims | undefined): void {
  const audiences = Array.isArray(claims?.aud) ? claims.aud : [claims?.aud];
  for (const aud of audiences) {
    const other = aud ? KIND_BY_AUD[aud] : undefined;
    if (other && other !== kind) {
      throw new Error(`The token given as ${kind} is for the ${other} service. Supply the right token: teams-cli auth`);
    }
  }
}

export function formatRemaining(ms: number): string {
  const minutes = Math.floor(ms / 60_000);
  const hours = Math.floor(minutes / 60);
  return hours > 0 ? `${hours}h ${minutes % 60}m` : `${minutes}m`;
}

/** One status line for a kind. */
export function describeKind(kind: TokenKind, exp: number | undefined, now = Date.now()): string {
  if (exp === undefined) return `${kind}: missing`;
  const remaining = exp * 1000 - now;
  if (remaining < 60_000) return `${kind}: expired at ${new Date(exp * 1000).toLocaleString()}`;
  return `${kind}: ok, expires in ${formatRemaining(remaining)}`;
}

let commandOutput: Promise<Partial<Record<TokenKind, string>>> | undefined;

/** Clears the cached TEAMS_TOKEN_COMMAND output. Tests use it between cases. */
export function resetTokenCache(): void {
  commandOutput = undefined;
}

async function fromCommand(): Promise<Partial<Record<TokenKind, string>>> {
  const command = process.env.TEAMS_TOKEN_COMMAND?.trim();
  if (!command) return {};
  commandOutput ??= (async () => {
    const proc = Bun.spawn(command.split(/\s+/), { stdout: "pipe", stderr: "ignore", env: { ...process.env } });
    const output = await new Response(proc.stdout).text();
    const code = await proc.exited;
    if (code !== 0) throw new Error(`TEAMS_TOKEN_COMMAND exited with status ${code}.`);
    try {
      const parsed = JSON.parse(output);
      if (!parsed || typeof parsed !== "object") throw new Error("not an object");
      return parsed as Partial<Record<TokenKind, string>>;
    } catch {
      throw new Error("TEAMS_TOKEN_COMMAND must print JSON with any of the keys chat, csa, spaces.");
    }
  })();
  return commandOutput;
}

function expOf(claims: Claims | undefined): number {
  return typeof claims?.exp === "number" ? claims.exp : 0;
}

/** Finds the token for a kind. Environment wins, then the command, then the stored profile. */
export async function resolveToken(kind: TokenKind): Promise<ResolvedToken | undefined> {
  const fromEnv = process.env[ENV_NAME[kind]];
  if (fromEnv) {
    const claims = decodeJwt(fromEnv);
    if (!claims) throw new Error(`${ENV_NAME[kind]} is not a JWT.`);
    checkAudience(kind, claims);
    return { kind, token: fromEnv, exp: expOf(claims), source: "environment", claims };
  }
  const fromCommandOutput = (await fromCommand())[kind];
  if (fromCommandOutput) {
    const claims = decodeJwt(fromCommandOutput);
    checkAudience(kind, claims);
    return { kind, token: fromCommandOutput, exp: expOf(claims), source: "command", claims };
  }
  const stored = (await loadProfile()).profile?.tokens[kind];
  if (stored) return { kind, token: stored.token, exp: stored.exp, source: "store", claims: decodeJwt(stored.token) };
  return undefined;
}

export async function hasToken(kind: TokenKind): Promise<boolean> {
  return (await resolveToken(kind)) !== undefined || !!(await loadProfile()).profile?.refresh;
}

/** Returns a usable token, or throws. A token that expires within 60 seconds counts as expired. */
export async function requireToken(kind: TokenKind): Promise<ResolvedToken> {
  let resolved = await resolveToken(kind);
  if ((!resolved || resolved.exp * 1000 - Date.now() < 60_000) &&
      (!resolved || resolved.source === "store") && (await loadProfile()).profile?.refresh) {
    // The sharepoint scope names the personal site host, so find it before the renewal takes the lock.
    if (kind === "sharepoint") await (await import("./sharepoint.ts")).sharepointHost();
    const { refreshToken } = await import("./refresh.ts");
    await refreshToken(kind);
    resolved = await resolveToken(kind);
  }
  if (!resolved) {
    // First run on this machine: a person at a terminal may sign in now. Then the same call finds the new tokens.
    if (await (await import("./first-run.ts")).offerSignIn()) return requireToken(kind);
    throw new Error(`No ${kind} token. Run: teams-cli auth login`);
  }
  if (resolved.exp * 1000 < Date.now()) throw new Error(expiredMessage(kind, resolved.exp));
  if (resolved.exp * 1000 - Date.now() < 60_000) {
    throw new Error(`The ${kind} token expires in under a minute. Supply a new one: teams-cli auth`);
  }
  return resolved;
}

export type Identity = {
  oid: string;
  mri: string;
  tenantId: string;
  upn: string;
  name: string;
  region: string;
  profile: string;
};

/** Returns who the tokens belong to. Stored profile data wins, then the token claims. */
export async function identity(): Promise<Identity> {
  const { name, profile } = await loadProfile();
  let oid = profile?.userId;
  let tid = profile?.tenantId;
  let upn = profile?.upn ?? "";
  let displayName = profile?.name ?? "";
  if (!oid || !tid) {
    for (const kind of KINDS) {
      const claims = (await resolveToken(kind))?.claims;
      if (!claims?.oid || !claims.tid) continue;
      oid ??= claims.oid;
      tid ??= claims.tid;
      upn ||= claims.upn ?? "";
      displayName ||= claims.name ?? "";
      break;
    }
  }
  if (!oid || !tid) {
    // Nearly every command asks for the identity first, so this is where a first run is caught.
    if (await (await import("./first-run.ts")).offerSignIn()) return identity();
    throw new Error("Teams identity is unknown. Run: teams-cli auth login");
  }
  return {
    oid,
    mri: `8:orgid:${oid}`,
    tenantId: tid,
    upn,
    name: displayName,
    region: process.env.TEAMS_REGION || profile?.region || "amer",
    profile: name,
  };
}
