import { loadProfile, removeProfile, saveProfile, type Profile, type TokenKind } from "../config.ts";
import { classifyToken, describeKind, KIND_BY_AUD, KINDS, redact, resolveToken, UNLOCKS, type Claims } from "../tokens.ts";
import { request } from "../transport.ts";
import { browserStatusLine } from "../browser-auth.ts";

const BEARER = /authorization:\s*Bearer\s+([A-Za-z0-9._-]+)/gi;
const REGION = /\/api\/(?:chatsvc|csa|mt)\/([A-Za-z0-9-]+)\//;

export type ParsedPaste = {
  tokens: Partial<Record<TokenKind, { token: string; exp: number }>>;
  claims?: Claims;
  region?: string;
  ignored: number;
  /** The audience of each ignored token, or a note when the token cannot be read. */
  audiences: string[];
};

/** Reads every bearer token in pasted cURL text. Each kind keeps the token that expires last. */
export function parsePaste(text: string): ParsedPaste {
  const parsed: ParsedPaste = { tokens: {}, ignored: 0, audiences: [] };
  for (const match of text.matchAll(BEARER)) {
    const token = match[1]!;
    const { kind, claims } = classifyToken(token);
    if (!kind || !claims) {
      parsed.ignored++;
      const aud = claims?.aud;
      const audiences = Array.isArray(aud) ? aud : [aud];
      parsed.audiences.push(...(claims ? audiences.map(value => String(value ?? "(none)")) : ["(unreadable)"]));
      continue;
    }
    const exp = typeof claims.exp === "number" ? claims.exp : 0;
    const existing = parsed.tokens[kind];
    if (!existing || exp > existing.exp) parsed.tokens[kind] = { token, exp };
    parsed.claims ??= claims;
  }
  parsed.region = text.match(REGION)?.[1];
  return parsed;
}

async function readPaste(): Promise<string> {
  if (process.stdin.isTTY) process.stderr.write("Paste curl command (then press Ctrl+D):\n");
  const chunks: Buffer[] = [];
  for await (const chunk of Bun.stdin.stream()) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString("utf8");
}

/** Names the audiences that were found when none of them is a Teams service. */
export function unknownAudienceMessage(paste: ParsedPaste): string {
  const found = [...new Set(paste.audiences)];
  return `Found ${paste.ignored} token(s), but none for a Teams service. Audience(s): ${found.join(", ")}. Expected one of: ${Object.keys(KIND_BY_AUD).join(", ")}.`;
}

/** Saves the tokens from a paste. A second paste replaces only the kinds it contains. */
export async function auth(): Promise<void> {
  const paste = parsePaste(await readPaste());
  const kinds = KINDS.filter(kind => paste.tokens[kind]);
  if (!kinds.length && paste.ignored) throw new Error(unknownAudienceMessage(paste));
  if (!kinds.length) throw new Error("No authorization token found in the pasted text. Paste a cURL command with an authorization header.");

  const { name, profile } = await loadProfile();
  const next: Profile = {
    refresh: profile?.refresh,
    browser: profile?.browser,
    sharepointHost: profile?.sharepointHost,
    tokens: { ...(profile?.tokens ?? {}), ...paste.tokens },
    userId: paste.claims?.oid ?? profile?.userId ?? "",
    tenantId: paste.claims?.tid ?? profile?.tenantId ?? "",
    upn: paste.claims?.upn ?? profile?.upn ?? "",
    name: paste.claims?.name ?? profile?.name ?? "",
    region: paste.region ?? profile?.region ?? "amer",
    updatedAt: new Date().toISOString(),
  };
  if (!next.userId || !next.tenantId) throw new Error("The pasted tokens have no oid or tid claim. Paste a chat, csa, or spaces token.");
  await saveProfile(name, next);

  const now = Date.now();
  for (const kind of KINDS) console.log(describeKind(kind, next.tokens[kind]?.exp, now));
  for (const kind of KINDS) console.log(`${kind} unlocks: ${UNLOCKS[kind]}`);
  if (paste.ignored) console.log(`Ignored ${paste.ignored} token(s) with an unknown audience.`);
  console.log(`Profile: ${name}`);
}

export async function authStatus(): Promise<void> {
  const { name, profile } = await loadProfile();
  if (!profile) throw new Error("No stored profile. Run: teams-cli auth login");
  console.log(`profile: ${name}`);
  console.log(`upn: ${profile.upn || "(unknown)"}`);
  console.log(`tenant: ${profile.tenantId}`);
  console.log(`region: ${process.env.TEAMS_REGION || profile.region}`);
  const now = Date.now();
  for (const kind of KINDS) {
    const resolved = await resolveToken(kind);
    const line = describeKind(kind, resolved?.exp, now);
    console.log(resolved ? `${line} (source: ${resolved.source})` : line);
  }
  console.log(await browserStatusLine(name, profile));
  try {
    await request("chat", "GET", "/users/ME/properties");
    console.log("online: ok");
  } catch (error) {
    console.log(`online: ${redact(error instanceof Error ? error.message : String(error))}`);
    process.exitCode = 1;
  }
}

export async function logout(): Promise<void> {
  console.log(await removeProfile() ? "Profile removed." : "No matching stored profile.");
}
