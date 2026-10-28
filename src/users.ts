import { join } from "node:path";
import { cacheDir, writeFileAtomic } from "./config.ts";
import { hasToken, identity } from "./tokens.ts";
import { request } from "./transport.ts";
import { readFile, rm } from "node:fs/promises";

export type User = {
  mri: string;
  displayName: string;
  realName?: string;
  mail?: string;
  upn?: string;
  fetchedAt: string;
  /** Set only on `@{everyone}` and `@{channel}`. A person has none. */
  mentionType?: "everyone" | "channel";
};

/** Version 2 stores mail from `email` too. Version 1 caches are ignored and refetched. */
type UserCache = { version: 2; updatedAt: string; users: Record<string, User> };

export const PLACEHOLDER_START = "";
export const PLACEHOLDER_END = "";

function usersPath(tenantId: string): string {
  return join(cacheDir(), tenantId, "users.json");
}

async function loadUsers(tenantId: string): Promise<UserCache> {
  try {
    const parsed = JSON.parse(await readFile(usersPath(tenantId), "utf8"));
    if (parsed?.version === 2 && parsed?.users && typeof parsed.users === "object") return parsed;
  } catch { /* Cache miss. */ }
  return { version: 2, updatedAt: "", users: {} };
}

async function saveUsers(tenantId: string, cache: UserCache): Promise<void> {
  await writeFileAtomic(usersPath(tenantId), `${JSON.stringify(cache, null, 2)}\n`, 0o600);
}

/** Scores a user against a query. 0 is exact, 1 is prefix, 2 is substring, 3 is no match. */
export function scoreUser(user: User, query: string): number {
  const values = [user.displayName, user.realName, user.mail, user.upn]
    .filter((value): value is string => !!value)
    .map(value => value.toLowerCase());
  const needle = query.trim().toLowerCase();
  if (values.some(value => value === needle)) return 0;
  if (values.some(value => value.startsWith(needle))) return 1;
  if (values.some(value => value.includes(needle))) return 2;
  return 3;
}

function searchCached(cache: UserCache, query: string): User[] {
  return Object.values(cache.users)
    .map(user => ({ user, score: scoreUser(user, query) }))
    .filter(item => item.score < 3)
    .sort((a, b) => a.score - b.score || a.user.displayName.localeCompare(b.user.displayName))
    .map(item => item.user);
}

export function toUser(raw: any): User | undefined {
  if (typeof raw?.mri !== "string" || !raw.mri.startsWith("8:orgid:")) return undefined;
  const realName = [raw.givenName, raw.surname].filter((part): part is string => typeof part === "string" && !!part).join(" ");
  const displayName = typeof raw.displayName === "string" && raw.displayName ? raw.displayName : realName || raw.userPrincipalName || raw.mri;
  return {
    mri: raw.mri,
    displayName,
    ...(realName ? { realName } : {}),
    ...(typeof (raw.mail || raw.email) === "string" && (raw.mail || raw.email) ? { mail: raw.mail || raw.email } : {}),
    ...(typeof raw.userPrincipalName === "string" && raw.userPrincipalName ? { upn: raw.userPrincipalName } : {}),
    fetchedAt: new Date().toISOString(),
  };
}

/** Finds users. The cache answers first. A miss calls searchV2 and caches the result. */
export async function findUsers(query: string): Promise<User[]> {
  const who = await identity();
  const cache = await loadUsers(who.tenantId);
  const cached = searchCached(cache, query);
  if (cached.length) return cached;
  if (!(await hasToken("spaces"))) throw new Error("User lookup needs a spaces token. Run: teams-cli auth login");

  const result = await request("spaces", "POST", "/users/searchV2", {
    query: { includeDLs: "false", enableGuest: "false", includeBots: "true", skypeTeamsInfo: "true", source: "shareContact" },
    rawBody: query,
    contentType: "text/plain;charset=UTF-8",
  });
  const found = ((result?.value ?? []) as unknown[]).map(toUser).filter((user): user is User => !!user);
  if (found.length) {
    for (const user of found) cache.users[user.mri] = user;
    cache.updatedAt = new Date().toISOString();
    await saveUsers(who.tenantId, cache);
  }
  return found;
}

/** Picks one user for a query. An exact match on name, mail, or upn wins over partial matches. */
export function pickUser(matches: User[], query: string, noun = "Mention"): User {
  const needle = query.toLowerCase();
  const exact = matches.filter(user => [user.displayName, user.mail, user.upn].some(value => value?.toLowerCase() === needle));
  const candidates = exact.length ? exact : matches;
  if (candidates.length !== 1) {
    const labels = candidates.slice(0, 5).map(user => `${user.displayName || user.upn || user.mri} (${user.mri})`).join(", ");
    throw new Error(candidates.length
      ? `${noun} "${query}" is ambiguous: ${labels}.`
      : `Could not find a Teams user matching "${query}".`);
  }
  return candidates[0]!;
}

/** Finds each `@{Name}` in order. Throws for an empty, unclosed, or nested mention before any lookup. */
export function mentionTokens(text: string): Array<{ start: number; end: number; query: string }> {
  const found: Array<{ start: number; end: number; query: string }> = [];
  let index = text.indexOf("@{");
  while (index >= 0) {
    const open = index + 2;
    const close = text.indexOf("}", open);
    if (close < 0) throw new Error(`Mention is not closed: "@{${text.slice(open, open + 20)}". Write @{Name}.`);
    const inner = text.slice(open, close);
    if (inner.includes("{")) throw new Error("Mention names cannot contain braces.");
    if (!inner.trim()) throw new Error("Empty mention. Write @{Name}.");
    found.push({ start: index, end: close + 1, query: inner.trim() });
    index = text.indexOf("@{", close + 1);
  }
  return found;
}

/** What the conversation is, so `@{everyone}` and `@{channel}` can be checked. `channelName` is read only when `@{channel}` is used. */
export type MentionScope = {
  conversationId: string;
  kind: "notes" | "stream" | "channel" | "oneOnOne" | "group";
  channelName?: () => Promise<string>;
};

const KIND_WORDS: Record<MentionScope["kind"], string> = {
  notes: "Notes to self",
  stream: "a personal stream",
  channel: "a channel",
  oneOnOne: "a 1:1 chat",
  group: "a group chat",
};

/**
 * `@{everyone}` and `@{channel}`, whole token and any case, are not person lookups. `@{everyone}` is valid only in a group or
 * meeting chat, and `@{channel}` only in a channel. Any other token, such as `@{Everyone Smith}` or an email address, is
 * a person lookup. A person whose name is Everyone or Channel is reached by email: `@{jo@contoso.com}`.
 */
async function specialMention(word: "everyone" | "channel", scope: MentionScope | undefined): Promise<User> {
  if (!scope) throw new Error(`@{${word}} needs a conversation to check against.`);
  if (word === "everyone") {
    if (scope.kind !== "group") throw new Error(`@{everyone} works only in a group or meeting chat. This is ${KIND_WORDS[scope.kind]}.`);
    return { mri: scope.conversationId, displayName: "Everyone", fetchedAt: "", mentionType: "everyone" };
  }
  if (scope.kind !== "channel" || !scope.channelName) throw new Error(`@{channel} works only in a channel. This is ${KIND_WORDS[scope.kind]}.`);
  return { mri: scope.conversationId, displayName: await scope.channelName(), fetchedAt: "", mentionType: "channel" };
}

/** Replaces each `@{Name}` with a private placeholder and returns the users in order of appearance. */
export async function resolveMentionTokens(text: string, scope?: MentionScope): Promise<{ text: string; mentions: User[] }> {
  const tokens = mentionTokens(text);
  if (!tokens.length) return { text, mentions: [] };
  const mentions: User[] = [];
  let output = "";
  let last = 0;
  for (const token of tokens) {
    const word = token.query.toLowerCase();
    const user = word === "everyone" || word === "channel"
      ? await specialMention(word, scope)
      : pickUser(await findUsers(token.query), token.query);
    mentions.push(user);
    output += text.slice(last, token.start) + `${PLACEHOLDER_START}${mentions.length - 1}${PLACEHOLDER_END}`;
    last = token.end;
  }
  output += text.slice(last);
  return { text: output, mentions };
}

export async function readUserCache(tenantId: string): Promise<UserCache> {
  return loadUsers(tenantId);
}

/** Names for user ids. The cache answers first. Other ids go to fetchShortProfile in batches, then fall back to the id. */
export async function resolveNames(tenantId: string, mris: string[]): Promise<Record<string, string>> {
  const users = await resolveUsers(tenantId, mris);
  return Object.fromEntries(Object.entries(users).map(([mri, user]) => [mri, user.displayName]));
}

/** Full user records for user ids, with the same cache-then-lookup order as resolveNames. Ids that cannot be resolved are left out. */
export async function resolveUsers(tenantId: string, mris: string[]): Promise<Record<string, User>> {
  const cache = await loadUsers(tenantId);
  const users: Record<string, User> = {};
  const missing: string[] = [];
  for (const mri of new Set(mris)) {
    const hit = cache.users[mri];
    if (hit) users[mri] = hit;
    else if (mri.startsWith("8:orgid:")) missing.push(mri);
  }
  if (!missing.length || !(await hasToken("spaces"))) return users;

  let changed = false;
  for (let start = 0; start < missing.length; start += 100) {
    try {
      const json: any = await request("spaces", "POST", "/users/fetchShortProfile", {
        query: {
          isMailAddress: "false", enableGuest: "true", skypeTeamsInfo: "true", canBeSmtpAddress: "false",
          includeIBBarredUsers: "true", includeDisabledAccounts: "true",
        },
        body: missing.slice(start, start + 100),
      });
      const profiles: unknown[] = Array.isArray(json) ? json : (Array.isArray(json?.value) ? json.value : []);
      for (const raw of profiles) {
        const user = toUser(raw);
        if (!user) continue;
        cache.users[user.mri] = user;
        users[user.mri] = user;
        changed = true;
      }
    } catch {
      // A failed lookup leaves these ids as they are. The caller shows the id.
    }
  }
  if (changed) {
    cache.updatedAt = new Date().toISOString();
    await saveUsers(tenantId, cache);
  }
  return users;
}

/**
 * Searches the directory for `users find`. A query of two or more characters always asks the
 * directory and merges the answer with the cache, cache hits first. A one-character query uses
 * the cache only. `refresh` skips the cache.
 */
export async function searchPeople(input: string, options: { refresh?: boolean } = {}): Promise<User[]> {
  const query = input.trim().replace(/\s+/g, " ");
  if (!query) throw new Error("Give a name or email to search for.");
  const who = await identity();
  const cache = await loadUsers(who.tenantId);
  if (query.length < 2) {
    if (options.refresh) throw new Error("--refresh needs a query of 2 or more characters.");
    console.error("Searched the cache only. A one-character query is not sent to the directory.");
    return searchCached(cache, query);
  }
  const cached = options.refresh ? [] : searchCached(cache, query);
  if (!(await hasToken("spaces"))) {
    console.error("No spaces token. Showing cached users only. Run: teams-cli auth login");
    return cached;
  }
  const result = await request("spaces", "POST", "/users/searchV2", {
    query: { includeDLs: "false", enableGuest: "false", includeBots: "true", skypeTeamsInfo: "true", source: "shareContact" },
    rawBody: query,
    contentType: "text/plain;charset=UTF-8",
  });
  const found = ((result?.value ?? []) as unknown[]).map(toUser).filter((user): user is User => !!user);
  if (found.length) {
    for (const user of found) cache.users[user.mri] = user;
    cache.updatedAt = new Date().toISOString();
    await saveUsers(who.tenantId, cache);
  }
  const seen = new Set<string>();
  return [...cached, ...found].filter(user => {
    if (seen.has(user.mri)) return false;
    seen.add(user.mri);
    return true;
  });
}

/** Clears the user-name cache for a tenant and returns how many entries it held. Teams has no bulk list to rebuild it from. */
export async function clearUserCache(tenantId: string): Promise<number> {
  const count = Object.keys((await loadUsers(tenantId)).users).length;
  await rm(usersPath(tenantId), { force: true });
  return count;
}
