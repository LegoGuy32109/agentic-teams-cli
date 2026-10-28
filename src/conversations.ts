import { directChatId } from "./direct.ts";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { cacheDir, writeFileAtomic } from "./config.ts";
import { NOTES_ID, STREAM_ALIASES, STREAM_LABELS, isChannelId, isConversationId, isEmail } from "./refs.ts";
import { findUsers, pickUser } from "./users.ts";
import { hasToken, identity } from "./tokens.ts";
import { request } from "./transport.ts";

/** `filesPath` is the channel's files folder on SharePoint, such as `/sites/x/Shared Documents/Name`. Teams has none until the channel's Files tab is first opened. */
export type ChannelInfo = { id: string; name: string; isRead?: boolean; consumptionHorizon?: string; isMember?: boolean; isDeleted?: boolean; filesPath?: string; isShared?: boolean };
/** `siteUrl` is the address of the team's SharePoint site. */
export type TeamInfo = { id: string; name: string; channels: ChannelInfo[]; siteUrl?: string; groupId?: string };
export type ChatInfo = { id: string; title: string; type: string; isOneOnOne: boolean; members: string[]; source?: "chatsvc"; isRead?: boolean; consumptionHorizon?: string };
export type ConversationCache = { version: 1; updatedAt: string; teams: TeamInfo[]; chats: ChatInfo[] };
/** `pending` marks a 1:1 chat that does not exist yet. Its id is the one Teams will give it. */
export type ResolvedConversation = { id: string; label: string; pending?: { mri: string } };

const EMPTY: ConversationCache = { version: 1, updatedAt: "", teams: [], chats: [] };

function text(value: unknown): string {
  return typeof value === "string" ? value : "";
}

/** Aggregate horizons are objects. Keep the message arrival timestamp for unread history reads. */
function horizon(value: any): string | undefined {
  if (typeof value === "string") return value;
  const arrival = value?.originalArrivalTime;
  return typeof arrival === "number" && Number.isFinite(arrival) ? String(arrival) : typeof arrival === "string" ? arrival : undefined;
}

function cachePath(tenantId: string): string {
  return join(cacheDir(), tenantId, "conversations.json");
}

export async function loadConversations(tenantId: string): Promise<ConversationCache> {
  try {
    const parsed = JSON.parse(await readFile(cachePath(tenantId), "utf8"));
    if (parsed?.version === 1 && Array.isArray(parsed.teams) && Array.isArray(parsed.chats)) return parsed;
  } catch { /* Cache miss. */ }
  return { ...EMPTY };
}

/** Reads the csa team and chat list. Every field may be missing, so each one is checked. */
export function parseTeamsMe(json: any): { teams: TeamInfo[]; chats: ChatInfo[] } {
  const teams: TeamInfo[] = (Array.isArray(json?.teams) ? json.teams : [])
    .filter((team: any) => text(team?.id))
    .map((team: any) => ({
      id: text(team.id),
      name: text(team.displayName),
      ...(text(team.teamSiteInformation?.sharepointSiteUrl) ? { siteUrl: text(team.teamSiteInformation.sharepointSiteUrl) } : {}),
      ...(text(team.teamSiteInformation?.groupId) ? { groupId: text(team.teamSiteInformation.groupId) } : {}),
      channels: (Array.isArray(team.channels) ? team.channels : [])
        .filter((channel: any) => text(channel?.id))
        .map((channel: any) => ({ id: text(channel.id), name: text(channel.displayName),
          ...(typeof channel.isMessageRead === "boolean" ? { isRead: channel.isMessageRead }
            : typeof channel.isRead === "boolean" ? { isRead: channel.isRead } : {}),
          ...(typeof channel.isMember === "boolean" ? { isMember: channel.isMember } : {}),
          ...(typeof channel.isDeleted === "boolean" ? { isDeleted: channel.isDeleted } : {}),
          ...(text(channel.defaultFileSettings?.filesRelativePath) ? { filesPath: text(channel.defaultFileSettings.filesRelativePath) } : {}),
          ...(channel.isShared === true ? { isShared: true } : {}),
          ...(horizon(channel.consumptionHorizon) !== undefined ? { consumptionHorizon: horizon(channel.consumptionHorizon) } : {}),
        })),
    }));
  const chats: ChatInfo[] = (Array.isArray(json?.chats) ? json.chats : [])
    .filter((chat: any) => text(chat?.id))
    .map((chat: any) => ({
      id: text(chat.id),
      title: text(chat.title),
      type: text(chat.threadType),
      isOneOnOne: chat.isOneOnOne === true,
      members: (Array.isArray(chat.members) ? chat.members : []).map((member: any) => text(member?.mri)).filter(Boolean),
      ...(typeof chat.isRead === "boolean" ? { isRead: chat.isRead } : {}),
      ...(horizon(chat.consumptionHorizon) !== undefined ? { consumptionHorizon: horizon(chat.consumptionHorizon) } : {}),
    }));
  return { teams, chats };
}

export async function saveConversations(tenantId: string, cache: ConversationCache): Promise<void> {
  await writeFileAtomic(cachePath(tenantId), `${JSON.stringify(cache, null, 2)}\n`, 0o600);
}

/** Downloads the team and chat list with the csa token and caches it for this tenant. */
export async function refreshConversations(): Promise<ConversationCache> {
  const who = await identity();
  const json = await request("csa", "GET", "/v3/teams/users/me", {
    query: { isPrefetch: "false", enableMembershipSummary: "true" },
  });
  const cache: ConversationCache = { version: 1, updatedAt: new Date().toISOString(), ...parseTeamsMe(json) };
  await saveConversations(who.tenantId, cache);
  return cache;
}

function candidatesFor(cache: ConversationCache, input: string): ResolvedConversation[] {
  const lower = (value: string) => value.trim().toLowerCase();
  if (input.includes("/")) {
    const slash = input.indexOf("/");
    const team = lower(input.slice(0, slash));
    const channel = lower(input.slice(slash + 1));
    return cache.teams
      .filter(entry => lower(entry.name) === team)
      .flatMap(entry => entry.channels
        .filter(item => lower(item.name) === channel)
        .map(item => ({ id: item.id, label: `${entry.name}/${item.name}` })));
  }
  const name = lower(input.replace(/^#/, ""));
  const channels = cache.teams.flatMap(entry => entry.channels
    .filter(item => lower(item.name) === name)
    .map(item => ({ id: item.id, label: `${entry.name}/${item.name}` })));
  if (channels.length) return channels;
  return cache.chats
    .filter(chat => chat.title && lower(chat.title) === name)
    .map(chat => ({ id: chat.id, label: chat.title }));
}

/** Reads the chat service's conversation list with the chat token. It is a second source for the csa list. */
export async function listChatService(): Promise<any[]> {
  const json: any = await request("chat", "GET", "/users/ME/conversations", {
    query: { view: "msnp24Equivalent", pageSize: "200", startTime: "1" },
  });
  return Array.isArray(json?.conversations) ? json.conversations : [];
}

/** Thread types that the chat service adds. Teams and channels (`space`, `topic`) come from the csa list. */
const MERGED_THREAD_TYPES = new Set(["meeting", "chat"]);

/**
 * Returns the chat-service conversations that the csa list lacks. Only `meeting` and `chat` threads are
 * merged. Personal streams (`48:*`) are not returned here. A merged thread with no title, no members,
 * and no last message is skipped.
 */
export function mergeChatService(cache: ConversationCache, items: any[]): ChatInfo[] {
  const known = new Set<string>([...cache.chats.map(chat => chat.id), ...cache.teams.flatMap(team => team.channels.map(channel => channel.id))]);
  const extra: ChatInfo[] = [];
  for (const item of items) {
    const id = text(item?.id);
    if (!id || id.startsWith("48:") || known.has(id)) continue;
    const type = text(item?.threadProperties?.threadType);
    if (!MERGED_THREAD_TYPES.has(type)) continue;
    // The chat service does not send members here, so a thread with no title and no last message is empty.
    const title = text(item?.threadProperties?.topic);
    if (!title && !item.lastMessage) continue;
    known.add(id);
    extra.push({ id, title, type, isOneOnOne: false, members: [], source: "chatsvc" });
  }
  return extra;
}

/** Looks up a label for a known conversation id. */
export function labelFor(cache: ConversationCache, id: string): string | undefined {
  if (STREAM_LABELS[id]) return STREAM_LABELS[id];
  for (const entry of cache.teams) {
    const channel = entry.channels.find(item => item.id === id);
    if (channel) return `${entry.name}/${channel.name}`;
  }
  const chat = cache.chats.find(item => item.id === id);
  return chat ? (chat.title || chat.id) : undefined;
}

function ambiguous(input: string, candidates: ResolvedConversation[]): Error {
  const list = candidates.map(item => `${item.label} (${item.id})`).join(", ");
  return new Error(`"${input}" matches more than one conversation: ${list}. Pass a conversation id.`);
}

/** Resolves a name against the cache. On a miss, refreshes once when a csa token exists. */
async function byName(input: string): Promise<ResolvedConversation> {
  const who = await identity();
  let cache = await loadConversations(who.tenantId);
  let found = candidatesFor(cache, input);
  if (!found.length) {
    if (!(await hasToken("csa"))) {
      throw new Error(`No conversation matches "${input}" in the cache. Name lookup needs a csa token. Run: teams-cli auth login, or pass a conversation id.`);
    }
    cache = await refreshConversations();
    found = candidatesFor(cache, input);
  }
  if (found.length === 1) return found[0]!;
  if (!found.length) throw new Error(`No conversation matches "${input}".`);
  throw ambiguous(input, found);
}

/**
 * Finds the cached 1:1 chat that has exactly me and one other user. It never creates a chat. With `allowNew`, a missing
 * chat comes back as pending, and the caller creates it when it sends.
 */
async function oneOnOneWith(input: string, allowNew: boolean): Promise<ResolvedConversation> {
  const who = await identity();
  const user = pickUser(await findUsers(input), input, "User");
  const pair = [who.mri, user.mri].sort().join("|");
  const matches = (cache: ConversationCache) => cache.chats.find(chat =>
    chat.isOneOnOne && [...new Set(chat.members)].sort().join("|") === pair);
  let cache = await loadConversations(who.tenantId);
  let chat = matches(cache);
  if (!chat && (await hasToken("csa"))) {
    cache = await refreshConversations();
    chat = matches(cache);
  }
  if (!chat && allowNew) return { id: directChatId(who.mri, user.mri), label: user.displayName, pending: { mri: user.mri } };
  if (!chat) throw new Error(`No 1:1 chat with ${input} exists in the cache. Send to ${input} with --allow-write to start one, or pass a chat id.`);
  return { id: chat.id, label: user.displayName };
}

/**
 * Resolves a conversation argument. Order: id, `me`, or a stream alias, `Team/Channel`, `@Name` or email, then
 * `#name` or a bare name (channels first, then chat titles).
 */
export async function resolveConversation(input: string, options: { allowNew?: boolean } = {}): Promise<ResolvedConversation> {
  const trimmed = input.trim();
  if (trimmed.toLowerCase() === "me") return { id: NOTES_ID, label: STREAM_LABELS[NOTES_ID]! };
  const alias = STREAM_ALIASES[trimmed.toLowerCase()];
  if (alias) return { id: alias, label: STREAM_LABELS[alias]! };
  if (isConversationId(trimmed)) {
    const who = await identity();
    const cache = await loadConversations(who.tenantId);
    return { id: trimmed, label: labelFor(cache, trimmed) ?? trimmed };
  }
  if (trimmed.startsWith("@") || isEmail(trimmed)) return oneOnOneWith(trimmed.replace(/^@/, ""), options.allowNew === true);
  return byName(trimmed);
}

export type ConversationKind = "notes" | "stream" | "channel" | "oneOnOne" | "group";

/**
 * What a conversation id is. A chat is one-to-one when the list says so, or when its id is a `@unq.gbl.spaces` pair id.
 * Any other `19:` chat, such as a group or meeting chat, is `group`.
 */
export function kindOf(cache: ConversationCache, id: string, pending = false): ConversationKind {
  if (id === NOTES_ID) return "notes";
  if (id.startsWith("48:")) return "stream";
  if (isChannelId(id)) return "channel";
  if (pending || id.endsWith("@unq.gbl.spaces") || cache.chats.some(chat => chat.id === id && chat.isOneOnOne)) return "oneOnOne";
  return "group";
}

/** The team and channel that own a channel id, from the cache. After a miss it refreshes the list once. `fresh` skips the cache. */
export async function findChannel(id: string, options: { fresh?: boolean } = {}): Promise<{ team: TeamInfo; channel: ChannelInfo }> {
  const who = await identity();
  const lookup = (cache: ConversationCache) => {
    for (const team of cache.teams) {
      const channel = team.channels.find(item => item.id === id);
      if (channel) return { team, channel };
    }
    return undefined;
  };
  let found = options.fresh ? undefined : lookup(await loadConversations(who.tenantId));
  if (!found && (await hasToken("csa"))) found = lookup(await refreshConversations());
  if (!found) throw new Error(`The channel ${id} is not in your team list, so its name and files folder are unknown. Run: teams-cli conversations`);
  return found;
}
