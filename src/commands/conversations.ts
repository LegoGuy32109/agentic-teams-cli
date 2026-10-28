import { STREAM_LABELS } from "../refs.ts";
import { listChatService, mergeChatService, refreshConversations, type ChatInfo, type ConversationCache } from "../conversations.ts";
import { resolveNames } from "../users.ts";
import { hasToken, identity } from "../tokens.ts";

/** The personal streams that every account has. Listed even when the service omits them. */
const STREAMS = ["48:notes", "48:mentions", "48:notifications"];
/** A chat with this many members or fewer lists the names. Larger chats keep the count only. */
const NAMED_MEMBER_LIMIT = 8;

type ChatRow = { id: string; title: string; type: string; source?: "chatsvc"; member_count: number | null; members?: string[] };

/**
 * Lists teams with channels, then chats, then the personal streams. The csa list is the main source.
 * The chat-service list adds conversations the csa list lacks. Without a csa token, only the chat list is used.
 */
export async function conversations(json: boolean): Promise<void> {
  const who = await identity();
  let cache: ConversationCache;
  if (await hasToken("csa")) {
    cache = await refreshConversations();
  } else {
    console.error("No csa token. Listing chat-service conversations only. Run: teams-cli auth login");
    cache = { version: 1, updatedAt: "", teams: [], chats: [] };
  }
  let items: any[] = [];
  if (await hasToken("chat")) items = await listChatService();
  else console.error("No chat token. Skipping the chat-service list. Run: teams-cli auth login");
  const chats: ChatInfo[] = [...cache.chats, ...mergeChatService(cache, items)];

  const named = chats.filter(chat => chat.members.length > 0 && chat.members.length <= NAMED_MEMBER_LIMIT).flatMap(chat => chat.members);
  const names = await resolveNames(who.tenantId, named);
  const rows: ChatRow[] = chats.map(chat => {
    const known = chat.source === "chatsvc" ? null : chat.members.length;
    const row: ChatRow = {
      id: chat.id,
      title: chat.title,
      type: chat.type,
      member_count: known,
      ...(chat.source ? { source: chat.source } : {}),
    };
    if (chat.members.length > 0 && chat.members.length <= NAMED_MEMBER_LIMIT) row.members = chat.members.map(mri => names[mri] ?? mri);
    return row;
  });
  const streams = STREAMS.map(id => ({ id, label: STREAM_LABELS[id]! }));

  if (json) {
    console.log(JSON.stringify({
      teams: cache.teams.map(team => ({ id: team.id, name: team.name, channels: team.channels })),
      chats: rows,
      streams,
    }, null, 2));
    return;
  }
  for (const team of cache.teams) {
    for (const channel of team.channels) console.log(`${team.name} / ${channel.name}  ${channel.id}`);
  }
  for (const row of rows) {
    const suffix = row.members ? ` (${row.members.join(", ")})` : row.member_count ? ` (${row.member_count} members)` : "";
    console.log(`chat: ${row.title || "(no title)"}${suffix}  ${row.id}`);
  }
  for (const stream of streams) console.log(`stream: ${stream.label}  ${stream.id}`);
}
