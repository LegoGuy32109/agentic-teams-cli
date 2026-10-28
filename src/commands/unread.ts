import { refreshConversations } from "../conversations.ts";
import { knownDenial, type AccessDenial } from "../access.ts";
import { identity } from "../tokens.ts";
import { encodeConversation } from "../refs.ts";
import { mapConcurrent, request, TeamsError } from "../transport.ts";

export type UnreadConversation = { id: string; label: string; kind: string; consumptionHorizon?: string };

/** Finds explicitly unread conversations. Missing state never implies unread. Known access denials go in `unavailable`. */
export async function findUnread(refreshAccess = false): Promise<{ conversations: UnreadConversation[]; unavailable: Array<UnreadConversation & AccessDenial> }> {
  const who = await identity();
  const cache = await refreshConversations();
  const candidates = [
    ...cache.teams.flatMap(team => team.channels.filter(channel => channel.isRead === false && channel.isMember !== false && channel.isDeleted !== true)
      .map(channel => ({ id: channel.id, label: `${team.name}/${channel.name}`, kind: "channel", consumptionHorizon: channel.consumptionHorizon }))),
    ...cache.chats.filter(chat => chat.isRead === false)
      .map(chat => ({ id: chat.id, label: chat.title || chat.id, kind: "chat", consumptionHorizon: chat.consumptionHorizon })),
  ];
  if (refreshAccess) await mapConcurrent(candidates, 4, async conversation => {
    try {
      await request("chat", "GET", `/users/ME/conversations/${encodeConversation(conversation.id)}/messages`, {
        query: { view: "msnp24Equivalent|supportsMessageProperties", pageSize: "1", startTime: "1" },
      });
    } catch (error) {
      if (!(error instanceof TeamsError) || error.status !== 403) throw error;
    }
  });
  const conversations: UnreadConversation[] = [];
  const unavailable: Array<UnreadConversation & AccessDenial> = [];
  for (const conversation of candidates) {
    const denied = await knownDenial(who, conversation.id);
    if (denied) unavailable.push({ ...conversation, ...denied });
    else conversations.push(conversation);
  }
  return { conversations, unavailable };
}

/** Lists only explicitly unread conversations. Missing state never implies unread. */
export async function unread(json: boolean, refreshAccess = false): Promise<void> {
  const { conversations, unavailable } = await findUnread(refreshAccess);
  if (json) console.log(JSON.stringify({ conversations, unavailable }, null, 2));
  else {
    for (const conversation of conversations) console.log(`${conversation.label} (${conversation.id})`);
    for (const conversation of unavailable) console.log(`unavailable: ${conversation.label} (${conversation.id}). ${conversation.reason}`);
  }
}
