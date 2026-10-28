import { DEFAULT_CONTEXT, MAX_CONTEXT } from "../limits.ts";
import { loadHistory } from "./history.ts";
import { parseMessageRef, encodeConversation, isChannelId } from "../refs.ts";
import { normalizeMessage, sortByTime, mentionedMris } from "../message.ts";
import { resolveNames } from "../users.ts";
import { cardMarker } from "../card.ts";
import { identity } from "../tokens.ts";
import { request } from "../transport.ts";

export async function context(ref: string, options: { json?: boolean; before?: number; after?: number }) {
  const parsed = parseMessageRef(ref);
  if (!parsed) throw new Error("Give a message ref: <conversation-id>#<message-id>.");
  const before = options.before ?? DEFAULT_CONTEXT;
  const after = options.after ?? DEFAULT_CONTEXT;
  if (![before, after].every(n => Number.isInteger(n) && n >= 0 && n <= MAX_CONTEXT)) throw new Error(`Context --before and --after must be whole numbers from 0 to ${MAX_CONTEXT}.`);
  const who = await identity();
  const raw = await request("chat", "GET", `/users/ME/conversations/${encodeConversation(parsed.conversationId)}/messages/${parsed.messageId}`);
  const names = await resolveNames(who.tenantId, mentionedMris(raw));
  const anchor = normalizeMessage(raw, { conversationId: parsed.conversationId, myMri: who.mri, nameFor: mri => names[mri], system: true });
  if (!anchor) throw new Error("The message could not be read.");
  const result = await loadHistory(parsed.conversationId, { json: true, content: true, system: true,
    around: { ts: parsed.messageId, before, after } });
  if (!result.messages.some(message => message.ts === anchor.ts)) result.messages = sortByTime([...result.messages, anchor]);
  let thread;
  if (isChannelId(parsed.conversationId)) thread = await loadHistory(parsed.conversationId, { json: true, content: true, all: true, thread: anchor.reply_to ?? anchor.ts });
  const output = { ...result, anchor, ...(thread ? { thread } : {}) };
  if (options.json) console.log(JSON.stringify(output, null, 2));
  else {
    console.log(`${result.channel} (${result.messages.length} messages)`);
    for (const message of result.messages) console.log(`${message.ts === anchor.ts ? ">" : " "} [${message.created}] ${message.user}: ${message.card ? `${cardMarker(message.card)} ` : ""}${message.content}`);
    if (result.capped) console.error("Context reached the history cap. Some surrounding messages may be missing.");
  }
}
