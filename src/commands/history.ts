import { DEFAULT_TOP, MAX_ALL, MAX_TOP } from "../limits.ts";
import { resolveNames } from "../users.ts";
import { labelFor, refreshConversations, resolveConversation } from "../conversations.ts";
import { encodeConversation, isChannelId, threadRootId } from "../refs.ts";
import { mentionedMris, normalizeMessage, sortByTime, type NormalizedMessage } from "../message.ts";
import { cardMarker } from "../card.ts";
import { identity } from "../tokens.ts";
import { request, TeamsError } from "../transport.ts";


export type HistoryOptions = {
  json: boolean;
  content: boolean;
  top?: number;
  afterTs?: string;
  thread?: string;
  all?: boolean;
  system?: boolean;
  /** Only for `mentions`: keep mentions that are newer than the read position of their conversation. */
  unread?: boolean;
  around?: { ts: string; before: number; after: number };
};

/** Reads an --after-ts value. It takes an ISO datetime or milliseconds of 12 or more digits. */
export function parseAfterTs(raw: string): number {
  if (/^\d{12,}$/.test(raw)) return Number(raw);
  const parsed = Date.parse(raw);
  if (Number.isNaN(parsed)) throw new Error("--after-ts must be an ISO datetime or milliseconds since the epoch.");
  return parsed;
}

function validate(options: HistoryOptions): void {
  if (options.top !== undefined && options.afterTs !== undefined) throw new Error("Use either --top or --after-ts, not both.");
  if (options.all && (options.top !== undefined || options.afterTs !== undefined)) throw new Error("Use --all alone, not with --top or --after-ts.");
  if (options.top !== undefined && (!Number.isInteger(options.top) || options.top < 1 || options.top > MAX_TOP)) {
    throw new Error(`--top must be a whole number from 1 to ${MAX_TOP}.`);
  }
  if (options.thread !== undefined) threadRootId(options.thread, "--thread");
}

/** Reads where a mention came from. The mentions stream puts it in `properties.activity`. */
function mentionSource(raw: any): { threadId: string; messageId: string } | undefined {
  let activity = raw?.properties?.activity;
  if (typeof activity === "string") { try { activity = JSON.parse(activity); } catch { return undefined; } }
  const threadId = activity?.sourceThreadId;
  const messageId = activity?.sourceMessageId;
  return typeof threadId === "string" && messageId !== undefined ? { threadId, messageId: String(messageId) } : undefined;
}

type ReadMode = { kind: "top"; top: number } | { kind: "after"; afterMs: number } | { kind: "all" } | { kind: "around"; ts: string; before: number; after: number };

/**
 * Reads pages backward until the mode is satisfied or the conversation starts. Each page follows the
 * `backwardLink` from the one before, one request at a time. `capped` means the read stopped at MAX_ALL.
 */
async function readHistory(pathId: string, mode: ReadMode, system: boolean, conversationId: string): Promise<{ raw: any[]; capped: boolean }> {
  const counting = { conversationId, myMri: "", nameFor: () => undefined, system };
  const raw: any[] = [];
  const seen = new Set<string>();
  const seenMessages = new Set<string>();
  let kept = 0;
  let oldest = Infinity;
  let before = 0;
  let page: any;
  try {
    page = await request("chat", "GET", `/users/ME/conversations/${encodeConversation(pathId)}/messages`, {
      query: { view: "msnp24Equivalent|supportsMessageProperties", pageSize: "200", startTime: "1" },
    });
  } catch (error) {
    if (error instanceof TeamsError && error.status === 403) {
      throw new Error("You are not a member of this conversation, or the chat token lacks access.");
    }
    throw error;
  }

  for (;;) {
    for (const item of (Array.isArray(page?.messages) ? page.messages : []) as unknown[]) {
      const id = (item as any)?.id;
      if (typeof id === "string" && seenMessages.has(id)) continue;
      if (typeof id === "string") seenMessages.add(id);
      raw.push(item);
      if (raw.length > MAX_ALL) { raw.pop(); return { raw, capped: true }; }
      if (normalizeMessage(item, counting) !== undefined) {
        kept++;
        if (mode.kind === "around" && Number((item as any)?.id) < Number(mode.ts)) before++;
      }
      const numeric = Number((item as any)?.id);
      if (Number.isFinite(numeric) && numeric < oldest) oldest = numeric;
    }
    const link = page?._metadata?.backwardLink;
    if (typeof link !== "string" || !link || seen.has(link)) return { raw, capped: false };
    if (raw.length >= MAX_ALL) return { raw, capped: true };
    if (mode.kind === "top" && kept >= mode.top) return { raw, capped: false };
    if (mode.kind === "after" && oldest <= mode.afterMs) return { raw, capped: false };
    if (mode.kind === "around" && oldest <= Number(mode.ts) && before >= mode.before) return { raw, capped: false };
    seen.add(link);
    page = await request("chat", "GET", "", { fullUrl: link });
  }
}

function reactionLine(message: NormalizedMessage): string | undefined {
  if (!message.reactions.length) return undefined;
  return `  reactions: ${message.reactions.map(item => `${item.emoji ?? item.key} ${item.count} (${item.users.join(", ")})`).join(", ")}`;
}

/** Reads messages from one conversation, oldest first. Paging follows the service's backward links. */
export async function loadHistory(input: string, options: HistoryOptions) {
  validate(options);
  const who = await identity();
  const conversation = await resolveConversation(input);
  if (options.thread !== undefined && !isChannelId(conversation.id)) throw new Error("--thread works only with channels.");
  if (options.unread && conversation.id !== "48:mentions") throw new Error("--unread works only with the mentions stream: teams-cli history mentions --unread.");
  const afterMs = options.afterTs === undefined ? undefined : parseAfterTs(options.afterTs);
  const top = options.top ?? DEFAULT_TOP;
  const mode: ReadMode = options.around ? { kind: "around", ...options.around } : options.all ? { kind: "all" }
    : afterMs !== undefined ? { kind: "after", afterMs }
      : { kind: "top", top };

  const pathId = options.thread ? `${conversation.id};messageid=${threadRootId(options.thread, "--thread")}` : conversation.id;
  const system = !!options.system;
  const { raw, capped } = await readHistory(pathId, mode, system, conversation.id);
  const names = await resolveNames(who.tenantId, raw.flatMap(item => mentionedMris(item)));
  let normalized = sortByTime(raw
    .map(item => normalizeMessage(item, {
      conversationId: pathId,
      myMri: who.mri,
      nameFor: mri => names[mri],
      system,
    }))
    .filter((message): message is NormalizedMessage => !!message));

  if (options.unread) {
    // The read position of each conversation is the arrival time of its last read message. A message id is its arrival time.
    const cache = await refreshConversations();
    const horizons = new Map<string, string | undefined>([
      ...cache.teams.flatMap(team => team.channels.map(channel => [channel.id, channel.consumptionHorizon] as const)),
      ...cache.chats.map(chat => [chat.id, chat.consumptionHorizon] as const),
    ]);
    const sources = new Map(raw.map(item => [String(item?.id), mentionSource(item)] as const));
    // A mention with no known source or no known read position cannot be shown to be read, so it stays.
    normalized = normalized.flatMap(message => {
      const source = sources.get(message.ts);
      const horizon = source ? horizons.get(source.threadId) : undefined;
      if (source && horizon !== undefined && Number(source.messageId) <= Number(horizon)) return [];
      return [source ? { ...message, source: {
        conversation_id: source.threadId, message_id: source.messageId, label: labelFor(cache, source.threadId) ?? source.threadId, horizon: horizon ?? null,
      } } : message];
    });
  }

  let messages: NormalizedMessage[];
  if (mode.kind === "after") messages = normalized.filter(message => Number(message.ts) > mode.afterMs);
  else if (mode.kind === "top") messages = normalized.slice(-top);
  else if (mode.kind === "around") messages = [
    ...(mode.before ? normalized.filter(message => Number(message.ts) < Number(mode.ts)).slice(-mode.before) : []),
    ...normalized.filter(message => message.ts === mode.ts),
    ...normalized.filter(message => Number(message.ts) > Number(mode.ts)).slice(0, mode.after),
  ];
  else messages = normalized;
  return { channel: conversation.label, channel_id: conversation.id, messages, capped };
}

export async function history(input: string, options: HistoryOptions): Promise<void> {
  const { channel, channel_id, messages, capped } = await loadHistory(input, options);
  if (capped) {
    console.error(`Stopped after ${MAX_ALL.toLocaleString("en-US")} messages. Older messages were not read. Use --after-ts to read a narrower range.`);
  }

  if (options.json) {
    console.log(JSON.stringify({ channel, channel_id, messages }, null, 2));
    return;
  }
  console.log(`${channel} (${messages.length} messages)`);
  for (const message of messages) {
    const edited = message.edited ? " (edited)" : "";
    const deleted = message.deleted ? " (deleted)" : "";
    console.log(`[${new Date(message.created).toLocaleString()}] ${message.user} (id: ${message.id})${edited}${deleted}`);
    const body = options.content ? message.content : message.text;
    // A card message has no text of its own. Without the marker its line would be empty.
    for (const line of body.split("\n")) if (line || !message.card) console.log(`  ${line}`);
    if (message.card) console.log(`  ${cardMarker(message.card)}`);
    if (message.source) console.log(`  in ${message.source.label} (${message.source.conversation_id}#${message.source.message_id})`);
    if (message.files.length) console.log(`  files: ${message.files.map(file => file.name).join(", ")} (files download ${message.id})`);
    const reactions = reactionLine(message);
    if (reactions) console.log(reactions);
  }
}
