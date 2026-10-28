export const NOTES_ID = "48:notes";

/** A conversation id: a chat, a channel, a meeting chat, or a personal stream (`48:*`). */
export function isConversationId(value: string): boolean {
  return /^48:\S+$/.test(value) || /^19:\S+@\S+$/.test(value);
}

/** Streams the service writes to. A send to one of these fails before any request. */
export const READ_ONLY_STREAMS = new Set(["48:mentions", "48:notifications"]);

/** The short names that stand for personal streams. `me` is any case. */
export const STREAM_ALIASES: Record<string, string> = {
  mentions: "48:mentions",
  notifications: "48:notifications",
};

export const STREAM_LABELS: Record<string, string> = {
  "48:notes": "Notes to self",
  "48:mentions": "Mentions",
  "48:notifications": "Notifications",
};

export function isChannelId(id: string): boolean {
  return /@thread\.(tacv2|skype)$/.test(id);
}

export function isEmail(value: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

/** Joins a conversation id and a message id as `<conversationId>#<messageId>`. */
export function formatMessageRef(conversationId: string, messageId: string): string {
  return `${conversationId}#${messageId}`;
}

/** Splits a message reference. A message id has 10 or more digits (milliseconds since the epoch). */
export function parseMessageRef(ref: string): { conversationId: string; messageId: string } | undefined {
  const index = ref.lastIndexOf("#");
  if (index <= 0) return undefined;
  const conversationId = ref.slice(0, index);
  const messageId = ref.slice(index + 1);
  if (!isConversationId(conversationId) || !/^\d{10,}$/.test(messageId)) return undefined;
  return { conversationId, messageId };
}

/** A thread root given as its message id or as a full message ref. Returns the message id. */
export function threadRootId(value: string, flag: string): string {
  if (/^\d{10,}$/.test(value)) return value;
  const ref = parseMessageRef(value);
  if (ref) return ref.messageId;
  throw new Error(`${flag} needs the root message id (10 or more digits) or its ref <conversation-id>#<message-id>.`);
}

/** Encodes a conversation id for a URL path. `:` becomes `%3A` and `@` becomes `%40`. */
export function encodeConversation(id: string): string {
  return encodeURIComponent(id);
}
