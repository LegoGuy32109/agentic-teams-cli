import { emojiByName } from "./emoji.ts";
import { formatMessageRef } from "./refs.ts";
import { cardOfMessage, type CardInfo } from "./card.ts";

export type Reaction = {
  key: string;
  emoji?: string;
  count: number;
  users: string[];
  mine: boolean;
};

export type CallParticipant = { name: string; seconds?: number };

export type NormalizedMessage = {
  id: string;
  ts: string;
  created: string;
  type: string;
  reply_to?: string;
  edited?: string;
  pinned: boolean;
  deleted?: true;
  user: string;
  user_id: string;
  text: string;
  content: string;
  mentions: string[];
  mentioned: boolean;
  files: Array<{ name: string; url: string }>;
  /** Set when the message carries an Adaptive Card: its first TextBlock text and its element counts. */
  card?: CardInfo;
  reactions: Reaction[];
  /** Set by `history mentions --unread`: where the mention came from and the read position there. */
  source?: { conversation_id: string; message_id: string; label: string; horizon: string | null };
  call?: { participants: CallParticipant[] };
  recording?: { title: string; url: string; organizer?: string };
  transcript?: { callId: string; storageId: string };
};

export type NormalizeContext = {
  conversationId: string;
  myMri: string;
  nameFor: (mri: string) => string | undefined;
  /** Include system events, call records, and deleted messages. Without it they are dropped. */
  system?: boolean;
};

/** The reaction keys Teams uses for the six standard reactions. Other keys take their emoji from the catalog. */
export const REACTION_EMOJI: Record<string, string> = {
  like: "👍",
  heart: "❤️",
  laugh: "😂",
  surprised: "😮",
  sad: "🙁",
  angry: "😠",
};

const KEEP_TYPES = new Set(["Text", "RichText/Html"]);
const CALL_TYPE = "Event/Call";
const RECORDING_TYPE = "RichText/Media_CallRecording";
const TRANSCRIPT_TYPE = "RichText/Media_CallTranscript";
const ACTIVITY_PREFIX = "ThreadActivity/";
const MENTION = "http://schema.skype.com/Mention";
const REPLY = "http://schema.skype.com/Reply";
const FORWARD = "http://schema.skype.com/Forward";

function stripTags(value: string): string {
  return value.replace(/<[^>]+>/g, "");
}

export function decodeEntities(value: string): string {
  return value
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");
}

function escapeRe(value: string): string {
  return value.replace(/[/.]/g, "\\$&");
}

/** The text of the first element with this tag, or undefined. */
function firstTag(html: string, tag: string): string | undefined {
  const match = html.match(new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)<\\/${tag}>`, "i"));
  return match ? decodeEntities(stripTags(match[1]!)).replace(/\s+/g, " ").trim() : undefined;
}

/** Converts a Teams message body to plain text. Mentions become @Name, quotes become `> Author: text`, and forwards get a marker. */
export function htmlToText(html: string): string {
  // The service stores paragraphs with CRLF between them. That whitespace between tags is not content.
  let out = html.replace(/\r\n?/g, "\n").replace(/>\s*\n\s*</g, "><");
  // The code block editor marker is an empty paragraph that holds the block's id. It is structure, not text.
  out = out.replace(/<p\b[^>]*itemtype="http:\/\/schema\.skype\.com\/CodeBlockEditor"[^>]*>\s*(?:&nbsp;)?\s*<\/p>/gi, "");
  out = out.replace(new RegExp(`<span\\b[^>]*itemtype="${escapeRe(MENTION)}"[^>]*>([\\s\\S]*?)<\\/span>`, "gi"), (_, name: string) => `@${stripTags(name)}`);
  out = out.replace(new RegExp(`<blockquote\\b[^>]*itemtype="${escapeRe(REPLY)}"[^>]*>([\\s\\S]*?)<\\/blockquote>`, "gi"), (_, inner: string) => {
    const author = firstTag(inner, "strong");
    const quoted = htmlToText(inner.replace(/<strong\b[\s\S]*?<\/strong>/i, "")).replace(/\s*\n\s*/g, " ").trim();
    return `\n> ${author ? `${author}: ` : ""}${quoted}\n`;
  });
  out = out.replace(new RegExp(`<(\\w+)\\b[^>]*itemtype="${escapeRe(FORWARD)}"[^>]*>([\\s\\S]*?)<\\/\\1>`, "gi"), (_, _tag: string, inner: string) =>
    `\n[forwarded] ${htmlToText(inner)}\n`);
  out = out.replace(/<a\b[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, (_, href: string, label: string) => {
    const shown = stripTags(label).trim();
    return shown === href ? shown : `${shown} (${href})`;
  });
  out = out.replace(/<img\b[^>]*>/gi, tag => {
    // An emoji image reads as its character.
    if (/itemtype="http:\/\/schema\.skype\.com\/Emoji"/i.test(tag)) return /\balt="([^"]*)"/i.exec(tag)?.[1] ?? "";
    const id = /itemtype="http:\/\/schema\.skype\.com\/AMSImage"/i.test(tag) ? /\bitemid="([^"]+)"/i.exec(tag)?.[1] : undefined;
    return id ? `[image ${id}]` : "[image]";
  });
  out = out.replace(/<br\s*\/?>/gi, "\n").replace(/<\/(p|div)>/gi, "\n").replace(/<li\b[^>]*>/gi, "- ");
  out = decodeEntities(stripTags(out));
  return out.replace(/\n{3,}/g, "\n\n").trim();
}

/**
 * True when two message bodies show the same text. Teams stores the body in its own markup, so a
 * list or a line break can come back with other whitespace than the one that was sent. The check
 * ignores all whitespace, and it still fails when a word, a mention, or a link changes.
 */
export function sameVisibleText(a: string, b: string): boolean {
  return a.replace(/\s+/g, "") === b.replace(/\s+/g, "");
}

function parseProp(value: unknown): any {
  if (typeof value !== "string") return value;
  try {
    return JSON.parse(value);
  } catch {
    return undefined;
  }
}

/** Collects text, title, and value strings from a card tree. */
function cardTexts(value: unknown, out: string[]): string[] {
  if (Array.isArray(value)) {
    for (const item of value) cardTexts(item, out);
  } else if (value && typeof value === "object") {
    for (const [key, child] of Object.entries(value)) {
      if (["text", "title", "value"].includes(key) && typeof child === "string") out.push(child.trim());
      else cardTexts(child, out);
    }
  }
  return out;
}

/** Keeps parts that are not contained in another part. */
function distinctParts(parts: string[]): string[] {
  const cleaned = parts.map(part => part.trim()).filter(Boolean);
  return cleaned.filter((part, index) => !cleaned.some((other, otherIndex) =>
    otherIndex !== index && other.includes(part) && (other !== part || otherIndex < index)));
}

function fileEntries(value: unknown): Array<{ name: string; url: string }> {
  if (!Array.isArray(value)) return [];
  return value.flatMap(entry => {
    if (!entry || typeof entry !== "object") return [];
    const name = (entry as any).fileName || (entry as any).title || (entry as any).name;
    const url = (entry as any).objectUrl || (entry as any).url || "";
    return typeof name === "string" && name ? [{ name, url: typeof url === "string" ? url : "" }] : [];
  });
}

/** Returns the user ID (`8:orgid:<guid>`) from the `from` URL of a message. */
export function userIdFrom(from: unknown): string {
  if (typeof from !== "string") return "";
  let decoded = from;
  try { decoded = decodeURIComponent(from); } catch { /* Keep the raw value. */ }
  return decoded.match(/(8:[^/]+)$/)?.[1] ?? "";
}

/** Parses `properties.emotions`. Returns each reaction key with the user ids that used it. */
export function reactionsFrom(value: unknown): Array<{ key: string; mris: string[] }> {
  const parsed = parseProp(value);
  if (!Array.isArray(parsed)) return [];
  return parsed.flatMap(entry => {
    if (typeof entry?.key !== "string" || !Array.isArray(entry.users)) return [];
    const mris = entry.users.map((user: any) => user?.mri).filter((mri: unknown): mri is string => typeof mri === "string");
    return mris.length ? [{ key: entry.key, mris }] : [];
  });
}

/** Lists every user id a message names. The history command resolves these before it formats output. */
export function mentionedMris(raw: any): string[] {
  const props = raw?.properties && typeof raw.properties === "object" ? raw.properties : {};
  const type = typeof raw?.messagetype === "string" ? raw.messagetype : "";
  const found = reactionsFrom(props.emotions).flatMap(reaction => reaction.mris);
  if (type.startsWith(ACTIVITY_PREFIX)) {
    const json = parseActivityJson(raw.content);
    if (json) {
      if (typeof json.initiator === "string") found.push(json.initiator);
      if (typeof json.user === "string") found.push(json.user);
      for (const member of Array.isArray(json.members) ? json.members : []) {
        if (typeof member?.id === "string" && member.id.startsWith("8:orgid:")) found.push(member.id);
      }
    } else {
      const xml = String(raw.content ?? "");
      found.push(...xmlValues(xml, "initiator"), ...xmlValues(xml, "target"), ...xmlValues(xml, "id"));
    }
  }
  return [...new Set(found.filter(mri => mri.startsWith("8:")))];
}

function xmlValues(xml: string, tag: string): string[] {
  return [...xml.matchAll(new RegExp(`<${tag}>([^<]*)<\\/${tag}>`, "g"))].map(match => decodeEntities(match[1]!).trim()).filter(Boolean);
}

/** Parses a ThreadActivity body that is JSON. Some are XML, and those return undefined. */
function parseActivityJson(content: unknown): any {
  if (typeof content !== "string" || !content.startsWith("{")) return undefined;
  try {
    return JSON.parse(content);
  } catch {
    return undefined;
  }
}

/** Parses a transcript body. The service stores it as a JSON string, sometimes encoded twice. */
function parseTranscript(content: string): any {
  // The body may hold JSON with backslash-escaped quotes. Unescape them once and try again.
  for (const candidate of [content, content.replace(/\\"/g, "\"")]) {
    let value: any = candidate;
    for (let depth = 0; depth < 2 && typeof value === "string"; depth++) {
      try {
        value = JSON.parse(value);
      } catch {
        value = undefined;
        break;
      }
    }
    if (value && typeof value === "object") return value;
  }
  return undefined;
}

/** A readable one-line summary of a system event. Names come from the cache, else the summary stays generic. */
function activitySummary(kind: string, raw: any, nameFor: (mri: string) => string | undefined): string {
  const json = parseActivityJson(raw.content);
  const xml = typeof raw.content === "string" ? raw.content : "";
  const name = (mri: string | undefined) => (mri && nameFor(mri)) || undefined;
  if (json) {
    if (kind === "memberleft" || kind === "memberjoined") {
      const names = (Array.isArray(json.members) ? json.members : []).map((member: any) => member?.friendlyname || name(member?.id)).filter(Boolean);
      const verb = kind === "memberleft" ? "left" : "joined";
      return names.length ? `${names.join(", ")} ${verb}` : `member ${verb}`;
    }
    if (kind === "topictopicupdated" || kind === "spacetopicupdated") return `topic changed to "${json.newValue}"`;
    if (kind === "spacethreadchannelrenamed") return "channel renamed";
    if (kind === "spacethreadchanneladded") return "channel added";
    if (kind === "spacethreadchanneldeleted") return "channel deleted";
    if (kind === "topicdeleted") return "topic deleted";
    if (kind === "spacethreadcreated") return "thread created";
    return "system event";
  }
  if (kind === "addmember" || kind === "deletemember") {
    const verb = kind === "addmember" ? "added" : "removed";
    const initiator = name(xmlValues(xml, "initiator")[0]);
    const targets = [...new Set(xmlValues(xml, "target"))].map(mri => name(mri));
    if (initiator && targets.length && targets.every(Boolean)) return `${initiator} ${verb} ${targets.join(", ")}`;
    return `member ${verb}`;
  }
  if (kind === "roleupdate") return "role changed";
  if (kind === "topicupdate") return "topic changed";
  if (kind === "meetingpolicyupdated") return "meeting policy updated";
  if (kind === "modalityupdate") return "modality updated";
  return "system event";
}

/** Reads the call participants and their durations in seconds from the call XML. */
function callParticipants(xml: string, nameFor: (mri: string) => string | undefined): CallParticipant[] {
  return [...xml.matchAll(/<part\b[^>]*>([\s\S]*?)<\/part>/g)].map(match => {
    const part = match[1]!;
    const displayName = firstTag(part, "displayName");
    const identity = firstTag(part, "name") ?? "";
    const seconds = Number(firstTag(part, "duration"));
    const name = displayName || (identity.startsWith("8:") ? nameFor(identity) : undefined) || identity || "unknown";
    return Number.isFinite(seconds) && firstTag(part, "duration") !== undefined ? { name, seconds } : { name };
  });
}

/** Normalizes one chatsvc message. Returns undefined for a message the CLI skips. */
export function normalizeMessage(raw: any, context: NormalizeContext): NormalizedMessage | undefined {
  const messagetype = typeof raw?.messagetype === "string" ? raw.messagetype : "";
  const props = raw?.properties && typeof raw.properties === "object" ? raw.properties : {};
  const content = typeof raw?.content === "string" ? raw.content : "";
  const id = String(raw?.id ?? "");
  const deleted = !content && !!props.deletetime;
  const isMessage = KEEP_TYPES.has(messagetype);
  if (deleted) return context.system ? build(raw, context, { type: "deleted", text: "message deleted", content: "", deleted: true }) : undefined;
  if (isMessage) return build(raw, context, { type: "message" });
  if (!context.system) return undefined;

  if (messagetype === CALL_TYPE) {
    const participants = callParticipants(content, context.nameFor);
    const names = participants.map(item => item.name);
    return build(raw, context, {
      type: "call",
      text: names.length ? `Call with ${names.join(", ")}` : "Call",
      call: { participants },
    });
  }
  if (messagetype === RECORDING_TYPE) {
    const title = firstTag(content, "Title") ?? "";
    const url = decodeEntities(content.match(/<a\b[^>]*href="([^"]*)"/i)?.[1] ?? "");
    const organizerId = content.match(/<MeetingOrganizerId\b[^>]*value="([^"]*)"/i)?.[1];
    const organizer = organizerId ? (context.nameFor(`8:orgid:${organizerId}`) ?? organizerId) : undefined;
    return build(raw, context, {
      type: "recording",
      text: title ? `Recording: ${title}` : "Recording",
      recording: { title, url, ...(organizer ? { organizer } : {}) },
    });
  }
  if (messagetype === TRANSCRIPT_TYPE) {
    const transcript = parseTranscript(content);
    return build(raw, context, {
      type: "transcript",
      text: "Call transcript",
      transcript: { callId: String(transcript?.callId ?? ""), storageId: String(transcript?.storageId ?? "") },
    });
  }
  if (messagetype.startsWith(ACTIVITY_PREFIX)) {
    const kind = messagetype.slice(ACTIVITY_PREFIX.length).toLowerCase();
    return build(raw, context, { type: `system:${kind}`, text: activitySummary(kind, raw, context.nameFor) });
  }
  return undefined;
}

type Extra = {
  type: string;
  text?: string;
  content?: string;
  deleted?: true;
  call?: NormalizedMessage["call"];
  recording?: NormalizedMessage["recording"];
  transcript?: NormalizedMessage["transcript"];
};

/** Builds the shared fields. Text and file parts apply to messages. System events keep their raw content. */
function build(raw: any, context: NormalizeContext, extra: Extra): NormalizedMessage {
  const props = raw?.properties && typeof raw.properties === "object" ? raw.properties : {};
  const content = typeof raw?.content === "string" ? raw.content : "";
  const id = String(raw?.id ?? "");
  const created = /^\d{10,}$/.test(id) ? new Date(Number(id)).toISOString() : String(raw?.composetime ?? "");
  const edited = props.edittime ? new Date(Number(props.edittime)).toISOString() : undefined;
  const userId = userIdFrom(raw?.from);
  const user = (typeof raw?.imdisplayname === "string" && raw.imdisplayname) || (userId && context.nameFor(userId)) || userId || "system";
  const pinnedInfo = parseProp(props.pinned);
  const pinned = !!(pinnedInfo && typeof pinnedInfo === "object" && (pinnedInfo.creatorId || pinnedInfo.pinnedTime));
  const root = raw?.rootMessageId;
  const replyTo = root !== undefined && root !== null && String(root) !== "" && String(root) !== id ? String(root) : undefined;

  const text = extra.text ?? htmlToText(content);
  const mentionEntries: any[] = Array.isArray(parseProp(props.mentions)) ? parseProp(props.mentions) : [];
  const files = fileEntries(parseProp(props.files));
  const cards = cardTexts(parseProp(props.cards) ?? [], []);
  const card = cardOfMessage(raw);
  const body = extra.content ?? distinctParts([text, ...files.map(file => file.name), ...cards]).join("\n");
  const reactions: Reaction[] = reactionsFrom(props.emotions).map(reaction => {
    const emoji = REACTION_EMOJI[reaction.key] ?? emojiByName(reaction.key)?.unicode;
    return {
      key: reaction.key,
      ...(emoji ? { emoji } : {}),
      count: reaction.mris.length,
      users: reaction.mris.map(mri => context.nameFor(mri) ?? mri),
      mine: reaction.mris.includes(context.myMri),
    };
  });

  return {
    id: formatMessageRef(context.conversationId, id),
    ts: id,
    created,
    type: extra.type,
    ...(replyTo ? { reply_to: replyTo } : {}),
    ...(edited ? { edited } : {}),
    pinned,
    ...(extra.deleted ? { deleted: true as const } : {}),
    user,
    user_id: userId,
    text: extra.type === "message" ? text : (extra.text ?? text),
    content: body,
    mentions: mentionEntries.map(entry => String(entry?.displayName ?? "")).filter(Boolean),
    mentioned: mentionEntries.some(entry => entry?.mri === context.myMri),
    files,
    ...(card ? { card } : {}),
    reactions,
    ...(extra.call ? { call: extra.call } : {}),
    ...(extra.recording ? { recording: extra.recording } : {}),
    ...(extra.transcript ? { transcript: extra.transcript } : {}),
  };
}

/** Sorts normalized messages by numeric id, oldest first. */
export function sortByTime<T extends { ts: string }>(messages: T[]): T[] {
  return [...messages].sort((a, b) => Number(a.ts) - Number(b.ts));
}
