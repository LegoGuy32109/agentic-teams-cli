import { cardTitle, messageCards } from "./card.ts";
import { loadConversations, refreshConversations } from "./conversations.ts";
import { userIdFrom } from "./message.ts";
import { encodeConversation, formatMessageRef } from "./refs.ts";
import { hasToken, identity } from "./tokens.ts";
import { request } from "./transport.ts";

/** The bot that posts a draft is shown by this name in the chat service (`imdisplayname`). */
export const WORKFLOWS_BOT_NAME = "Workflows";
/** `--wait` reads the target this many times, this far apart. That is 20 seconds in all. */
export const WAIT_POLLS = 11;
export const WAIT_INTERVAL_MS = 2000;
/** A message that is older than the send by more than this is not the card that was sent. It also covers clock skew. */
export const WAIT_SKEW_MS = 60_000;
const VIEW = "msnp24Equivalent|supportsMessageProperties";

let sleep = (ms: number) => Bun.sleep(ms);

/** Replaces the wait between polls. Tests call it to skip real waits. */
export function setWaitSleepForTests(fn: (ms: number) => Promise<unknown>): void {
  sleep = fn as (ms: number) => Promise<void>;
}

async function recentMessages(path: string, size: number): Promise<any[]> {
  const page = await request("chat", "GET", `/users/ME/conversations/${encodeConversation(path)}/messages`, {
    query: { view: VIEW, pageSize: String(size), startTime: "1" },
  });
  return Array.isArray(page?.messages) ? page.messages : [];
}

/**
 * Finds the user's Workflows chat: the 1:1 chat with a `28:` bot member that Teams shows as "Workflows".
 *
 * What the code relies on, checked live on 2026-10-19:
 * - `/v3/teams/users/me` lists the chat with `isOneOnOne: true`, an empty `title`, and two members: the user (`8:orgid:<guid>`)
 *   and the bot (`28:<guid>`). The cache keeps only those fields.
 * - The bot's messages in that chat carry `imdisplayname: "Workflows"`.
 * So a candidate is a 1:1 chat with exactly one `28:` member. The name is read from its recent messages. A chat that the
 * list titles "Workflows" is taken without a read. A chat that has no message from the bot yet is not found.
 */
export async function findWorkflowsChat(): Promise<string | undefined> {
  const who = await identity();
  const candidatesOf = (chats: Awaited<ReturnType<typeof loadConversations>>["chats"]) =>
    chats.filter(chat => chat.isOneOnOne && chat.members.filter(mri => mri.startsWith("28:")).length === 1);
  let candidates = candidatesOf((await loadConversations(who.tenantId)).chats);
  if (!candidates.length && (await hasToken("csa"))) candidates = candidatesOf((await refreshConversations()).chats);
  const titled = candidates.find(chat => chat.title.trim().toLowerCase() === WORKFLOWS_BOT_NAME.toLowerCase());
  if (titled) return titled.id;
  for (const chat of candidates.slice(0, 5)) {
    const messages = await recentMessages(chat.id, 10);
    if (messages.some(item => typeof item?.imdisplayname === "string" && item.imdisplayname.trim().toLowerCase() === WORKFLOWS_BOT_NAME.toLowerCase())) return chat.id;
  }
  return undefined;
}

export type WaitTarget = {
  /** The conversation that holds the posted message: a channel, a Workflows chat. */
  conversationId: string;
  /** For a thread reply: the root message id. The reply is read from the thread. */
  thread?: string;
  /** True when the user posted the card. False for a draft, which the Workflows bot posts. */
  fromUser: boolean;
};

/** The ref of one posted card, for `card update`. */
export function postedRef(target: WaitTarget, messageId: string): string {
  return formatMessageRef(target.thread ? `${target.conversationId};messageid=${target.thread}` : target.conversationId, messageId);
}

/**
 * Finds the newest card message whose card title is `title`, from the right author, and not older than
 * `since` minus the skew. It reads the conversation only. It polls up to WAIT_POLLS times, WAIT_INTERVAL_MS apart.
 * Returns the ref, or undefined when the card did not show.
 */
export async function waitForPostedRef(target: WaitTarget, title: string | null, since: number): Promise<string | undefined> {
  const who = await identity();
  const path = target.thread ? `${target.conversationId};messageid=${target.thread}` : target.conversationId;
  for (let poll = 0; poll < WAIT_POLLS; poll++) {
    if (poll > 0) await sleep(WAIT_INTERVAL_MS);
    let best: number | undefined;
    for (const item of await recentMessages(path, 25)) {
      const id = Number(item?.id);
      if (!Number.isFinite(id) || id < since - WAIT_SKEW_MS) continue;
      if ((userIdFrom(item?.from) === who.mri) !== target.fromUser) continue;
      // Teams rewrites a stored card, so the title is the stable part to compare.
      if (!messageCards(item).some(content => cardTitle(content) === title)) continue;
      if (best === undefined || id > best) best = id;
    }
    if (best !== undefined) return postedRef(target, String(best));
  }
  return undefined;
}
