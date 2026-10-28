import { resolveConversation } from "../conversations.ts";
import { attachPreviews } from "../preview.ts";
import { buildSendBody, checkBodySize, newClientMessageId } from "../compose.ts";
import { htmlToText, sameVisibleText, userIdFrom } from "../message.ts";
import { permalinkFor, type PreviewResult, type ReceiptResult } from "../operations.ts";
import { paceWrite } from "../pace.ts";
import { READ_ONLY_STREAMS, encodeConversation, formatMessageRef } from "../refs.ts";
import { identity } from "../tokens.ts";
import { request, TeamsError } from "../transport.ts";
import { parseMessageRef } from "./delete.ts";
import { composeMessage, mentionScope, validateComposeArgs } from "./send.ts";

export type EditOptions = {
  json: boolean;
  format?: string;
  html?: string;
  /** `plain` keeps emoji characters as typed. The default writes the Teams emoji markup. */
  emoji?: "teams" | "plain";
  allowWrite: boolean;
  noWait?: boolean;
  /** No preview card on the first link. */
  noPreview?: boolean;
};

/**
 * Replaces the text of one of your own messages. Preview by default. The edit uses the same composing
 * rules as send, and it is not covered by the send-once ledger.
 */
export async function editMessage(input: string, ref: string, words: string[], options: EditOptions): Promise<void> {
  const text = validateComposeArgs(words, options, "Usage: teams-cli edit <conversation> <message-id-or-ref> <text...> [--format=rich] [--html X] [--allow-write]");
  const who = await identity();
  const target = await resolveConversation(input);
  if (READ_ONLY_STREAMS.has(target.id)) throw new Error("That stream is read-only.");
  const messageId = parseMessageRef(target.id, ref);
  const path = `/users/ME/conversations/${encodeConversation(target.id)}/messages/${messageId}`;
  const fullPath = `/api/chatsvc/${who.region}/v1${path}`;

  const original: any = await request("chat", "GET", path);
  const existing = Array.isArray(original?.messages) ? original.messages[0] : original;
  if (existing?.properties?.deletetime) throw new Error("That message was deleted. It cannot be edited.");
  if (userIdFromSender(existing) !== who.mri) throw new Error("You can only edit your own messages. This message was sent by someone else.");
  const root = typeof existing?.rootMessageId === "string" ? existing.rootMessageId : undefined;
  const replyRoot = root && root !== messageId ? root : undefined;

  const { content, visible, entries, links, warning } = await composeMessage(text, options, await mentionScope(target.id));
  if (warning) console.error(warning);
  if (!(await attachPreviews(links, who.region, !options.noPreview))) console.error("Link preview unavailable. The link is sent without a card.");
  const full = buildSendBody({
    conversationId: target.id,
    content,
    mentions: entries,
    links,
    mri: who.mri,
    name: who.name,
    now: new Date().toISOString(),
    clientMessageId: newClientMessageId(),
  });
  // The edit carries the same fields as a send, plus the id it edits. It has no client message id.
  const body = {
    content: full.content,
    messagetype: full.messagetype,
    contenttype: full.contenttype,
    imdisplayname: full.imdisplayname,
    properties: full.properties,
    skypeeditedid: messageId,
    // A thread reply names its thread, as the web client's reply edit did.
    ...(replyRoot ? { replyChainId: replyRoot, conversationLink: `conversation/${target.id};messageid=${replyRoot}` } : {}),
  };
  checkBodySize(body);

  if (!options.allowWrite) {
    const preview: Omit<PreviewResult, "method"> & { method: "PUT"; message_id: string } = {
      ok: true,
      dry_run: true,
      effect: "message",
      method: "PUT",
      path: fullPath,
      conversation: { id: target.id, label: target.label },
      message_id: messageId,
      body,
      ...(warning ? { warning } : {}),
    };
    if (options.json) console.log(JSON.stringify(preview, null, 2));
    else console.log(`Dry run (effect: message): would edit ${formatMessageRef(target.id, messageId)} in ${target.label} to: ${visible}`);
    return;
  }

  await paceWrite(who.tenantId, { noWait: options.noWait, send: false });
  try {
    await request("chat", "PUT", path, { body });
  } catch (error) {
    if (error instanceof TeamsError && error.uncertain) {
      throw new Error(`Outcome unknown: Teams did not confirm this edit. Check with: teams-cli history ${input} --top=5`);
    }
    throw error;
  }

  const expected = htmlToText(content);
  const sent: any = await request("chat", "GET", path);
  const message = Array.isArray(sent?.messages) ? sent.messages[0] : sent;
  const actual = htmlToText(typeof message?.content === "string" ? message.content : "");
  const ts = messageId;
  const receipt: ReceiptResult & { text: string } = {
    ok: true,
    effect: "message",
    conversation: target.id,
    id: formatMessageRef(target.id, ts),
    ts,
    created: /^\d{10,}$/.test(ts) ? new Date(Number(ts)).toISOString() : "",
    permalink: permalinkFor(target.id, ts),
    text: actual,
    verification: { ok: sameVisibleText(actual, expected), expected, actual },
  };
  if (!receipt.verification!.ok) process.exitCode = 1;
  if (options.json) {
    console.log(JSON.stringify(receipt, null, 2));
    return;
  }
  console.log(`Edited ${receipt.id}`);
  console.log(receipt.permalink);
  console.log(receipt.verification!.ok ? "Verification: ok" : `Verification failed. Expected: ${expected}. Actual: ${actual}`);
}

function userIdFromSender(message: any): string {
  return userIdFrom(message?.from);
}
