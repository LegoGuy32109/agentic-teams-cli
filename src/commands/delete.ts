import { resolveConversation } from "../conversations.ts";
import { htmlToText } from "../message.ts";
import { READ_ONLY_STREAMS, encodeConversation, isChannelId, threadRootId } from "../refs.ts";
import { paceWrite } from "../pace.ts";
import { identity } from "../tokens.ts";
import { request } from "../transport.ts";

/** Said in the preview and in the help: the files of a deleted message cannot be found again from the message. */
export const FILES_WARNING = "Deleting a message leaves its files on SharePoint. Run `files delete <ref>` first.";

export type DeleteOptions = { json: boolean; allowWrite: boolean; noWait?: boolean; threadTs?: string };

export type DeleteTarget = {
  id: string;
  text: string;
  sender: string;
  time: string;
  alreadyDeleted: boolean;
  error?: string;
};

/** Reads each message reference. A ref is `<conversationId>#<messageId>`. A bare id has 10 or more digits. */
export function parseDeleteRefs(conversationId: string, refs: string[]): string[] {
  const ids: string[] = [];
  for (const ref of refs) {
    const index = ref.lastIndexOf("#");
    if (index < 0) {
      if (!/^\d{10,}$/.test(ref)) throw new Error(`"${ref}" is not a message id. Use <conversationId>#<messageId>, or a message id of 10 or more digits.`);
      ids.push(ref);
      continue;
    }
    const refConversation = ref.slice(0, index);
    const messageId = ref.slice(index + 1);
    if (!/^\d{10,}$/.test(messageId)) throw new Error(`"${ref}" has no valid message id after the #. Use <conversationId>#<messageId>.`);
    if (refConversation !== conversationId) throw new Error(`"${ref}" names another conversation. Every message must belong to the conversation given.`);
    ids.push(messageId);
  }
  return [...new Set(ids)];
}

/**
 * Parses one ref for react, unreact, and edit. It also takes a thread reply ref, `<channel>;messageid=<root>#<reply>`:
 * the web client reacts to and edits a reply on the channel path, with the reply id (captured 2026-10-19).
 */
export function parseMessageRef(conversationId: string, ref: string): string {
  const thread = /^(.+);messageid=\d{10,}#(\d{10,})$/.exec(ref);
  if (thread && thread[1] === conversationId) return thread[2]!;
  return parseDeleteRefs(conversationId, [ref])[0]!;
}

function messagePath(conversationId: string, messageId: string): string {
  return `/users/ME/conversations/${encodeConversation(conversationId)}/messages/${messageId}`;
}

/** Reads one message for the preview and the delete. A read failure is kept on the target, not thrown. */
async function readTarget(conversationId: string, messageId: string): Promise<DeleteTarget> {
  const time = /^\d{10,}$/.test(messageId) ? new Date(Number(messageId)).toISOString() : "";
  try {
    const raw: any = await request("chat", "GET", messagePath(conversationId, messageId));
    const message = Array.isArray(raw?.messages) ? raw.messages[0] : raw;
    const content = typeof message?.content === "string" ? message.content : "";
    const deleted = !!message?.properties?.deletetime;
    return {
      id: messageId,
      text: htmlToText(content).replace(/\s+/g, " ").trim().slice(0, 40),
      sender: typeof message?.imdisplayname === "string" && message.imdisplayname ? message.imdisplayname : "unknown",
      time,
      alreadyDeleted: deleted && content === "",
    };
  } catch (error) {
    return { id: messageId, text: "", sender: "", time, alreadyDeleted: false, error: error instanceof Error ? error.message : String(error) };
  }
}

/**
 * Deletes messages in one conversation. Preview by default. Only `--allow-write` deletes,
 * one at a time, paced to the write limit. Deleting is not covered by the send-once ledger.
 */
export async function deleteMessages(input: string, refs: string[], options: DeleteOptions): Promise<void> {
  const target = await resolveConversation(input);
  if (READ_ONLY_STREAMS.has(target.id)) throw new Error("That stream is read-only.");
  if (options.threadTs !== undefined) options = { ...options, threadTs: threadRootId(options.threadTs, "--thread-ts") };
  if (options.threadTs && !isChannelId(target.id)) throw new Error("--thread-ts works only with channels.");
  const pathId = options.threadTs ? `${target.id};messageid=${options.threadTs}` : target.id;
  const ids = parseDeleteRefs(pathId, refs);
  const targets: DeleteTarget[] = [];
  for (const id of ids) targets.push(await readTarget(pathId, id));

  if (!options.allowWrite) {
    // A message that is already deleted is listed, but it is not counted.
    const toDelete = targets.filter(item => !item.alreadyDeleted).length;
    if (options.json) {
      console.log(JSON.stringify({
        ok: true,
        dry_run: true,
        effect: "message",
        method: "DELETE",
        path: messagePath(pathId, "{messageId}"),
        query: { behavior: "softDelete" },
        conversation: { id: target.id, label: target.label },
        count: toDelete,
        warning: FILES_WARNING,
        messages: targets.map(item => ({
          id: item.id,
          text: item.text,
          sender: item.sender,
          time: item.time,
          status: item.error ? "unreadable" : item.alreadyDeleted ? "already deleted" : "delete",
          ...(item.error ? { error: item.error } : {}),
        })),
      }, null, 2));
      return;
    }
    console.log(`Dry run (effect: message): would delete ${toDelete} message(s) in ${target.label}:`);
    for (const item of targets) {
      const status = item.error ? `unreadable: ${item.error}` : item.alreadyDeleted ? "already deleted" : `"${item.text}" by ${item.sender} at ${item.time}`;
      console.log(`  ${item.id}  ${status}`);
    }
    console.log(FILES_WARNING);
    return;
  }

  const who = await identity();
  const deleted: string[] = [];
  const skipped: string[] = [];
  const failed: Array<{ id: string; error: string }> = [];
  for (const item of targets) {
    if (item.error) {
      failed.push({ id: item.id, error: item.error });
      if (!options.json) console.log(`failed ${item.id}: ${item.error}`);
      continue;
    }
    if (item.alreadyDeleted) {
      skipped.push(item.id);
      if (!options.json) console.log(`skipped ${item.id}: already deleted`);
      continue;
    }
    try {
      await paceWrite(who.tenantId, { noWait: options.noWait, send: false });
      await request("chat", "DELETE", messagePath(pathId, item.id), { query: { behavior: "softDelete" } });
      deleted.push(item.id);
      if (!options.json) console.log(`deleted ${item.id}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      failed.push({ id: item.id, error: message });
      if (!options.json) console.log(`failed ${item.id}: ${message}`);
    }
  }

  if (options.json) {
    console.log(JSON.stringify({ ok: failed.length === 0, deleted, skipped, failed }, null, 2));
  } else {
    console.log(`Deleted ${deleted.length}, skipped ${skipped.length}, failed ${failed.length}.`);
  }
  if (failed.length) process.exitCode = 1;
}
