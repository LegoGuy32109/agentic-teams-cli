import { resolveConversation } from "../conversations.ts";
import { reactionsFrom } from "../message.ts";
import { paceWrite } from "../pace.ts";
import { resolveReactionKey } from "../reactions.ts";
import { READ_ONLY_STREAMS, encodeConversation, formatMessageRef } from "../refs.ts";
import { identity } from "../tokens.ts";
import { request, TeamsError } from "../transport.ts";
import { parseMessageRef } from "./delete.ts";

export type ReactOptions = { json: boolean; allowWrite: boolean; noWait?: boolean; verify?: boolean };

/** Reads one message and says whether my user id is under the key. */
async function readState(path: string, key: string, mri: string): Promise<{ deleted: boolean; mine: boolean }> {
  const raw: any = await request("chat", "GET", path);
  const message = Array.isArray(raw?.messages) ? raw.messages[0] : raw;
  const mine = reactionsFrom(message?.properties?.emotions).some(item => item.key === key && item.mris.includes(mri));
  return { deleted: !!message?.properties?.deletetime, mine };
}

/** Writes one reaction request. It uses the write pacer, and it does not retry an uncertain outcome. */
async function sendReaction(writePath: string, method: "PUT" | "DELETE", body: unknown, remove: boolean, input: string, tenantId: string, options: { noWait?: boolean }): Promise<void> {
  await paceWrite(tenantId, { noWait: options.noWait, send: false });
  try {
    // The web client sends the DELETE body as text/plain. Keep that shape.
    await request("chat", method, writePath, { query: { name: "emotions" }, body, ...(remove ? { contentType: "text/plain;charset=UTF-8" } : {}) });
  } catch (error) {
    if (error instanceof TeamsError && error.uncertain) {
      throw new Error(`Outcome unknown: Teams did not confirm this reaction. Check with: teams-cli history ${input} --top=5`);
    }
    throw error;
  }
}

/**
 * Adds your reaction and prints nothing. It is the write path of `react`, for a command that reacts for you, such as
 * `card send` on a vote card. It makes no request when you already reacted. It returns the key and whether it wrote.
 */
export async function addReaction(input: string, ref: string, emoji: string, options: { noWait?: boolean }): Promise<{ key: string; changed: boolean }> {
  const key = resolveReactionKey(emoji);
  const who = await identity();
  const target = await resolveConversation(input);
  if (READ_ONLY_STREAMS.has(target.id)) throw new Error("That stream is read-only.");
  const path = `/users/ME/conversations/${encodeConversation(target.id)}/messages/${parseMessageRef(target.id, ref)}`;
  const state = await readState(path, key, who.mri);
  if (state.deleted) throw new Error("That message was deleted. It cannot take a reaction.");
  if (state.mine) return { key, changed: false };
  await sendReaction(`${path}/properties`, "PUT", { emotions: { key, value: Date.now() } }, false, input, who.tenantId, options);
  return { key, changed: true };
}

/**
 * Adds or removes one of your reactions. Preview by default. A reaction is not a send: it has no ledger record.
 * It uses the write pacer, as `edit` does. `react` on a key you already used, and `unreact` on a key you did not use,
 * make no write request.
 */
export async function reactToMessage(input: string, ref: string, emoji: string, remove: boolean, options: ReactOptions): Promise<void> {
  const key = resolveReactionKey(emoji);
  const who = await identity();
  const target = await resolveConversation(input);
  if (READ_ONLY_STREAMS.has(target.id)) throw new Error("That stream is read-only.");
  const messageId = parseMessageRef(target.id, ref);
  const path = `/users/ME/conversations/${encodeConversation(target.id)}/messages/${messageId}`;
  const writePath = `${path}/properties`;
  const messageRef = formatMessageRef(target.id, messageId);
  const action = remove ? "unreact" : "react";

  const state = await readState(path, key, who.mri);
  if (state.deleted) throw new Error("That message was deleted. It cannot take a reaction.");
  const unchanged = remove ? !state.mine : state.mine;
  const base = { conversation: { id: target.id, label: target.label }, message_id: messageId, key };
  if (unchanged) {
    const note = remove ? `You have no "${key}" reaction on ${messageRef}. Nothing to remove.` : `You already reacted with "${key}" on ${messageRef}. Nothing to add.`;
    if (options.json) console.log(JSON.stringify({ ok: true, action, changed: false, ...base, note }, null, 2));
    else console.log(note);
    return;
  }

  const body = { emotions: remove ? { key } : { key, value: Date.now() } };
  const method = remove ? "DELETE" : "PUT";
  if (!options.allowWrite) {
    if (options.json) {
      console.log(JSON.stringify({ ok: true, dry_run: true, effect: "reaction", action, method, path: `/api/chatsvc/${who.region}/v1${writePath}?name=emotions`, ...base, body }, null, 2));
    } else {
      console.log(`Dry run (effect: reaction): would ${remove ? "remove" : "add"} "${key}" ${remove ? "from" : "on"} ${messageRef} in ${target.label}.`);
    }
    return;
  }

  await sendReaction(writePath, method, body, remove, input, who.tenantId, options);

  let verification: { ok: boolean } | undefined;
  if (options.verify) {
    const after = await readState(path, key, who.mri);
    verification = { ok: remove ? !after.mine : after.mine };
    if (!verification.ok) process.exitCode = 1;
  }
  const verb = remove ? "Removed" : "Added";
  if (options.json) {
    console.log(JSON.stringify({ ok: true, action, changed: true, ...base, ...(verification ? { verification } : {}) }, null, 2));
    return;
  }
  console.log(`${verb} "${key}" ${remove ? "from" : "on"} ${messageRef}`);
  if (verification) console.log(verification.ok ? "Verification: ok" : `Verification failed: the read-back ${remove ? "still has" : "lacks"} your "${key}" reaction.`);
}
