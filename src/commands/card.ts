import { cardOutline, cardTitle, cardPayload, checkPayloadSize, loadCardFile, loadTemplateCard, postToRouter, routerUrl, type CardMode, type LoadedCard } from "../card.ts";
import { findWorkflowsChat, waitForPostedRef, WAIT_INTERVAL_MS, WAIT_POLLS, type WaitTarget } from "../card-ref.ts";
import { resolveConversation } from "../conversations.ts";
import { ledgerKey, markSent, markUncertain, removeRecord, reserve } from "../ledger.ts";
import { paceWrite } from "../pace.ts";
import { isChannelId, threadRootId } from "../refs.ts";
import { resolveReactionKey } from "../reactions.ts";
import { identity } from "../tokens.ts";
import { addReaction } from "./react.ts";

/** A card file, or a template with the data file that fills it. */
export type CardSource = string | { template: string; data: string };
export type CardSendOptions = { json: boolean; allowWrite: boolean; draft?: boolean; thread?: string; noWait?: boolean; wait?: boolean; noReact?: boolean };
export type CardUpdateOptions = { json: boolean; allowWrite: boolean; noWait?: boolean };

export const ACCEPTED = "Accepted by Workflows. The flow posts within a few seconds. A channel that is not on the flow's allowlist is refused in the flow.";
const DRAFT_REMINDER = "Run with --draft --allow-write first. The draft shows the real rendering in your Workflows chat.";
const UPDATED = "Accepted by Workflows. The flow edits the card within a few seconds. A channel that is not on the flow's allowlist is refused in the flow.";

function describe(mode: CardMode, label: string, root?: string): string {
  if (mode === "draft") return `a draft in your private Workflows chat (the final target would be ${root ? `a reply in thread ${root} of ${label}` : `a new post in ${label}`})`;
  return mode === "reply" ? `a reply in thread ${root} of ${label}` : `a new post in ${label}`;
}

type SourcedCard = LoadedCard & { template?: string; votes: string[] };

async function loadSource(source: CardSource): Promise<SourcedCard> {
  if (typeof source === "string") return { ...(await loadCardFile(source)), votes: [] };
  return loadTemplateCard(source.template, source.data);
}

/** The reactions that a send of this card adds: each vote emoji and the Teams key it becomes. */
function voteReactions(card: SourcedCard): Array<{ emoji: string; key: string }> {
  return card.votes.map(emoji => ({ emoji, key: resolveReactionKey(emoji) }));
}

function imageLines(card: LoadedCard): string[] {
  return card.images.map(image => `image: ${image.path}, ${image.bytes} bytes, ${image.type}`);
}

function imageJson(card: LoadedCard) {
  return card.images.map(image => ({ path: image.path, type: image.type, bytes: image.bytes }));
}

type Submission = { key: string; mode: CardMode; context: { label: string; conversationId: string }; input: string; url: URL; payload: unknown; tenantId: string };

/** Reserves the ledger key, posts once, and settles the record. Returns normally only for a 202. */
async function submit(job: Submission): Promise<void> {
  await reserve(job.tenantId, job.key, { ...job.context, input: job.input });
  const result = await postToRouter(job.url, job.payload);
  if (result.status === 202) {
    await markSent(job.tenantId, job.key, `card:${job.mode}`, job.context);
    return;
  }
  if (result.uncertain) {
    await markUncertain(job.tenantId, job.key, job.context);
    const seen = result.status === undefined ? "a network error" : `status ${result.status}`;
    throw new Error(`Outcome unknown: the card router answered with ${seen}. Do not resend. Check with: teams-cli history ${job.input} --top=5`);
  }
  await removeRecord(job.tenantId, job.key);
  throw new Error(`The card router refused the card: status ${result.status}${result.body ? ` ${result.body}` : ""}`);
}

/** Previews a card by default. Only `--allow-write` posts it, through the Workflows card router, once per send-once key. */
export async function cardSend(input: string, source: CardSource, options: CardSendOptions): Promise<void> {
  const root = options.thread === undefined ? undefined : threadRootId(options.thread, "--thread");
  const loaded = await loadSource(source);
  const reactions = options.noReact ? [] : voteReactions(loaded);
  const target = await resolveConversation(input);
  const draft = options.draft === true;
  if (!isChannelId(target.id) && !draft) {
    throw new Error(`${target.label} is not a channel. Cards to chats go through --draft only for now.`);
  }
  if (root && !isChannelId(target.id)) throw new Error("--thread works only with channels.");

  const mode: CardMode = draft ? "draft" : root ? "reply" : "post";
  // A draft takes no channel. The final target is still in the preview, so Josh sees where the card would go.
  const payload = cardPayload(loaded.card, { mode, channel: target.id, replyTo: root });
  const payloadBytes = checkPayloadSize(payload);
  const where = describe(mode, target.label, root);

  if (!options.allowWrite) {
    if (options.json) {
      console.log(JSON.stringify({
        ok: true, dry_run: true, effect: "card", mode, conversation: { id: target.id, label: target.label },
        ...(root ? { thread: root } : {}), version: loaded.card.version, card_bytes: loaded.bytes, payload_bytes: payloadBytes, outline: cardOutline(loaded.card),
        ...(loaded.template ? { template: loaded.template } : {}), ...(reactions.length ? { reactions } : {}),
        ...(loaded.images.length ? { images: imageJson(loaded) } : {}), reminder: DRAFT_REMINDER,
      }, null, 2));
      return;
    }
    console.log(`Dry run (effect: card): would send ${where}.`);
    if (loaded.template) console.log(`  template: ${loaded.template}`);
    console.log(`  version: ${loaded.card.version}, ${loaded.bytes} bytes`);
    for (const line of cardOutline(loaded.card)) console.log(`  ${line}`);
    for (const line of imageLines(loaded)) console.log(`  ${line}`);
    if (reactions.length) console.log(`  reactions: after the post, would add ${reactions.map(item => `${item.emoji} (${item.key})`).join(", ")} as your votes. --no-react skips this.`);
    console.log(DRAFT_REMINDER);
    return;
  }

  // Read the URL before the ledger reserves anything, so a missing variable leaves no record.
  const url = await routerUrl();
  const who = await identity();
  if (mode !== "draft") await paceWrite(who.tenantId, { noWait: options.noWait });
  const key = ledgerKey({
    conversationId: mode === "draft" ? "card:draft" : `card:${mode}:${target.id}${root ? `;messageid=${root}` : ""}`,
    content: loaded.digest, mentionMris: [], contentType: "AdaptiveCard",
  });
  const context = { label: mode === "draft" ? "your Workflows chat" : target.label, conversationId: target.id };
  const startedAt = Date.now();
  await submit({ key, mode, context, input, url, payload, tenantId: who.tenantId });

  // The webhook answers only 202 and no id. With --wait, the posted message is found through the chat service.
  let ref: string | undefined;
  let waited = "";
  // Seeding the votes needs the ref of the new message, so a vote card always waits for it.
  const wait = options.wait || reactions.length > 0;
  let place: WaitTarget | undefined;
  if (wait) {
    place = await waitTarget(mode, target.id, root);
    if (!place) waited = "Could not find your Workflows chat, so the draft's ref is not known.";
    else {
      ref = await waitForPostedRef(place, cardTitle(loaded.card), startedAt);
      if (ref) await markSent(who.tenantId, key, ref, context);
      else waited = `The card did not show within ${(WAIT_POLLS - 1) * WAIT_INTERVAL_MS / 1000} s.`;
    }
  }
  const findRef = `teams-cli history ${mode === "draft" ? "<your Workflows chat id>" : input}${root ? ` --thread ${root}` : ""} --top=5 --json`;
  const seeded = reactions.length ? await seedVotes(reactions, ref, place?.conversationId ?? target.id, options.noWait) : [];
  const byHand = seeded.filter(item => item.status !== "added" && item.status !== "already");
  const handCommands = reactions.length && byHand.length
    ? byHand.map(item => `teams-cli react ${place?.conversationId ?? target.id} ${ref ?? "<ref>"} ${item.emoji} --allow-write`)
    : [];
  if (options.json) {
    console.log(JSON.stringify({
      ok: true, effect: "card", mode, conversation: target.id, ...(root ? { thread: root } : {}), status: 202, message: ACCEPTED,
      ...(loaded.template ? { template: loaded.template } : {}),
      ...(wait ? { ref: ref ?? null, ...(waited ? { wait_note: waited } : {}) } : { find_ref: findRef }),
      ...(reactions.length ? { reactions: seeded, ...(handCommands.length ? { react_by_hand: handCommands, ...(ref ? {} : { find_ref: findRef }) } : {}) } : {}),
    }, null, 2));
    if (seeded.some(item => item.status === "failed")) process.exitCode = 1;
    return;
  }
  console.log(ACCEPTED);
  if (ref) console.log(`Ref: ${ref}`);
  else if (wait) console.log(`${waited} Find the ref later with: ${findRef}`);
  else console.log(`To update this card later, find its ref with --wait, or with: ${findRef}`);
  for (const item of seeded) {
    if (item.status === "added") console.log(`Added your vote reaction ${item.emoji} (${item.key}).`);
    else if (item.status === "already") console.log(`You already had the reaction ${item.emoji} (${item.key}). Nothing to add.`);
    else if (item.status === "failed") console.log(`Could not add ${item.emoji} (${item.key}): ${item.error}`);
  }
  if (handCommands.length) {
    console.log(ref ? "Add the missing reactions by hand:" : "The vote reactions were not added, because the ref of the card was not found. After you find the ref, run:");
    for (const command of handCommands) console.log(`  ${command}`);
  }
  if (seeded.some(item => item.status === "failed")) process.exitCode = 1;
}

type Seed = { emoji: string; key: string; status: "added" | "already" | "failed" | "no-ref"; error?: string };

/**
 * Adds each vote emoji as your reaction on the new card. It goes through the `react` write path: the same pacing, no
 * second request for a reaction you already have. One failed reaction does not stop the rest. The send-once ledger
 * guards the post, so a repeated send does not reach this code.
 */
async function seedVotes(reactions: Array<{ emoji: string; key: string }>, ref: string | undefined, conversation: string, noWait: boolean | undefined): Promise<Seed[]> {
  if (!ref) return reactions.map(item => ({ ...item, status: "no-ref" as const }));
  const done: Seed[] = [];
  for (const item of reactions) {
    try {
      const result = await addReaction(conversation, ref, item.emoji, { noWait });
      done.push({ ...item, status: result.changed ? "added" : "already" });
    } catch (error) {
      done.push({ ...item, status: "failed", error: error instanceof Error ? error.message : String(error) });
    }
  }
  return done;
}

async function waitTarget(mode: CardMode, channelId: string, root: string | undefined): Promise<WaitTarget | undefined> {
  if (mode === "draft") {
    const chat = await findWorkflowsChat();
    return chat ? { conversationId: chat, fromUser: false } : undefined;
  }
  return { conversationId: channelId, ...(root ? { thread: root } : {}), fromUser: true };
}

type UpdateTarget = { mode: "update" | "update-draft"; channel?: string; messageId: string; conversationId: string; label: string; key: string };

/**
 * Reads the ref of a posted card. The forms:
 * - `<channel>#<id>`: a channel post.
 * - `<channel>;messageid=<root>#<reply>`: a thread reply. The flow takes `<root>/replies/<reply>`. The bare reply id gives 404.
 * - `draft#<id>`, or `<chat id>#<id>` for your Workflows chat: a draft.
 * `<channel>` may be a name that resolves to a channel.
 */
export async function parseUpdateTarget(input: string): Promise<UpdateTarget> {
  const forms = "Use <channel>#<id>, <channel>;messageid=<root>#<reply>, or draft#<id>.";
  const index = input.lastIndexOf("#");
  const messageId = index > 0 ? input.slice(index + 1) : "";
  if (!/^\d{10,}$/.test(messageId)) throw new Error(`"${input}" is not the ref of a posted card. ${forms}`);
  const left = input.slice(0, index);
  if (left.toLowerCase() === "draft") {
    const chat = await findWorkflowsChat();
    if (!chat) throw new Error("Cannot find your Workflows chat, so a draft cannot be updated. Open Workflows in Teams once, or pass the chat ref.");
    return { mode: "update-draft", messageId, conversationId: chat, label: "your Workflows chat", key: `card:update:draft:${messageId}` };
  }
  const thread = /^(.*);messageid=(\d{10,})$/.exec(left);
  const base = thread ? thread[1]! : left;
  const conversation = await resolveConversation(base);
  if (isChannelId(conversation.id)) {
    const id = thread ? `${thread[2]}/replies/${messageId}` : messageId;
    return {
      mode: "update", channel: conversation.id, messageId: id, conversationId: conversation.id, label: conversation.label,
      key: `card:update:${conversation.id}:${id}`,
    };
  }
  if (thread) throw new Error("A thread reply must be in a channel.");
  const chat = await findWorkflowsChat();
  if (!chat || chat !== conversation.id) throw new Error(`${conversation.label} is not a channel or your Workflows chat. A card update works only in those. ${forms}`);
  return { mode: "update-draft", messageId, conversationId: chat, label: "your Workflows chat", key: `card:update:draft:${messageId}` };
}

/** Previews an edit of a posted card. Only `--allow-write` sends it, through the card router, once per target and card hash. */
export async function cardUpdate(input: string, source: CardSource, options: CardUpdateOptions): Promise<void> {
  const loaded = await loadSource(source);
  const target = await parseUpdateTarget(input);
  const payload = cardPayload(loaded.card, { mode: target.mode, channel: target.channel, messageId: target.messageId });
  const payloadBytes = checkPayloadSize(payload);
  const where = target.mode === "update-draft" ? `the draft ${target.messageId} in your Workflows chat` : `the card ${target.messageId} in ${target.label}`;

  if (!options.allowWrite) {
    if (options.json) {
      console.log(JSON.stringify({
        ok: true, dry_run: true, effect: "card-update", mode: target.mode, conversation: { id: target.conversationId, label: target.label },
        message_id: target.messageId, ...(loaded.template ? { template: loaded.template } : {}), version: loaded.card.version, card_bytes: loaded.bytes, payload_bytes: payloadBytes, outline: cardOutline(loaded.card),
        ...(loaded.images.length ? { images: imageJson(loaded) } : {}),
      }, null, 2));
      return;
    }
    console.log(`Dry run (effect: card update): would replace ${where} with this card.`);
    if (loaded.template) console.log(`  template: ${loaded.template}`);
    console.log(`  version: ${loaded.card.version}, ${loaded.bytes} bytes`);
    for (const line of cardOutline(loaded.card)) console.log(`  ${line}`);
    for (const line of imageLines(loaded)) console.log(`  ${line}`);
    return;
  }

  const url = await routerUrl();
  const who = await identity();
  if (target.mode === "update") await paceWrite(who.tenantId, { noWait: options.noWait });
  const key = ledgerKey({ conversationId: target.key, content: loaded.hash, mentionMris: [], contentType: "AdaptiveCardUpdate" });
  const context = { label: target.label, conversationId: target.conversationId };
  await submit({ key, mode: target.mode, context, input, url, payload, tenantId: who.tenantId });
  if (options.json) console.log(JSON.stringify({ ok: true, effect: "card-update", mode: target.mode, conversation: target.conversationId, message_id: target.messageId, status: 202, message: UPDATED }, null, 2));
  else console.log(UPDATED);
}
