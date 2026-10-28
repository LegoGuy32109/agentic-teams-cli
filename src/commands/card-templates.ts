import { messageCards } from "../card.ts";
import { findWorkflowsChat } from "../card-ref.ts";
import { TEMPLATES, findTemplate } from "../card-templates/templates.ts";
import { resolveConversation } from "../conversations.ts";
import { findEmoji } from "../emoji.ts";
import { reactionsFrom } from "../message.ts";
import { resolveReactionKey } from "../reactions.ts";
import { encodeConversation, isChannelId } from "../refs.ts";
import { identity } from "../tokens.ts";
import { request } from "../transport.ts";

/** Prints the templates, each with its purpose and data fields. */
export function cardTemplates(json: boolean): void {
  if (json) {
    console.log(JSON.stringify(TEMPLATES.map(template => ({ name: template.name, purpose: template.purpose, fields: template.fields, votes: template.votes !== undefined })), null, 2));
    return;
  }
  const width = Math.max(...TEMPLATES.map(template => template.name.length));
  for (const template of TEMPLATES) {
    console.log(`${template.name.padEnd(width)}  ${template.purpose}`);
    console.log(`${" ".repeat(width)}  data: ${template.fields}`);
  }
  console.log("\nExample data: teams-cli card example <template>. Send: teams-cli card send <channel> --template <template> --data <file.json>");
}

/** Prints the example data of one template, as JSON that `card send --data` accepts. */
export function cardExample(name: string): void {
  console.log(JSON.stringify(findTemplate(name).example, null, 2));
}

type Place = { conversationId: string; messageId: string; label: string };

/** Reads the same refs as `card update`: `<channel>#<id>`, `<channel>;messageid=<root>#<reply>`, or `draft#<id>`. A chat works too. */
async function votePlace(input: string): Promise<Place> {
  const forms = "Use <conversation>#<id>, <channel>;messageid=<root>#<reply>, or draft#<id>.";
  const index = input.lastIndexOf("#");
  const messageId = index > 0 ? input.slice(index + 1) : "";
  if (!/^\d{10,}$/.test(messageId)) throw new Error(`"${input}" is not the ref of a message. ${forms}`);
  const left = input.slice(0, index);
  if (left.toLowerCase() === "draft") {
    const chat = await findWorkflowsChat();
    if (!chat) throw new Error("Cannot find your Workflows chat, so a draft cannot be read. Pass the chat ref instead.");
    return { conversationId: chat, messageId, label: "your Workflows chat" };
  }
  const thread = /^(.*);messageid=(\d{10,})$/.exec(left);
  const conversation = await resolveConversation(thread ? thread[1]! : left);
  if (thread && !isChannelId(conversation.id)) throw new Error("A thread reply must be in a channel.");
  return { conversationId: conversation.id, messageId, label: conversation.label };
}

/** The option emoji and labels of a vote card, in card order. The templates mark each option text with the id `vote-<n>`. */
export function voteOptions(cards: unknown[]): Array<{ emoji: string; label: string }> {
  const found: Array<{ n: number; text: string }> = [];
  const visit = (node: unknown): void => {
    if (Array.isArray(node)) node.forEach(visit);
    else if (node && typeof node === "object") {
      const item = node as Record<string, unknown>;
      const match = typeof item.id === "string" ? /^vote-(\d+)$/.exec(item.id) : null;
      if (match && item.type === "TextBlock" && typeof item.text === "string") found.push({ n: Number(match[1]), text: item.text });
      Object.values(item).forEach(visit);
    }
  };
  for (const card of cards) visit(card);
  return found.sort((a, b) => a.n - b.n).map(item => {
    const [emoji = "", ...rest] = item.text.trim().split(/\s+/);
    return { emoji, label: rest.join(" ") };
  });
}

const baseKey = (key: string) => key.replace(/-tone\d$/, "");

/** The keys that count as one option: its key, its catalog key, and the same without a skin tone. */
function optionKeys(emoji: string): Set<string> {
  const keys = new Set<string>();
  try { keys.add(baseKey(resolveReactionKey(emoji))); } catch { /* A character with no key counts for nothing. */ }
  const catalog = findEmoji(emoji)?.base.id;
  if (catalog) keys.add(catalog);
  return keys;
}

export type VoteCount = { emoji: string; label: string; key: string; votes: number };
export type VoteResult = { options: VoteCount[]; other: Array<{ key: string; emoji?: string; votes: number }>; ownExcluded: string[] };

/**
 * Counts the votes on a message. A vote is a reaction of someone else under the option's key. Your own reaction is the
 * seed that `card send` added, so it is left out. Reactions that are not options are listed apart.
 */
export function countVotes(reactions: Array<{ key: string; mris: string[] }>, options: Array<{ emoji: string; label: string }>, mine: string): VoteResult {
  const ownExcluded = reactions.filter(reaction => reaction.mris.includes(mine)).map(reaction => reaction.key);
  const others = reactions.map(reaction => ({ key: reaction.key, count: new Set(reaction.mris.filter(mri => mri !== mine)).size }));
  const used = new Set<string>();
  const counted = options.map(option => {
    const keys = optionKeys(option.emoji);
    let votes = 0;
    others.forEach((entry, index) => {
      if (keys.has(baseKey(entry.key))) { votes += entry.count; used.add(String(index)); }
    });
    let key = "";
    try { key = resolveReactionKey(option.emoji); } catch { key = ""; }
    return { ...option, key, votes };
  });
  const other = others
    .map((entry, index) => ({ ...entry, index }))
    .filter(entry => !used.has(String(entry.index)) && entry.count > 0)
    .map(entry => ({ key: entry.key, emoji: findEmoji(entry.key)?.unicode, votes: entry.count }));
  return { options: counted, other, ownExcluded };
}

/** Reads the reactions of a card message and prints the votes per option. It reads only. */
export async function cardVotes(input: string, options: { json: boolean }): Promise<void> {
  const place = await votePlace(input);
  const who = await identity();
  const raw: any = await request("chat", "GET", `/users/ME/conversations/${encodeConversation(place.conversationId)}/messages/${place.messageId}`);
  const message = Array.isArray(raw?.messages) ? raw.messages[0] : raw;
  if (message?.properties?.deletetime) throw new Error("That message was deleted.");
  const reactions = reactionsFrom(message?.properties?.emotions);
  const choices = voteOptions(messageCards(message));
  // A card without vote markers has no known options, so every reaction is listed.
  const result = countVotes(reactions, choices, who.mri);
  const total = [...result.options, ...(choices.length ? [] : result.other)].reduce((sum, item) => sum + item.votes, 0);

  if (options.json) {
    console.log(JSON.stringify({
      ok: true, ref: input, conversation: { id: place.conversationId, label: place.label }, message_id: place.messageId,
      options_found: choices.length > 0, options: result.options, other: result.other, total_votes: total, own_reactions_excluded: result.ownExcluded,
    }, null, 2));
    return;
  }
  console.log(`Votes on ${input} in ${place.label}. Your own reactions are not counted.`);
  if (!choices.length) console.log("This message has no vote options that the CLI knows, so every reaction is listed.");
  const width = Math.max(0, ...result.options.map(item => `${item.emoji} ${item.label}`.length));
  for (const item of result.options) console.log(`  ${`${item.emoji} ${item.label}`.padEnd(width)}  ${item.votes}`);
  for (const item of result.other) console.log(`  ${(item.emoji ?? item.key).padEnd(width)}  ${item.votes}${choices.length ? "  (not an option)" : ""}`);
  console.log(`Total votes: ${total}`);
}
