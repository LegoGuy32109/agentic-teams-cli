import { readFile } from "node:fs/promises";
import { findChannel, kindOf, loadConversations, resolveConversation } from "../conversations.ts";
import { createDirectChat } from "../direct.ts";
import { attachPreviews } from "../preview.ts";
import { buildScheduledBody, buildSendBody, checkBodySize, composeHtml, composePlain, composeRich, extractFences, newClientMessageId, unresolvedMentionWarning, type LinkEntry, type MentionEntry } from "../compose.ts";
import { htmlToText, sameVisibleText } from "../message.ts";
import { clearRecord, ledgerKey, listRecords, markSent, markUncertain, removeRecord, reserve } from "../ledger.ts";
import { READ_ONLY_STREAMS, encodeConversation, isChannelId, threadRootId } from "../refs.ts";
import { permalinkFor, type PreviewResult, type ReceiptResult } from "../operations.ts";
import { paceWrite } from "../pace.ts";
import { describeSendAt, parseSendAt } from "../schedule.ts";
import { fileSharing } from "../settings.ts";
import { withEmojiStyle, type EmojiStyle } from "../emoji.ts";
import { identity } from "../tokens.ts";
import { amsHosts, request, TeamsError } from "../transport.ts";
import { imageHtml, loadImage, MAX_IMAGES, placeholderId, uploadImage, displaySize, type LoadedImage } from "../image.ts";
import { resolveMentionTokens, type MentionScope } from "../users.ts";
import { checkFileSet, chatFileEntry, channelFileEntry, describeDestination, loadFile, placeholderChannelUpload, placeholderChatUpload, planUploads, runUploads, UploadFailure, MAX_FILES, type LoadedFile, type UploadPlan } from "../upload.ts";

export type SendOptions = {
  json: boolean;
  format?: string;
  html?: string;
  verify: boolean;
  allowWrite: boolean;
  threadTs?: string;
  noWait?: boolean;
  /** No preview card on the first link. */
  noPreview?: boolean;
  /** Schedules the message for this time instead of sending it now. See `parseSendAt()`. */
  at?: string;
  /** Replaces the clock for `--at`. Tests only. */
  now?: Date;
  /** Paths of images to attach, up to 10. */
  images?: string[];
  /** Who can open a file sent to a chat: link, people, or private. See `fileSharing()`. */
  access?: string;
  /** The access an invited member gets: edit or view. */
  role?: string;
  /** Paths of files to upload and attach, up to 10, each up to 250 MB. */
  files?: string[];
  /** `plain` keeps emoji characters as typed. The default writes the Teams emoji markup. */
  emoji?: EmojiStyle;
};

/** Checks the text, --format, and --html arguments. Shared by send and edit. */
export function validateComposeArgs(words: string[], options: { html?: string; format?: string; emoji?: string }, usage: string): string {
  if (options.emoji !== undefined && options.emoji !== "teams" && options.emoji !== "plain") throw new Error("--emoji must be teams or plain.");
  if (options.html !== undefined && words.length) throw new Error("Use either <text> or --html.");
  if (options.html !== undefined && options.format !== undefined) throw new Error("Use either --format or --html.");
  if (options.format !== undefined && options.format !== "plain" && options.format !== "rich") throw new Error("--format must be plain or rich.");
  const text = words.join(" ");
  if (options.html === undefined && words.length === 0) throw new Error(usage);
  if (options.html === undefined && text.trim() === "") throw new Error("The message is empty.");
  return text;
}

/**
 * Turns text or --html into the stored content, with mention entries. Every write uses this, so the
 * rules are the same for send and edit: `@{Name}` resolves in text and in the text of --html, --html
 * cannot hold hand-written mention markup, and an empty body is refused. `warning` names plain `@Word`
 * text that will not notify anyone.
 */
export async function composeMessage(text: string, options: { html?: string; format?: string; emoji?: EmojiStyle }, scope?: MentionScope): Promise<{ content: string; visible: string; entries: MentionEntry[]; links: LinkEntry[]; warning?: string }> {
  return withEmojiStyle(options.emoji, () => composeWith(text, options, scope));
}

/** What `@{everyone}` and `@{channel}` are checked against: the kind of the conversation, and the channel name when asked. */
export async function mentionScope(conversationId: string, pending = false): Promise<MentionScope> {
  const who = await identity();
  const kind = kindOf(await loadConversations(who.tenantId), conversationId, pending);
  return { conversationId, kind, channelName: async () => (await findChannel(conversationId)).channel.name };
}

async function composeWith(text: string, options: { html?: string; format?: string }, scope?: MentionScope): Promise<{ content: string; visible: string; entries: MentionEntry[]; links: LinkEntry[]; warning?: string }> {
  if (options.html !== undefined) {
    const raw = options.html.startsWith("@") ? await readFile(options.html.slice(1), "utf8") : options.html;
    if (raw.trim() === "") throw new Error("The message is empty.");
    if (/schema\.skype\.com\/Mention/i.test(raw)) throw new Error("--html cannot contain hand-written mention markup. Write @{Name} in the text and the CLI builds the mention.");
    const composed = await composeHtml(raw, scope);
    const warning = unresolvedMentionWarning(composed.warnText);
    return { content: composed.content, visible: composed.visible, entries: composed.entries, links: [], ...(warning ? { warning } : {}) };
  }
  // Rich text sets fenced code aside first, so a mention token inside code stays code.
  const fenced = options.format === "rich" ? extractFences(text) : { text, blocks: [] };
  const resolved = await resolveMentionTokens(fenced.text, scope);
  const composed = options.format === "rich" ? composeRich(resolved.text, resolved.mentions, fenced.blocks) : composePlain(resolved.text, resolved.mentions);
  const warning = unresolvedMentionWarning(resolved.text);
  return { content: composed.html, visible: resolved.text, entries: composed.entries, links: composed.links, ...(warning ? { warning } : {}) };
}

/** Who can open the files, in words, for a preview. */
function accessWords(plan: Extract<UploadPlan, { route: "onedrive" }>): string {
  if (plan.sharing.access === "private" || !plan.invite) return "only you can open it (no invite, no link)";
  const { recipients, skipped } = plan.invite;
  const people = `invite ${recipients.length} member${recipients.length === 1 ? "" : "s"} to ${plan.sharing.role}: ${recipients.map(item => item.name).join(", ")}`
    + (skipped.length ? ` (skipped ${skipped.length}: ${skipped.map(item => item.reason).join(", ")})` : "");
  if (plan.sharing.access === "people") return `${people}, no link`;
  // With the default link asked for, SharePoint makes only the link and adds nobody by name (proven 2026-10-19).
  return `share with the tenant's default sharing link for ${recipients.length} member${recipients.length === 1 ? "" : "s"}: ${recipients.map(item => item.name).join(", ")}`
    + (skipped.length ? ` (skipped ${skipped.length}: ${skipped.map(item => item.reason).join(", ")})` : "")
    + " (nobody is added by name; --access=people for that)";
}

/** Previews a message by default. Only `--allow-write` sends, and only once per send-once key. */
export async function send(input: string, words: string[], options: SendOptions): Promise<void> {
  if (options.threadTs !== undefined) options = { ...options, threadTs: threadRootId(options.threadTs, "--thread-ts") };
  const imagePaths = options.images ?? [];
  if (imagePaths.length > MAX_IMAGES) throw new Error(`A message can hold ${MAX_IMAGES} images at most.`);
  const filePaths = options.files ?? [];
  if (filePaths.length > MAX_FILES) throw new Error(`A message can hold ${MAX_FILES} files at most.`);
  // An image or a file alone is a valid message. The text rules then apply only when text or --html is given.
  const imageOnly = (imagePaths.length > 0 || filePaths.length > 0) && words.length === 0 && options.html === undefined;
  const text = imageOnly ? "" : validateComposeArgs(words, options, "Usage: teams-cli send <conversation> <text...> [--image PATH]... [--file PATH]... [--format=rich] [--html X] [--at TIME] [--allow-write]");
  if (imageOnly && options.format !== undefined) throw new Error("--format needs message text.");
  if (options.at !== undefined && options.verify) throw new Error("--verify cannot be used with --at. A scheduled message has not posted yet.");
  // A scheduled message with images is not captured, so the pair is refused.
  if (options.at !== undefined && imagePaths.length) throw new Error("--at cannot be used with --image. A scheduled message with images is not verified.");
  // A scheduled message with files is not captured, so the pair is refused.
  const sendAt = options.at === undefined ? undefined : parseSendAt(options.at, options.now);

  const who = await identity();
  const target = await resolveConversation(input, { allowNew: true });
  if (READ_ONLY_STREAMS.has(target.id)) throw new Error("That stream is read-only.");
  if (options.threadTs && !isChannelId(target.id)) throw new Error("--thread-ts works only with channels.");
  const pathId = options.threadTs ? `${target.id};messageid=${options.threadTs}` : target.id;
  const path = sendAt ? "/users/ME/drafts" : `/users/ME/conversations/${encodeConversation(pathId)}/messages`;

  if (imagePaths.length) amsHosts(who.region);
  const images: LoadedImage[] = [];
  for (const imagePath of imagePaths) images.push(await loadImage(imagePath));
  const files: LoadedFile[] = [];
  for (const filePath of filePaths) files.push(await loadFile(filePath));
  checkFileSet(files);

  const scope = await mentionScope(target.id, !!target.pending);
  const { content: textContent, visible, entries, links, warning } = imageOnly
    ? { content: "", visible: "", entries: [] as MentionEntry[], links: [] as LinkEntry[], warning: undefined }
    : await composeMessage(text, options, scope);
  // Where the files go and who can open them. Reads only, so a preview can show it.
  const sharing = await fileSharing({ access: options.access, role: options.role });
  const plan = files.length ? await planUploads({ id: target.id, kind: scope.kind, pending: target.pending }, files, who.mri, sharing) : undefined;
  if (warning) console.error(warning);
  if (!(await attachPreviews(links, who.region, !options.noPreview))) console.error("Link preview unavailable. The link is sent without a card.");

  // The text comes first, then one paragraph per image. A preview and the send-once key use placeholder ids,
  // because the real ids exist only after the upload, and the upload comes after the ledger check.
  const entriesFor = (plan: UploadPlan, files: LoadedFile[]) => files.map((file, index) =>
    plan.route === "channel" ? channelFileEntry(placeholderChannelUpload(file, index, plan.target)) : chatFileEntry(placeholderChatUpload(file, index, plan.site)));
  const bodyFor = (ids: string[], fileEntries: unknown[] = plan ? entriesFor(plan, files) : []) => {
    const content = textContent + images.map((image, index) => imageHtml(who.region, ids[index]!, image)).join("");
    return buildSendBody({
      conversationId: target.id,
      content,
      mentions: entries,
      links,
      mri: who.mri,
      name: who.name,
      now: new Date().toISOString(),
      clientMessageId: newClientMessageId(),
      threadTs: options.threadTs,
      channelPost: isChannelId(target.id),
      amsReferences: ids,
      files: fileEntries,
    });
  };
  const placeholders = images.map((_, index) => placeholderId(index));
  const message = bodyFor(placeholders);
  // A scheduled reply names the thread as `<channel>;messageid=<root>`, as the web client did on 2026-10-19.
  const body = sendAt ? buildScheduledBody(message, sendAt, pathId) : message;
  const content = message.content;
  checkBodySize(message);
  const fullPath = `/api/chatsvc/${who.region}/v1${path}`;

  const when = sendAt ? describeSendAt(sendAt) : undefined;
  if (plan?.route === "channel" && plan.conflicts.length && options.allowWrite) {
    throw new Error(`${plan.conflicts.join(", ")} already ${plan.conflicts.length === 1 ? "exists" : "exist"} in ${plan.target.folder}. The upload would not replace ${plan.conflicts.length === 1 ? "it" : "them"}. Nothing was uploaded. Rename the file.`);
  }
  if (!options.allowWrite) {
    const preview: PreviewResult = {
      ok: true,
      dry_run: true,
      effect: sendAt ? "scheduled_message" : "message",
      ...(sendAt && when ? { send_at: when.utc, send_at_local: when.local } : {}),
      method: "POST",
      path: fullPath,
      conversation: { id: target.id, label: target.label },
      ...(target.pending ? { new_chat: true } : {}),
      body,
      ...(images.length ? { images: images.map(image => ({ file: image.name, bytes: image.size, type: image.kind, width: image.width, height: image.height, display: displaySize(image.width, image.height) })) } : {}),
      ...(plan ? {
        files: files.map(file => ({ file: file.name, bytes: file.size, destination: describeDestination(plan, file), ...(plan.route === "channel" && plan.conflicts.includes(file.name) ? { conflict: "A file with this name exists. A send is refused." } : {}) })),
        ...(plan.route === "channel" ? { channel_folder: plan.target.folder } : {}),
        ...(plan.route === "onedrive" ? { sharing: plan.sharing } : {}),
        ...(plan.route === "onedrive" && plan.invite ? { invite: { roles: [plan.sharing.role === "view" ? "read" : "write"], default_link: plan.sharing.access === "link", recipients: plan.invite.recipients.map(item => ({ name: item.name, object_id: item.objectId })), skipped: plan.invite.skipped } } : {}),
      } : {}),
      ...(warning ? { warning } : {}),
    };
    if (options.json) console.log(JSON.stringify(preview, null, 2));
    else if (sendAt && when) console.log(`Dry run (effect: scheduled message): would schedule for ${target.label}${target.pending ? " (new chat, created on send)" : ""} at ${when.local} (${when.utc}): ${options.html !== undefined ? visible : text}${plan ? ` [${files.length} file${files.length === 1 ? "" : "s"} upload now: ${files.map(file => `${file.name} to ${describeDestination(plan, file)}`).join(", ")}${plan.route === "onedrive" ? `; ${accessWords(plan)}, now` : ""}]` : ""}`);
    else console.log(`Dry run (effect: message): would send to ${target.label}${target.pending ? " (new chat, created on send)" : ""}${options.threadTs ? ` (thread ${options.threadTs})` : ""}: ${options.html !== undefined ? visible : text}${images.length ? `${imageOnly ? "" : " "}[${images.length} image${images.length === 1 ? "" : "s"}: ${images.map(image => `${image.name} ${image.size} bytes ${image.width}x${image.height}`).join(", ")}; no upload in a preview]` : ""}${plan ? `${imageOnly && !images.length ? "" : " "}[${files.length} file${files.length === 1 ? "" : "s"}: ${files.map(file => `${file.name} ${file.size} bytes to ${describeDestination(plan, file)}`).join(", ")}${plan.route === "onedrive" ? `; ${accessWords(plan)}` : ""}${plan.route === "channel" && plan.conflicts.length ? `; ${plan.conflicts.join(", ")} already exists, so a send is refused` : ""}; no upload in a preview]` : ""}`);
    return;
  }

  if (target.pending) {
    // The chat is created only now, after the message composed and the duplicate guard would run. The service returns the existing chat for the pair.
    const created = await createDirectChat(who.mri, target.pending.mri);
    if (created !== target.id) throw new Error("Teams returned a different chat than expected. Nothing was sent.");
  }
  await paceWrite(who.tenantId, { noWait: options.noWait });
  const key = ledgerKey({ conversationId: pathId, content, mentionMris: entries.map(entry => entry.mri), contentType: "Text", images: images.map(image => image.digest), files: files.map(file => file.digest), ...(sendAt ? { sendAt: sendAt.getTime() } : {}) });
  const context = { label: target.label, conversationId: pathId };
  await reserve(who.tenantId, key, { ...context, input });

  // Images and --at are exclusive. Files may be scheduled: they upload now, and the draft carries properties.files,
  // as the web client's scheduled channel post with a file did (capture 13).
  let sendMessage = message;
  let uploadedFiles: string[] = [];
  if (images.length || files.length) {
    const ids: string[] = [];
    let fileEntries: unknown[] = [];
    try {
      for (const image of images) ids.push(await uploadImage(target.id, image));
      if (plan) fileEntries = await runUploads(plan, files);
    } catch (error) {
      // No message was posted, so the send-once record would only block a retry.
      await removeRecord(who.tenantId, key);
      if (error instanceof UploadFailure) {
        throw new Error(`${error.message} No message was sent.${error.leftovers.length ? ` These files are already on SharePoint and can be deleted: ${error.leftovers.join("; ")}` : ""}`);
      }
      throw error;
    }
    uploadedFiles = fileEntries.map((entry: any) => String(entry?.fileInfo?.fileUrl ?? ""));
    sendMessage = bodyFor(ids, fileEntries);
    checkBodySize(sendMessage);
  }
  const sendBody = images.length || files.length ? (sendAt ? buildScheduledBody(sendMessage, sendAt, pathId) : sendMessage) : body;
  const filesLeft = uploadedFiles.length ? ` The files are already uploaded: ${uploadedFiles.join("; ")}` : "";

  let response: any;
  try {
    response = await request("chat", "POST", path, { body: sendBody, send: !sendAt });
  } catch (error) {
    if (error instanceof TeamsError && error.uncertain) {
      await markUncertain(who.tenantId, key, context);
      throw new Error(`Outcome unknown: Teams did not confirm this ${sendAt ? "schedule" : "send"}. Do not resend. Check with: ${sendAt ? "teams-cli scheduled" : `teams-cli history ${input} --top=5`}${filesLeft}`);
    }
    await removeRecord(who.tenantId, key);
    if (filesLeft) throw new Error(`${error instanceof Error ? error.message : String(error)} The message was not sent.${filesLeft}`);
    throw error;
  }

  const ts = String(response?.OriginalArrivalTime ?? "");
  if (sendAt && when) {
    const draftId = `draft:${ts}`;
    await markSent(who.tenantId, key, draftId, context);
    const scheduledReceipt = { ok: true, effect: "scheduled_message", conversation: target.id, draft_id: ts, send_at: when.utc, send_at_local: when.local, ...(warning ? { warning } : {}) };
    if (options.json) console.log(JSON.stringify(scheduledReceipt, null, 2));
    else console.log(`Scheduled for ${target.label} at ${when.local} (${when.utc}) as ${ts}. Cancel with: teams-cli scheduled cancel ${ts} --allow-write`);
    return;
  }
  const id = `${pathId}#${ts}`;
  await markSent(who.tenantId, key, id, context);

  let verification: ReceiptResult["verification"];
  if (options.verify) {
    const expected = htmlToText(sendMessage.content);
    try {
      const sent: any = await request("chat", "GET", `${path}/${ts}`);
      const actualHtml = typeof sent?.content === "string" ? sent.content : (typeof sent?.messages?.[0]?.content === "string" ? sent.messages[0].content : "");
      const actual = htmlToText(actualHtml);
      verification = { ok: sameVisibleText(actual, expected), expected, actual };
    } catch (error) {
      console.error(error instanceof Error ? error.message : String(error));
      verification = { ok: false, expected, actual: "" };
    }
    if (!verification.ok) process.exitCode = 1;
  }

  const receipt: ReceiptResult = {
    ok: true,
    effect: "message",
    conversation: target.id,
    id,
    ts,
    created: /^\d{10,}$/.test(ts) ? new Date(Number(ts)).toISOString() : "",
    permalink: permalinkFor(target.id, options.threadTs ?? ts),
    ...(options.threadTs ? { reply_to: options.threadTs } : {}),
    ...(warning ? { warning } : {}),
    ...(verification ? { verification } : {}),
  };
  if (options.json) {
    console.log(JSON.stringify(receipt, null, 2));
    return;
  }
  console.log(`Sent to ${target.label} as ${id}`);
  console.log(receipt.permalink);
  if (verification) {
    console.log(verification.ok ? "Verification: ok" : `Verification failed. Expected: ${verification.expected}. Actual: ${verification.actual}`);
  }
}

/** Lists send-once records for this tenant. */
export async function ledgerList(json: boolean): Promise<void> {
  const who = await identity();
  const records = await listRecords(who.tenantId);
  if (json) {
    console.log(JSON.stringify(records, null, 2));
    return;
  }
  if (!records.length) console.log("No send records.");
  for (const record of records) console.log(`${record.key}  ${record.state}  ${record.at}  ${record.label ?? ""}`);
}

/** Removes one send-once record. Use it after you confirm a message did not post. */
export async function ledgerClear(key: string): Promise<void> {
  const who = await identity();
  console.log(await clearRecord(who.tenantId, key) ? "Cleared." : "No send record with that key.");
}
