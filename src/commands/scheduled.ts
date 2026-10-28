import { labelFor, loadConversations } from "../conversations.ts";

/** A scheduled reply names `<channel>;messageid=<root>`. It shows as the channel name and the thread root. */
function threadLabel(cache: Parameters<typeof labelFor>[0], conversation: string): string {
  const thread = /^(.+);messageid=(\d+)$/.exec(conversation);
  const channel = thread ? thread[1]! : conversation;
  const label = labelFor(cache, channel) ?? channel;
  return thread ? `${label} (thread ${thread[2]})` : label;
}
import { htmlToText } from "../message.ts";
import { filesOf } from "./files.ts";
import { paceWrite } from "../pace.ts";
import { describeSendAt } from "../schedule.ts";
import { identity } from "../tokens.ts";
import { request, TeamsError } from "../transport.ts";

export type ScheduledMessage = {
  id: string;
  conversation: string;
  label: string;
  send_at: string;
  send_at_local: string;
  text: string;
  /** Addresses of files the draft carries. They are uploaded when the message is scheduled, not when it sends. */
  files?: string[];
};

/** Most pages of scheduled drafts to follow. A page holds up to 200. */
export const MAX_DRAFT_PAGES = 5;
const PREVIEW_CHARS = 80;

const LIST_QUERY = { view: "msnp24Equivalent|supportsMessageProperties", pageSize: "200", startTime: "1" };

/** Takes the `syncState` value out of the paging link. Returns undefined when there is none. */
function nextSyncState(link: unknown): string | undefined {
  if (typeof link !== "string") return undefined;
  try {
    return new URL(link).searchParams.get("syncState") || undefined;
  } catch {
    return undefined;
  }
}

/**
 * Reads the live scheduled messages. A cancelled draft stays in the list with `properties.deletetime`,
 * so it is skipped, and so is any draft that is not a scheduled one. Plain drafts never reach the server.
 */
export async function listScheduled(): Promise<ScheduledMessage[]> {
  const who = await identity();
  const cache = await loadConversations(who.tenantId);
  const found: ScheduledMessage[] = [];
  let page: any = await request("chat", "GET", "/users/ME/drafts", { query: LIST_QUERY });
  let last = "";
  for (let pages = 1; ; pages++) {
    for (const draft of Array.isArray(page?.drafts) ? page.drafts : []) {
      if (draft?.draftType !== "ScheduledDraft" || draft?.properties?.deletetime) continue;
      const sendAt = Number(draft?.draftDetails?.sendAt);
      if (typeof draft.id !== "string" && typeof draft.id !== "number") continue;
      if (!Number.isFinite(sendAt)) continue;
      const conversation = typeof draft.innerThreadId === "string" ? draft.innerThreadId : "";
      const at = new Date(sendAt);
      found.push({
        id: String(draft.id),
        conversation,
        label: threadLabel(cache, conversation),
        send_at: at.toISOString(),
        send_at_local: describeSendAt(at).local,
        text: htmlToText(typeof draft.content === "string" ? draft.content : "").replace(/\s+/g, " ").trim().slice(0, PREVIEW_CHARS),
        ...(filesOf(draft).length ? { files: filesOf(draft).map(file => file.url) } : {}),
      });
    }
    // The `syncState` link points at `conversations/48:drafts/messages`, and Teams answers it with 400
    // (checked 2026-10-19). Only its `syncState` value is used, sent back to the drafts endpoint.
    const next = nextSyncState(page?._metadata?.syncState);
    if (!next || next === last || pages >= MAX_DRAFT_PAGES || !Array.isArray(page?.drafts) || page.drafts.length === 0) break;
    last = next;
    page = await request("chat", "GET", "/users/ME/drafts", { query: { ...LIST_QUERY, syncState: next } });
  }
  return found.sort((a, b) => a.send_at.localeCompare(b.send_at));
}

export async function scheduled(json: boolean): Promise<void> {
  const items = await listScheduled();
  if (json) {
    console.log(JSON.stringify(items, null, 2));
    return;
  }
  if (!items.length) console.log("No scheduled messages.");
  for (const item of items) console.log(`${item.id}  ${item.send_at_local}  ${item.label}  ${item.text}`);
}

/** Cancels one scheduled message. Preview by default. Only `--allow-write` deletes it. */
export async function scheduledCancel(id: string, options: { json: boolean; allowWrite: boolean; noWait?: boolean }): Promise<void> {
  if (!/^\d{10,}$/.test(id)) throw new Error(`"${id}" is not a draft id. Run: teams-cli scheduled`);
  const item = (await listScheduled()).find(entry => entry.id === id);
  if (!item) throw new Error(`No scheduled message has the id ${id}. Run: teams-cli scheduled`);
  const who = await identity();
  const path = `/api/chatsvc/${who.region}/v1/users/ME/drafts/${id}`;

  if (!options.allowWrite) {
    if (options.json) console.log(JSON.stringify({ ok: true, dry_run: true, effect: "scheduled_cancel", method: "DELETE", path, scheduled: item }, null, 2));
    else console.log(`Dry run (effect: scheduled cancel): would cancel ${id}, to ${item.label} at ${item.send_at_local} (${item.send_at}): ${item.text}`);
    return;
  }

  await paceWrite(who.tenantId, { noWait: options.noWait });
  try {
    await request("chat", "DELETE", `/users/ME/drafts/${id}`);
  } catch (error) {
    if (error instanceof TeamsError && error.uncertain) {
      throw new Error(`Outcome unknown: Teams did not confirm the cancel. Check with: teams-cli scheduled`);
    }
    throw error;
  }
  if (options.json) console.log(JSON.stringify({ ok: true, effect: "scheduled_cancel", id, scheduled: item }, null, 2));
  else console.log(`Cancelled scheduled message ${id} (${item.label}, ${item.send_at_local}).`);
  // Cancelling the draft does not remove its files, and chat members keep the access the upload gave them.
  if (item.files?.length && !options.json) console.log(`Its files stay on SharePoint, and anyone invited keeps access until they are deleted: ${item.files.join("; ")}`);
}
