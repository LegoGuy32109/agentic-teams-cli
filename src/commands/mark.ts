import { resolveConversation } from "../conversations.ts";
import { paceWrite } from "../pace.ts";
import { READ_ONLY_STREAMS, encodeConversation } from "../refs.ts";
import { identity } from "../tokens.ts";
import { request, TeamsError } from "../transport.ts";
import { findUnread } from "./unread.ts";

export type MarkOptions = { json: boolean; allowWrite: boolean; noWait?: boolean; allUnread?: boolean };

type Latest = { id: string; clientMessageId: string };

/** Builds the consumption horizon: `<last message id>;<now ms>;<last client message id>`. */
export function horizonValue(latest: Latest, now: number): string {
  return `${latest.id};${now};${latest.clientMessageId}`;
}

/** Reads the newest message that has a client message id. Returns undefined for an empty conversation. */
export async function latestMessage(conversationId: string): Promise<Latest | undefined> {
  const page: any = await request("chat", "GET", `/users/ME/conversations/${encodeConversation(conversationId)}/messages`, {
    query: { view: "msnp24Equivalent|supportsMessageProperties", pageSize: "10", startTime: "1" },
  });
  const found = (Array.isArray(page?.messages) ? page.messages : [])
    .filter((item: any) => /^\d{10,}$/.test(String(item?.id)) && /^\d+$/.test(String(item?.clientmessageid ?? "")))
    .sort((a: any, b: any) => Number(b.id) - Number(a.id))[0];
  return found ? { id: String(found.id), clientMessageId: String(found.clientmessageid) } : undefined;
}

type Result = { id: string; label: string; status: "marked" | "would mark" | "empty" | "failed"; horizon?: string; error?: string };

/**
 * Marks conversations read by writing the consumption horizon. Preview by default. Only `--allow-write` writes,
 * one conversation at a time, paced to the write limit. It is not covered by the send-once ledger.
 */
export async function markRead(input: string | undefined, options: MarkOptions): Promise<void> {
  const who = await identity();
  let targets: Array<{ id: string; label: string }>;
  if (options.allUnread) {
    targets = (await findUnread()).conversations;
  } else {
    const target = await resolveConversation(input!);
    if (READ_ONLY_STREAMS.has(target.id)) throw new Error("That stream is read-only.");
    targets = [target];
  }

  const results: Result[] = [];
  for (const target of targets) {
    try {
      const latest = await latestMessage(target.id);
      if (!latest) {
        results.push({ id: target.id, label: target.label, status: "empty" });
        continue;
      }
      const horizon = horizonValue(latest, Date.now());
      if (!options.allowWrite) {
        results.push({ id: target.id, label: target.label, status: "would mark", horizon });
        continue;
      }
      await paceWrite(who.tenantId, { noWait: options.noWait, send: false });
      try {
        await request("chat", "PUT", `/users/ME/conversations/${encodeConversation(target.id)}/properties`, {
          query: { name: "consumptionhorizon" },
          body: { consumptionhorizon: horizon },
        });
      } catch (error) {
        if (error instanceof TeamsError && error.uncertain) throw new Error("Outcome unknown: Teams did not confirm this change.");
        throw error;
      }
      results.push({ id: target.id, label: target.label, status: "marked", horizon });
    } catch (error) {
      results.push({ id: target.id, label: target.label, status: "failed", error: error instanceof Error ? error.message : String(error) });
    }
  }

  const failed = results.filter(item => item.status === "failed").length;
  if (failed) process.exitCode = 1;
  if (options.json) {
    const path = (id: string) => `/api/chatsvc/${who.region}/v1/users/ME/conversations/${encodeConversation(id)}/properties?name=consumptionhorizon`;
    console.log(JSON.stringify({
      ok: failed === 0,
      ...(options.allowWrite ? {} : { dry_run: true, effect: "read-state", method: "PUT" }),
      conversations: results.map(item => ({
        id: item.id,
        label: item.label,
        status: item.status,
        ...(item.horizon ? { path: path(item.id), body: { consumptionhorizon: item.horizon } } : {}),
        ...(item.error ? { error: item.error } : {}),
      })),
    }, null, 2));
    return;
  }
  if (!options.allowWrite) console.log(`Dry run (effect: read-state): would mark ${results.filter(item => item.status === "would mark").length} conversation(s) read:`);
  for (const item of results) {
    const detail = item.status === "failed" ? `failed: ${item.error}` : item.status === "empty" ? "no messages, skipped" : item.status === "marked" ? "marked read" : `horizon ${item.horizon}`;
    console.log(`  ${item.label} (${item.id})  ${detail}`);
  }
  if (options.allowWrite) console.log(`Marked ${results.filter(item => item.status === "marked").length}, skipped ${results.filter(item => item.status === "empty").length}, failed ${failed}.`);
}
