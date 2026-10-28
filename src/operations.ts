/** The preview and receipt shapes for every write. Keep them the same across commands. */
export type PreviewResult = {
  ok: true;
  dry_run: true;
  effect: "message" | "scheduled_message";
  /** Scheduled messages only: the send time in UTC and in the local zone. */
  send_at?: string;
  send_at_local?: string;
  method: "POST";
  path: string;
  conversation: { id: string; label: string };
  /** The 1:1 chat does not exist yet. A real send creates it first. */
  new_chat?: boolean;
  body: unknown;
  /** The files that a real send would upload first. A preview uploads nothing. */
  images?: Array<{ file: string; bytes: number; type: string; width: number; height: number; display: { width: number; height: number } }>;
  /** The files that a real send would upload, where they go, and who it would invite to open them. A preview uploads and invites no one. */
  files?: Array<{ file: string; bytes: number; destination: string; conflict?: string }>;
  channel_folder?: string;
  sharing?: { access: string; role: string };
  invite?: { roles: string[]; default_link?: boolean; recipients: Array<{ name: string; object_id: string }>; skipped: Array<{ mri: string; reason: string }> };
  warning?: string;
};

export type ReceiptResult = {
  ok: true;
  effect: "message";
  conversation: string;
  id: string;
  ts: string;
  created: string;
  permalink: string;
  reply_to?: string;
  warning?: string;
  verification?: { ok: boolean; expected: string; actual: string };
};

export function permalinkFor(conversationId: string, ts: string): string {
  return `https://teams.cloud.microsoft/l/message/${encodeURIComponent(conversationId)}/${ts}`;
}
