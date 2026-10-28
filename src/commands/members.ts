import { resolveConversation } from "../conversations.ts";
import { encodeConversation } from "../refs.ts";
import { identity } from "../tokens.ts";
import { request } from "../transport.ts";
import { resolveUsers } from "../users.ts";

export type MemberRow = { mri: string; name: string; email: string; role: string };

/** Lists a conversation's members with names, mail, and role. Names come from the cache or fetchShortProfile. */
export async function membersOf(input: string, json: boolean): Promise<void> {
  const who = await identity();
  const target = await resolveConversation(input);
  const thread: any = await request("chat", "GET", `/threads/${encodeConversation(target.id)}`, { query: { view: "msnp24Equivalent" } });
  const listed = (Array.isArray(thread?.members) ? thread.members : [])
    .map((member: any) => ({ mri: String(member?.id ?? ""), role: String(member?.role ?? "") }))
    .filter((member: { mri: string }) => member.mri);
  const users = await resolveUsers(who.tenantId, listed.map((member: { mri: string }) => member.mri));
  const rows: MemberRow[] = listed.map((member: { mri: string; role: string }) => ({
    mri: member.mri,
    name: users[member.mri]?.displayName ?? member.mri,
    email: users[member.mri]?.mail ?? "",
    role: member.role,
  }));

  if (json) {
    console.log(JSON.stringify(rows, null, 2));
    return;
  }
  if (!rows.length) console.log("No members returned.");
  for (const row of rows) console.log(`${row.mri}\t${row.name}\t${row.email}\t${row.role}`);
}
