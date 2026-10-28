import { request } from "../transport.ts";
import { encodeConversation, parseMessageRef } from "../refs.ts";
import { resolveConversation } from "../conversations.ts";

export const CATALOG = [
  { name: "self", service: "chat", method: "GET", path: "/users/ME/properties", argument: "none" },
  { name: "teams", service: "csa", method: "GET", path: "/v3/teams/users/me", argument: "none" },
  { name: "conversations", service: "chat", method: "GET", path: "/users/ME/conversations", argument: "none" },
  { name: "members", service: "chat", method: "GET", path: "/threads/<conversation>", argument: "conversation" },
  { name: "message", service: "chat", method: "GET", path: "/users/ME/conversations/<conversation>/messages/<message>", argument: "message-ref" },
] as const;

/** Calls only named, verified read operations. Arbitrary URLs and methods are refused. */
export async function api(name: string, argument: string | undefined, json: boolean): Promise<void> {
  if (name === "list") {
    if (argument) throw new Error("api list takes no argument.");
    if (json) console.log(JSON.stringify({ operations: CATALOG }, null, 2));
    else for (const operation of CATALOG) console.log(`${operation.name}: ${operation.method} ${operation.path} (${operation.argument})`);
    return;
  }
  const operation = CATALOG.find(item => item.name === name);
  if (!operation) throw new Error("Unknown API operation. Run: teams-cli api list");
  if ((operation.argument === "none") !== (argument === undefined)) throw new Error(`api ${name} requires ${operation.argument === "none" ? "no argument" : operation.argument}.`);
  let path: string = operation.path;
  let query: Record<string, string> = {};
  if (name === "teams") query = { isPrefetch: "false", enableMembershipSummary: "true" };
  if (name === "conversations") query = { view: "msnp24Equivalent", pageSize: "200", startTime: "1" };
  if (name === "members") {
    const conversation = await resolveConversation(argument!);
    path = `/threads/${encodeConversation(conversation.id)}`;
    query = { view: "msnp24Equivalent" };
  }
  if (name === "message") {
    const ref = parseMessageRef(argument!);
    if (!ref) throw new Error("Give a message ref: <conversation-id>#<message-id>.");
    path = `/users/ME/conversations/${encodeConversation(ref.conversationId)}/messages/${ref.messageId}`;
  }
  const result = await request(operation.service, "GET", path, { query });
  console.log(JSON.stringify(result ?? null, null, 2));
}
