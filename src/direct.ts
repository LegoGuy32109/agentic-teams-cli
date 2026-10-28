import { request } from "./transport.ts";

/** The 1:1 chat id that Teams derives from the two people: their object ids in ascending order. */
export function directChatId(firstMri: string, secondMri: string): string {
  const [a, b] = [firstMri, secondMri].map(mri => mri.replace(/^8:orgid:/i, "").toLowerCase()).sort();
  return `19:${a}_${b}@unq.gbl.spaces`;
}

/**
 * Starts the 1:1 chat with a person, the way the Teams web client does (captured 2026-10-16): one `POST /threads` with
 * both people as admins and the unique-roster flag, which is lower case. A new chat comes back with its id in
 * `threadResource.id`. For a pair that already has a chat the service answers `{}` with 200, so a repeat call is safe and
 * the id is the one that both object ids give (verified live, 2026-10-16).
 */
export async function createDirectChat(myMri: string, theirMri: string): Promise<string> {
  const response = await request("chat", "POST", "/threads", {
    body: {
      members: [{ id: myMri, role: "Admin" }, { id: theirMri, role: "Admin" }],
      properties: { threadType: "chat", fixedRoster: true, uniquerosterthread: true },
    },
    send: true,
  });
  const id = response?.threadResource?.id;
  if (id === undefined) return directChatId(myMri, theirMri);
  if (typeof id !== "string" || !id.endsWith("@unq.gbl.spaces")) throw new Error("Teams did not return the new chat. Open the chat in Teams once, then try again.");
  return id;
}
