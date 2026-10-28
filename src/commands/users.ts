import { searchPeople } from "../users.ts";

/** Finds users by name, mail, or upn. Prints one ID per user. */
export async function usersFind(query: string, json: boolean, refresh = false): Promise<void> {
  const users = await searchPeople(query, { refresh });
  if (json) {
    console.log(JSON.stringify(users.map(user => ({
      id: user.mri,
      handle: user.upn ?? user.mail ?? "",
      display_name: user.displayName,
      real_name: user.realName ?? "",
      active: null,
    })), null, 2));
    return;
  }
  if (!users.length) console.log("No users found.");
  for (const user of users) console.log(`${user.mri}\t${user.displayName}\t@${user.upn ?? user.mail ?? ""}`);
}
