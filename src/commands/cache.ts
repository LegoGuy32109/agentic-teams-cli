import { clearUserCache } from "../users.ts";
import { identity } from "../tokens.ts";

/** Clears the user-name cache. Teams has no verified bulk user list, so the cache refills on demand. */
export async function cacheUsers(): Promise<void> {
  const who = await identity();
  console.log(`Cleared ${await clearUserCache(who.tenantId)} cached users. Names refill on demand.`);
}
