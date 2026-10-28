import { chmod, mkdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

export const DEFAULT_PROFILE = "default";

export type TokenKind = "chat" | "csa" | "spaces" | "substrate" | "sharepoint";

export type RefreshConfig = {
  clientId: string;
  tenantId: string;
  refreshToken: string;
  origin?: string;
  scopes?: Partial<Record<TokenKind, string>>;
};

export type StoredToken = { token: string; exp: number };

/** When the dedicated browser profile last produced tokens, and how. */
export type BrowserMeta = { lastAt: string; lastMethod: "login" | "renew" };

export type Profile = {
  refresh?: RefreshConfig;
  browser?: BrowserMeta;
  /** The host of the personal SharePoint site, such as `contoso-my.sharepoint.com`. Found once, then kept. */
  sharepointHost?: string;
  /** The webhook URL of the card router flow. It is a secret, so it stays in this 0600 store and is never printed. */
  cardRouterUrl?: string;
  tokens: Partial<Record<TokenKind, StoredToken>>;
  userId: string;
  tenantId: string;
  upn: string;
  name: string;
  region: string;
  updatedAt: string;
};

type CredentialStore = {
  version: 1;
  activeProfile?: string;
  profiles: Record<string, Profile>;
};

let selectedProfile: string | undefined;

/** Sets the profile named by `--profile`. Commands read it through `loadProfile()`. */
export function selectProfile(name: string | undefined): void {
  selectedProfile = name;
}

/**
 * Sets a file mode. Windows has no POSIX modes, so `chmod` does nothing there. Skip it, and rely on the per-user
 * `%APPDATA%` folder. The platform is a parameter so tests can cover both branches.
 */
export async function restrictMode(path: string, mode: number, platform: NodeJS.Platform = process.platform): Promise<void> {
  if (platform === "win32") return;
  await chmod(path, mode);
}

export function configDir(): string {
  if (process.platform === "darwin") return join(homedir(), "Library", "Application Support", "teams-cli");
  if (process.platform === "win32") return join(process.env.APPDATA || join(homedir(), "AppData", "Roaming"), "teams-cli");
  return join(process.env.XDG_CONFIG_HOME || join(homedir(), ".config"), "teams-cli");
}

export function cacheDir(): string {
  if (process.platform === "darwin") return join(homedir(), "Library", "Caches", "teams-cli");
  if (process.platform === "win32") return join(process.env.LOCALAPPDATA || process.env.APPDATA || join(homedir(), "AppData", "Local"), "teams-cli", "Cache");
  return join(process.env.XDG_CACHE_HOME || join(homedir(), ".cache"), "teams-cli");
}

export function credentialsPath(): string {
  return join(configDir(), "credentials.json");
}

/** Writes through a temp file and a rename, so a crash never leaves half a file. */
export async function writeFileAtomic(path: string, content: string, mode: number): Promise<void> {
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const temp = `${path}.${process.pid}.${Date.now()}.${Math.random().toString(36).slice(2)}.tmp`;
  try {
    await writeFile(temp, content, { mode });
    await restrictMode(temp, mode).catch(() => {});
    await rename(temp, path);
    await restrictMode(path, mode).catch(() => {});
  } finally {
    await rm(temp, { force: true }).catch(() => {});
  }
}

async function readStore(): Promise<CredentialStore> {
  try {
    const parsed = JSON.parse(await readFile(credentialsPath(), "utf8"));
    if (parsed?.version === 1 && parsed?.profiles && typeof parsed.profiles === "object") return parsed;
  } catch { /* First run or unreadable file. */ }
  return { version: 1, profiles: {} };
}

async function writeStore(store: CredentialStore): Promise<void> {
  await writeFileAtomic(credentialsPath(), `${JSON.stringify(store, null, 2)}\n`, 0o600);
}

/** Returns the selected profile name and its stored data, if any. */
export async function loadProfile(name = selectedProfile): Promise<{ name: string; profile?: Profile }> {
  const store = await readStore();
  const resolved = name || store.activeProfile || DEFAULT_PROFILE;
  return { name: resolved, profile: store.profiles[resolved] };
}

export async function saveProfile(name: string, profile: Profile): Promise<void> {
  const store = await readStore();
  store.profiles[name] = profile;
  store.activeProfile = name;
  await writeStore(store);
}

export async function removeProfile(name = selectedProfile): Promise<boolean> {
  const store = await readStore();
  const selected = name || store.activeProfile;
  if (!selected || !store.profiles[selected]) return false;
  delete store.profiles[selected];
  if (store.activeProfile === selected) store.activeProfile = Object.keys(store.profiles)[0];
  await writeStore(store);
  return true;
}

/** Keeps the card router webhook URL in the selected profile. The profile must exist already. */
export async function saveCardRouterUrl(url: string): Promise<void> {
  const { name, profile } = await loadProfile();
  if (!profile) throw new Error("No stored profile exists, so the router URL cannot be kept. Run: teams-cli auth");
  await saveProfile(name, { ...profile, cardRouterUrl: url });
}
