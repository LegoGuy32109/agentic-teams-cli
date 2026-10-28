import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { configDir, writeFileAtomic } from "./config.ts";

/** Preferences that are not credentials. They apply to every profile. */
export type Settings = { skinTone?: number; fileAccess?: FileAccess; fileRole?: FileRole };

/**
 * Who can open a file sent to a chat. `link` is what the Teams web client does: invite the members and ask for the
 * tenant's default sharing link (the tenant may allow anyone with the link to edit). `people` invites the members
 * with no link. `private` invites nobody and makes no link, so only you can open it. Channels are not affected.
 */
export type FileAccess = "link" | "people" | "private";
/** The access an invited member gets. The web client gives `edit`. */
export type FileRole = "edit" | "view";
export const FILE_ACCESS: FileAccess[] = ["link", "people", "private"];
export const FILE_ROLES: FileRole[] = ["edit", "view"];

export function parseFileAccess(value: string): FileAccess {
  const word = value.trim().toLowerCase() as FileAccess;
  if (!FILE_ACCESS.includes(word)) throw new Error(`"${value}" is not a file access. Use link, people, or private.`);
  return word;
}

export function parseFileRole(value: string): FileRole {
  const word = value.trim().toLowerCase() as FileRole;
  if (!FILE_ROLES.includes(word)) throw new Error(`"${value}" is not a file role. Use edit or view.`);
  return word;
}

/** The flag wins, then `TEAMS_FILE_ACCESS` and `TEAMS_FILE_ROLE`, then the saved setting, then the web client's choice. */
export async function fileSharing(flags: { access?: string; role?: string }): Promise<{ access: FileAccess; role: FileRole }> {
  const saved = await loadSettings();
  const access = flags.access ?? process.env.TEAMS_FILE_ACCESS ?? saved.fileAccess ?? "link";
  const role = flags.role ?? process.env.TEAMS_FILE_ROLE ?? saved.fileRole ?? "edit";
  return { access: parseFileAccess(access), role: parseFileRole(role) };
}

/** `config file-access` and `config file-role` print the setting. With a value, they save it. `default` clears it. */
export async function configFileSetting(name: "file-access" | "file-role", value: string | undefined): Promise<void> {
  const settings = await loadSettings();
  const key = name === "file-access" ? "fileAccess" : "fileRole";
  if (value === undefined) {
    const current = await fileSharing({});
    console.log(`${name}: ${name === "file-access" ? current.access : current.role}${settings[key] ? "" : " (default)"}`);
    return;
  }
  if (value.trim().toLowerCase() === "default") delete settings[key];
  else if (key === "fileAccess") settings.fileAccess = parseFileAccess(value);
  else settings.fileRole = parseFileRole(value);
  await saveSettings(settings);
  console.log(settings[key] ? `${name} set to ${settings[key]}.` : `${name} cleared. The default applies.`);
}

export function settingsPath(): string {
  return join(configDir(), "settings.json");
}

export async function loadSettings(): Promise<Settings> {
  try {
    const parsed = JSON.parse(await readFile(settingsPath(), "utf8"));
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

export async function saveSettings(settings: Settings): Promise<void> {
  await writeFileAtomic(settingsPath(), `${JSON.stringify(settings, null, 2)}\n`, 0o600);
}

/** Tone names as Teams titles them, plus `white` and `black` as people say them. `none` clears the preference. */
const TONE_WORDS: Record<string, number> = {
  light: 1, white: 1, "medium-light": 2, medium: 3, "medium-dark": 4, dark: 5, black: 5,
  "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "tone1": 1, "tone2": 2, "tone3": 3, "tone4": 4, "tone5": 5,
  "🏻": 1, "🏼": 2, "🏽": 3, "🏾": 4, "🏿": 5,
};
export const TONE_LABELS = ["light", "medium-light", "medium", "medium-dark", "dark"];

/** Returns 1 to 5, 0 for `none`, or throws for a word that is not a tone. */
export function parseSkinTone(value: string): number {
  const word = value.trim().toLowerCase();
  if (word === "none" || word === "default" || word === "0") return 0;
  const tone = TONE_WORDS[word];
  if (!tone) throw new Error(`"${value}" is not a skin tone. Use light (white), medium-light, medium, medium-dark, dark, 1 to 5, or none.`);
  return tone;
}

/** `TEAMS_SKIN_TONE` wins over the saved setting. */
export async function preferredSkinTone(): Promise<number | undefined> {
  const fromEnv = process.env.TEAMS_SKIN_TONE;
  if (fromEnv) return parseSkinTone(fromEnv) || undefined;
  const tone = (await loadSettings()).skinTone;
  return typeof tone === "number" && tone >= 1 && tone <= 5 ? tone : undefined;
}

/** `config skin-tone` prints the preference. `config skin-tone <tone>` saves it. */
export async function configSkinTone(value: string | undefined): Promise<void> {
  if (value === undefined) {
    const tone = await preferredSkinTone();
    console.log(tone ? `Skin tone: ${TONE_LABELS[tone - 1]} (-tone${tone})${process.env.TEAMS_SKIN_TONE ? ", from TEAMS_SKIN_TONE" : ""}` : "Skin tone: none (Teams default yellow)");
    return;
  }
  const tone = parseSkinTone(value);
  const settings = await loadSettings();
  if (tone) settings.skinTone = tone; else delete settings.skinTone;
  await saveSettings(settings);
  console.log(tone ? `Skin tone set to ${TONE_LABELS[tone - 1]} (-tone${tone}). Emoji that take tones use it unless you give one.` : "Skin tone cleared.");
}
