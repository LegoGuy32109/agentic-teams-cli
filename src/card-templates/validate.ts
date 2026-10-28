import { findEmoji } from "../emoji.ts";
import { resolveReactionKey } from "../reactions.ts";

/** A field of the template data is wrong. The message names the field by its path, such as `actions[1].owner`. */
export class FieldError extends Error {
  constructor(path: string, problem: string) {
    super(`${path} ${problem}`);
    this.name = "FieldError";
  }
}

export type Rec = Record<string, unknown>;

export const BADGE_STYLES = ["Default", "Accent", "Attention", "Good", "Informative", "Subtle", "Warning"] as const;
export type BadgeStyle = (typeof BADGE_STYLES)[number];
export type StatusBadge = { text: string; style: BadgeStyle };

const isRecord = (value: unknown): value is Rec => !!value && typeof value === "object" && !Array.isArray(value);
const describe = (value: unknown) => (value === null ? "null" : Array.isArray(value) ? "a list" : typeof value === "object" ? "an object" : typeof value);

/** An object. `allowed` lists every key it may have. A key that is not in the list is a typo, and it fails. */
export function object(value: unknown, path: string, allowed: readonly string[]): Rec {
  if (!isRecord(value)) throw new FieldError(path, `must be an object, not ${describe(value)}.`);
  for (const key of Object.keys(value)) {
    if (!allowed.includes(key)) throw new FieldError(path === "data" ? `"${key}"` : `${path}.${key}`, `is not a field here. The fields are: ${allowed.join(", ")}.`);
  }
  return value;
}

export function str(value: unknown, path: string): string {
  if (typeof value !== "string") throw new FieldError(path, value === undefined ? "is required." : `must be a string, not ${describe(value)}.`);
  if (!value.trim()) throw new FieldError(path, "must not be empty.");
  return value.trim();
}

export function optStr(value: unknown, path: string): string | undefined {
  return value === undefined ? undefined : str(value, path);
}

export function int(value: unknown, path: string, min: number, max = Number.MAX_SAFE_INTEGER): number {
  if (typeof value !== "number" || !Number.isInteger(value)) throw new FieldError(path, value === undefined ? "is required." : "must be a whole number.");
  if (value < min || value > max) throw new FieldError(path, `must be from ${min} to ${max}.`);
  return value;
}

export function optInt(value: unknown, path: string, min: number, max?: number): number | undefined {
  return value === undefined ? undefined : int(value, path, min, max);
}

export function bool(value: unknown, path: string): boolean {
  if (value === undefined) return false;
  if (typeof value !== "boolean") throw new FieldError(path, "must be true or false.");
  return value;
}

/** A list. Each item goes through `each`, which gets the item and its path. A missing list is empty unless `required`. */
export function list<T>(value: unknown, path: string, each: (item: unknown, path: string) => T, options: { required?: boolean; min?: number; max?: number } = {}): T[] {
  if (value === undefined) {
    if (options.required || options.min) throw new FieldError(path, "is required.");
    return [];
  }
  if (!Array.isArray(value)) throw new FieldError(path, `must be a list, not ${describe(value)}.`);
  if (options.min && value.length < options.min) throw new FieldError(path, `needs at least ${options.min} item${options.min === 1 ? "" : "s"}.`);
  if (options.max && value.length > options.max) throw new FieldError(path, `has ${value.length} items. The most is ${options.max}.`);
  return value.map((item, index) => each(item, `${path}[${index}]`));
}

export const strings = (value: unknown, path: string, options?: { required?: boolean; min?: number; max?: number }) => list(value, path, str, options);

/** A badge style, in any letter case. */
export function style(value: unknown, path: string): BadgeStyle {
  const text = str(value, path);
  const found = BADGE_STYLES.find(item => item.toLowerCase() === text.toLowerCase());
  if (!found) throw new FieldError(path, `"${text}" is not a badge style. Use one of: ${BADGE_STYLES.join(", ")}.`);
  return found;
}

export function status(value: unknown, path: string): StatusBadge {
  const rec = object(value, path, ["text", "style"]);
  return { text: str(rec.text, `${path}.text`), style: style(rec.style, `${path}.style`) };
}

export function oneOf<T extends string>(value: unknown, path: string, allowed: readonly T[]): T {
  const text = str(value, path);
  const found = allowed.find(item => item === text);
  if (!found) throw new FieldError(path, `"${text}" is not allowed. Use one of: ${allowed.join(", ")}.`);
  return found;
}

/** An http or https URL. */
export function url(value: unknown, path: string): string {
  const text = str(value, path);
  let parsed: URL | undefined;
  try { parsed = new URL(text); } catch { parsed = undefined; }
  if (!parsed || (parsed.protocol !== "https:" && parsed.protocol !== "http:")) throw new FieldError(path, `must be an http or https URL, not "${text}".`);
  return text;
}

export const optUrl = (value: unknown, path: string) => (value === undefined ? undefined : url(value, path));

/** A link as a URL, or as `{title, url}`. */
export type Link = { title: string; url: string };
export function link(value: unknown, path: string, defaultTitle: string): Link | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "string") return { title: defaultTitle, url: url(value, path) };
  const rec = object(value, path, ["title", "url"]);
  return { title: optStr(rec.title, `${path}.title`) ?? defaultTitle, url: url(rec.url, `${path}.url`) };
}

/** An emoji that Teams takes as a reaction. A character that is not in the catalog fails here, not after the post. */
export function reactionEmoji(value: unknown, path: string): string {
  const text = str(value, path);
  // A name such as "launch" also resolves, but a vote shows its emoji in the card, so it must be the character.
  if (!/[^\x00-\x7F]/.test(text) || !findEmoji(text)) {
    throw new FieldError(path, `"${text}" is not an emoji in the Teams catalog, so it cannot be a vote. Find one with: teams-cli emoji find <word>.`);
  }
  try {
    resolveReactionKey(text);
  } catch {
    throw new FieldError(path, `"${text}" has no Teams reaction key.`);
  }
  return text;
}
