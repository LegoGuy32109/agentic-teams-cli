import catalog from "./emoji-catalog.json" with { type: "json" };
import { escapeHtml } from "./html.ts";

/** The type of an emoji `<img>` in message HTML. Captured 2026-10-19. */
export const EMOJI_TYPE = "http://schema.skype.com/Emoji";
const ASSETS = "https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons";

type Row = [id: string, unicode: string, description: string, category: string, shortcuts: string[], keywords: string[], diverse: 0 | 1];

export type Emoji = { id: string; unicode: string; description: string; category: string; shortcuts: string[]; keywords: string[]; diverse: boolean };

/** One resolved emoji. `tone` is 1 to 5 for a skin tone, and then `id` ends with `-toneN`. */
export type ResolvedEmoji = { id: string; unicode: string; title: string; base: Emoji; tone?: number };

const ENTRIES: Emoji[] = (catalog.entries as Row[]).map(([id, unicode, description, category, shortcuts, keywords, diverse]) => ({ id, unicode, description, category, shortcuts, keywords, diverse: diverse === 1 }));

/** U+1F3FB to U+1F3FF, in tone order. Teams names them `-tone1` to `-tone5`. */
const TONES = ["\u{1F3FB}", "\u{1F3FC}", "\u{1F3FD}", "\u{1F3FE}", "\u{1F3FF}"];
const TONE_NAMES = ["light", "medium-light", "medium", "medium-dark", "dark"];

/** The variation selector U+FE0F changes nothing about which emoji it is. */
const bare = (value: string) => value.replaceAll("️", "");

// The first entry wins when two entries share a character, as `like` and `yes` share 👍. The catalog lists Smilies first.
// A skin tone takes the first entry that has tones, so 👍🏽 is `yes-tone3`, as the web client sent it.
const BY_UNICODE = new Map<string, Emoji[]>();
const BY_NAME = new Map<string, Emoji>();
for (const entry of ENTRIES) {
  const key = bare(entry.unicode);
  BY_UNICODE.set(key, [...(BY_UNICODE.get(key) ?? []), entry]);
  for (const name of [entry.id, ...entry.shortcuts]) if (!BY_NAME.has(name.toLowerCase())) BY_NAME.set(name.toLowerCase(), entry);
}

/** The saved skin tone, 1 to 5. It applies to an emoji that takes tones when no tone is given. */
let preferredTone: number | undefined;

export function setPreferredTone(tone: number | undefined): void {
  preferredTone = tone;
}

function withTone(base: Emoji, tone?: number): ResolvedEmoji | undefined {
  if (tone === undefined && preferredTone && base.diverse) tone = preferredTone;
  if (tone === undefined) return { id: base.id, unicode: base.unicode, title: base.description, base };
  if (!base.diverse) return undefined;
  return {
    id: `${base.id}-tone${tone}`,
    unicode: base.unicode + TONES[tone - 1],
    title: `${base.description} with ${TONE_NAMES[tone - 1]} skin tone`,
    base,
    tone,
  };
}

/** Finds an emoji by its character, with or without a skin tone. */
export function emojiByUnicode(value: string): ResolvedEmoji | undefined {
  const key = bare(value);
  const entries = BY_UNICODE.get(key);
  // With a preferred tone, the entry that takes tones wins, so 👍 is `yes-toneN` and not the toneless `like`.
  const exact = (preferredTone ? entries?.find(entry => entry.diverse) : undefined) ?? entries?.[0];
  if (exact) return withTone(exact);
  const tone = TONES.findIndex(mark => key.includes(mark));
  if (tone < 0) return undefined;
  const base = BY_UNICODE.get(key.replace(TONES[tone]!, ""))?.find(entry => entry.diverse);
  return base ? withTone(base, tone + 1) : undefined;
}

/** Finds an emoji by key or shortcut: `launch`, `(launch)`, `:launch:`, or `yes-tone3`. */
export function emojiByName(value: string): ResolvedEmoji | undefined {
  const name = value.trim().replace(/^\((.*)\)$/, "$1").replace(/^:(.*):$/, "$1").toLowerCase();
  const toned = /^(.+)-tone([1-5])$/.exec(name);
  const base = BY_NAME.get(toned ? toned[1]! : name);
  return base ? withTone(base, toned ? Number(toned[2]) : undefined) : undefined;
}

/** Finds an emoji by character, key, or shortcut. */
export function findEmoji(value: string): ResolvedEmoji | undefined {
  return emojiByUnicode(value.trim()) ?? emojiByName(value);
}

/** Searches keys, shortcuts, names, and keywords. Exact key and shortcut matches come first. */
export function searchEmoji(query: string, limit = 25): Emoji[] {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  if (!words.length) return [];
  const direct = findEmoji(query);
  const scored: Array<{ entry: Emoji; score: number }> = [];
  for (const entry of ENTRIES) {
    const haystack = [entry.id, ...entry.shortcuts, entry.description.toLowerCase(), ...entry.keywords.map(word => word.toLowerCase())];
    if (!words.every(word => haystack.some(item => item.includes(word)))) continue;
    const score = entry === direct?.base ? 0 : entry.description.toLowerCase().includes(query.toLowerCase()) ? 1 : 2;
    scored.push({ entry, score });
  }
  return scored.sort((a, b) => a.score - b.score).slice(0, limit).map(item => item.entry);
}

/** The message markup for one emoji, in the shape the web client sent on 2026-10-19. */
export function emojiHtml(emoji: ResolvedEmoji): string {
  // The capture showed `20_f_s4.png` for tone 3, so the image suffix is the tone plus one.
  const src = `${ASSETS}/${emoji.base.id}/default/20_f${emoji.tone ? `_s${emoji.tone + 1}` : ""}.png`;
  const title = escapeHtml(emoji.title);
  const id = escapeHtml(emoji.id);
  return `<span contenteditable="false" title="${title}" type="(${id})" class="animated-emoticon-20-${id}" itemscope>`
    + `<img itemscope itemtype="${EMOJI_TYPE}" itemid="${id}" src="${src}" title="${title}" alt="${emoji.unicode}" style="width:20px;height:20px;"></span>`;
}

// An emoji: a pictograph with an optional variation selector and skin tone, joined by ZWJ to more of the same.
// Keycaps (#️⃣) and flags are not matched and stay as plain characters.
const PICTOGRAPH = "\\p{Extended_Pictographic}\\uFE0F?[\\u{1F3FB}-\\u{1F3FF}]?";
const EMOJI_OR_NAME = new RegExp(`(${PICTOGRAPH}(?:\\u200D${PICTOGRAPH})*)|(?<![\\w:]):([a-z0-9_+-]*[a-z][a-z0-9_+-]*):(?![\\w:])`, "gu");

/**
 * `teams` writes the emoji markup that the Teams picker writes. `plain` keeps the characters, as when an emoji is
 * typed from the keyboard. The web client sent both shapes on 2026-10-19. A command sets this around one compose.
 */
export type EmojiStyle = "teams" | "plain";
let style: EmojiStyle = "teams";

/** Runs `compose` with one emoji style, then restores the default. */
export async function withEmojiStyle<T>(next: EmojiStyle | undefined, compose: () => Promise<T>): Promise<T> {
  const previous = style;
  style = next ?? "teams";
  try { return await compose(); } finally { style = previous; }
}

/**
 * Escapes text for HTML and turns each emoji character, and each `:name:` that names a catalog emoji, into
 * the Teams emoji markup. In the plain style, a character stays as written and a `:name:` becomes its character.
 * Other text, and a `:name:` that is not in the catalog, is escaped unchanged.
 */
export function emojifyText(text: string, escape: (value: string) => string = escapeHtml): string {
  let html = "";
  let last = 0;
  for (const match of text.matchAll(EMOJI_OR_NAME)) {
    if (style === "plain" && match[1]) continue;
    const emoji = match[1] ? emojiByUnicode(match[1]) : emojiByName(match[2]!);
    if (!emoji) continue;
    html += escape(text.slice(last, match.index)) + (style === "plain" ? emoji.unicode : emojiHtml(emoji));
    last = match.index + match[0].length;
  }
  return html + escape(text.slice(last));
}

/** The number of entries in the bundled catalog. */
export const EMOJI_COUNT = ENTRIES.length;
