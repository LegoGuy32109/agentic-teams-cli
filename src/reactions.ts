import { findEmoji } from "./emoji.ts";

/** Emoji whose quick-reaction key differs from the first catalog key. Captured 2026-10-19. */
const ALIASES: Record<string, string> = {
  "👍": "like",
  "❤": "heart",
  "😆": "laugh",
  "😂": "laugh",
  "😮": "surprised",
  "😢": "sad",
  "🙁": "sad",
  "😠": "angry",
  "✅": "2705_whiteheavycheckmark",
};

/** A raw key: a named key, a codepoint key (`2705_name`), or a custom org emoji (`name;id`). */
const RAW_KEY = /^[A-Za-z0-9_-]+(;[A-Za-z0-9_-]+)?$/;

/**
 * Turns the emoji argument into the key Teams stores. The quick reactions keep their captured keys. Any other
 * emoji character, with or without a skin tone, maps through the bundled Teams catalog, as does a key or shortcut
 * such as `launch`, `(launch)`, or `:launch:`. A raw key that is not in the catalog passes unchanged.
 */
export function resolveReactionKey(input: string): string {
  const trimmed = input.trim();
  // Emoji come with or without the variation selector (U+FE0F). Both mean the same emoji.
  const emoji = findEmoji(trimmed);
  // A skin tone, given or preferred, takes the toned catalog key over the toneless quick-reaction key.
  if (emoji?.tone) return emoji.id;
  const alias = ALIASES[trimmed.replaceAll("️", "")];
  if (alias) return alias;
  if (emoji) return emoji.id;
  if (RAW_KEY.test(trimmed)) return trimmed;
  throw new Error(`"${trimmed}" is not in the Teams emoji catalog. Find one with: teams-cli emoji find <word>. A custom org emoji is name;id.`);
}
