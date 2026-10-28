import { EMOJI_COUNT, findEmoji, searchEmoji } from "../emoji.ts";

/** Searches the bundled Teams emoji catalog. It reads no network data. */
export function emojiFind(query: string, json: boolean): void {
  const exact = findEmoji(query);
  const found = searchEmoji(query);
  const rows = found.map(entry => ({ key: entry.id, emoji: entry.unicode, name: entry.description, category: entry.category, shortcuts: entry.shortcuts, skin_tones: entry.diverse }));
  // A toned key or toned emoji is an exact match that the search list holds only as its base.
  if (exact?.tone && !rows.some(row => row.key === exact.id)) {
    rows.unshift({ key: exact.id, emoji: exact.unicode, name: exact.title, category: exact.base.category, shortcuts: [], skin_tones: true });
  }
  if (json) {
    console.log(JSON.stringify({ query, catalog_size: EMOJI_COUNT, results: rows }, null, 2));
    return;
  }
  if (!rows.length) {
    console.log(`No emoji match "${query}".`);
    return;
  }
  for (const row of rows) console.log(`${row.emoji}  ${row.key}  ${row.name}${row.skin_tones ? "  (skin tones: -tone1 light, -tone2 medium-light, -tone3 medium, -tone4 medium-dark, -tone5 dark)" : ""}`);
}
