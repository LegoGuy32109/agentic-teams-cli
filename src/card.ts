import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { cardBaseDir, inlineCardImages, type CardImage } from "./card-image.ts";
import { loadProfile } from "./config.ts";
import { MAX_CARD_BYTES, MAX_TEAMS_CARD_BYTES } from "./limits.ts";
import { buildTemplateCard } from "./card-templates/templates.ts";
import { redact } from "./tokens.ts";

export const CARD_CONTENT_TYPE = "application/vnd.microsoft.card.adaptive";
export const ROUTER_VARIABLE = "TEAMS_CARD_ROUTER";
const ROUTER_HOST_SUFFIXES = [".powerplatform.com", ".logic.azure.com"];
const TIMEOUT_MS = 30_000;

export type CardMode = "draft" | "post" | "reply" | "update" | "update-draft";
export type CardInfo = { title: string | null; elements: Record<string, number> };

function isObject(value: unknown): value is Record<string, any> {
  return !!value && typeof value === "object" && !Array.isArray(value);
}

/** Visits every typed node below a value, in document order. */
function walk(value: unknown, visit: (node: Record<string, any>) => void): void {
  if (Array.isArray(value)) {
    for (const item of value) walk(item, visit);
  } else if (isObject(value)) {
    if (typeof value.type === "string") visit(value);
    for (const child of Object.values(value)) walk(child, visit);
  }
}

/** The card's own typed nodes: its body, then its top-level actions. The card itself is not one. */
function walkCard(card: unknown, visit: (node: Record<string, any>) => void): void {
  if (!isObject(card)) return;
  walk(card.body, visit);
  walk(card.actions, visit);
}

/** The text of the first TextBlock, on one line and at most 80 characters. */
export function cardTitle(card: unknown): string | null {
  let title: string | null = null;
  walkCard(card, node => {
    if (title === null && node.type === "TextBlock" && typeof node.text === "string" && node.text.trim()) {
      const flat = node.text.replace(/\s+/g, " ").trim();
      title = flat.length > 80 ? `${flat.slice(0, 77)}...` : flat;
    }
  });
  return title;
}

/** How many nodes of each type the card has. */
export function elementCounts(card: unknown): Record<string, number> {
  const counts: Record<string, number> = {};
  walkCard(card, node => { counts[node.type] = (counts[node.type] ?? 0) + 1; });
  return counts;
}

export function actionTitles(card: unknown): string[] {
  const titles: string[] = [];
  walkCard(card, node => {
    if (node.type.startsWith("Action.") && typeof node.title === "string" && node.title.trim()) titles.push(node.title.trim());
  });
  return titles;
}

export function cardInfo(card: unknown): CardInfo {
  return { title: cardTitle(card), elements: elementCounts(card) };
}

/** The marker that `history` and `context` print for a card message. */
export function cardMarker(card: CardInfo): string {
  return card.title ? `[card: ${card.title}]` : "[card]";
}

/** Reads `properties.cards` of a Teams message. It is a JSON string that holds a list of card entries. */
export function cardFromProperties(value: unknown): CardInfo | undefined {
  let parsed: unknown = value;
  if (typeof value === "string") {
    try { parsed = JSON.parse(value); } catch { return undefined; }
  }
  if (!Array.isArray(parsed)) return undefined;
  const entry = parsed.find(item => isObject(item) && isObject(item.content));
  return entry ? cardInfo(entry.content) : undefined;
}

/** The first card in a message, from `properties.cards` or a Workflows bot `<Swift>` payload. */
export function cardOfMessage(message: unknown): CardInfo | undefined {
  const first = messageCards(message)[0];
  return first ? cardInfo(first) : undefined;
}

/** The outline lines for the preview. */
export function cardOutline(card: unknown): string[] {
  const counts = Object.entries(elementCounts(card)).map(([type, count]) => `${type} ${count}`);
  const actions = actionTitles(card);
  return [
    `title: ${cardTitle(card) ?? "(no TextBlock)"}`,
    `elements: ${counts.length ? counts.join(", ") : "none"}`,
    `actions: ${actions.length ? actions.join(", ") : "none"}`,
  ];
}

export type LoadedCard = {
  card: Record<string, any>;
  /** Bytes of the card as JSON. */
  bytes: number;
  /** SHA-256 of the card as JSON. It goes into the send-once key. */
  digest: string;
  /** SHA-256 of the card with its keys sorted. It does not depend on key order, so a stored card can be matched. */
  hash: string;
  /** The `file:` images that were put into the card as data URIs. */
  images: CardImage[];
};

/** JSON with the keys of every object in sorted order. */
export function canonicalJson(value: unknown): string {
  const sort = (node: unknown): unknown => {
    if (Array.isArray(node)) return node.map(sort);
    if (isObject(node)) return Object.fromEntries(Object.keys(node).sort().map(key => [key, sort(node[key])]));
    return node;
  };
  return JSON.stringify(sort(value));
}

/** The card hash. `card send --wait` matches a posted message with it, and `card update` keys the ledger with it. */
export function cardHash(card: unknown): string {
  return createHash("sha256").update(canonicalJson(card)).digest("hex");
}

/** The cards in `properties.cards` of a Teams message. It is a JSON string of `[{cardClientId, contentType, content}]`. */
export function cardContents(value: unknown): unknown[] {
  let parsed: unknown = value;
  if (typeof value === "string") {
    try { parsed = JSON.parse(value); } catch { return []; }
  }
  if (!Array.isArray(parsed)) return [];
  return parsed.filter(item => isObject(item) && isObject(item.content)).map(item => (item as Record<string, any>).content);
}

/**
 * Every card in a Teams message. A card the user posted is in `properties.cards`. A card the Workflows bot posted is in
 * `content` as `<Swift b64="...">`, base64 JSON of `{attachments: [{content}]}`. Teams rewrites both on the way in (enum
 * case, key order, added nulls), so compare cards by title, never by hash.
 */
export function messageCards(message: any): unknown[] {
  const cards = cardContents(message?.properties?.cards);
  const swift = /<Swift\b[^>]*\bb64="([^"]+)"/i.exec(String(message?.content ?? ""))?.[1];
  if (swift) {
    try {
      const decoded = JSON.parse(Buffer.from(swift, "base64").toString("utf8"));
      for (const item of Array.isArray(decoded?.attachments) ? decoded.attachments : []) if (isObject(item?.content)) cards.push(item.content);
    } catch { /* Not a card payload. */ }
  }
  return cards;
}

function checkCard(card: unknown): Record<string, any> {
  if (!isObject(card)) throw new Error("The card must be a JSON object.");
  if (card.type !== "AdaptiveCard") throw new Error('The card needs "type": "AdaptiveCard".');
  if (typeof card.version !== "string" || !card.version.trim()) throw new Error('The card needs a "version", for example "1.5".');
  if (!Array.isArray(card.body)) throw new Error('The card needs a "body" that is an array.');
  return card;
}

/** Reads a card file. It holds an AdaptiveCard, or a `{type: "message", attachments: [...]}` payload with one card. */
export async function loadCardFile(path: string): Promise<LoadedCard> {
  let text: string;
  try {
    text = await readFile(path, "utf8");
  } catch {
    throw new Error(`Cannot read the card file ${path}.`);
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch {
    throw new Error(`The card file ${path} is not valid JSON.`);
  }
  let card: unknown = parsed;
  if (isObject(parsed) && parsed.type === "message") {
    const attachments = parsed.attachments;
    if (!Array.isArray(attachments) || attachments.length !== 1) throw new Error("A message payload must hold exactly one attachment, the card.");
    card = attachments[0]?.content;
  }
  const inlined = await inlineCardImages(checkCard(card), cardBaseDir(path));
  return finishCard(inlined.card, inlined.images);
}

function finishCard(card: Record<string, any>, images: CardImage[]): LoadedCard {
  const json = JSON.stringify(card);
  return { card, bytes: Buffer.byteLength(json), digest: createHash("sha256").update(json).digest("hex"), hash: cardHash(card), images };
}

/** A card that a template built, and the emoji that people vote with, when the template asks for a vote. */
export type LoadedTemplateCard = LoadedCard & { template: string; votes: string[] };

/** Reads the data file of a template and builds the card. The data and the card are checked before anything is sent. */
export async function loadTemplateCard(name: string, dataFile: string): Promise<LoadedTemplateCard> {
  let text: string;
  try {
    text = await readFile(dataFile, "utf8");
  } catch {
    throw new Error(`Cannot read the data file ${dataFile}.`);
  }
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch {
    throw new Error(`The data file ${dataFile} is not valid JSON.`);
  }
  const built = buildTemplateCard(name, data);
  return { ...finishCard(checkCard(built.card), []), template: built.template, votes: built.votes };
}

export type CardTarget = { mode: CardMode; channel?: string; replyTo?: string; messageId?: string };

export function cardPayload(card: Record<string, any>, target: CardTarget): Record<string, unknown> {
  return {
    type: "message",
    target: {
      mode: target.mode,
      ...(target.mode === "draft" || target.mode === "update-draft" ? {} : { channel: target.channel }),
      ...(target.mode === "reply" ? { replyTo: target.replyTo } : {}),
      ...(target.mode === "update" || target.mode === "update-draft" ? { messageId: target.messageId } : {}),
    },
    attachments: [{ contentType: CARD_CONTENT_TYPE, contentUrl: null, content: card }],
  };
}

/** Fails when the payload is larger than the flow accepts. Returns the size. */
export function checkPayloadSize(payload: unknown): number {
  const bytes = Buffer.byteLength(JSON.stringify(payload));
  if (bytes > MAX_CARD_BYTES) {
    throw new Error(`The payload is ${bytes} bytes. The card router accepts ${MAX_CARD_BYTES} bytes at most (${MAX_CARD_BYTES / 1024} KB). Make the card smaller.`);
  }
  // Teams has its own, smaller limit for the card itself. It refuses the post after the flow has accepted it.
  const attachments = (payload as { attachments?: Array<{ content?: unknown }> })?.attachments;
  const cardSize = attachments?.[0]?.content === undefined ? 0 : Buffer.byteLength(JSON.stringify(attachments[0].content));
  if (cardSize > MAX_TEAMS_CARD_BYTES) {
    throw new Error(`The card is ${cardSize} bytes. Teams refuses a card over ${MAX_TEAMS_CARD_BYTES} bytes (${MAX_TEAMS_CARD_BYTES / 1024} KB): "The payload is too large. Please make sure the size is less than 28KB." Make the card smaller.`);
  }
  return bytes;
}

/** The URL must be https on a Power Automate host. The error never shows the value. */
export function checkRouterUrl(raw: string, source: string): URL {
  let url: URL | undefined;
  try { url = new URL(raw); } catch { url = undefined; }
  if (!url || url.protocol !== "https:" || !ROUTER_HOST_SUFFIXES.some(suffix => url.hostname.endsWith(suffix))) {
    throw new Error(`${source} is not accepted. It must be an https URL on a host that ends in ${ROUTER_HOST_SUFFIXES.join(" or ")}.`);
  }
  return url;
}

/** Reads the router URL: `TEAMS_CARD_ROUTER` first, then the URL that `card setup` stored in the profile. */
export async function routerUrl(): Promise<URL> {
  const raw = process.env[ROUTER_VARIABLE]?.trim();
  if (raw) return checkRouterUrl(raw, ROUTER_VARIABLE);
  const stored = (await loadProfile()).profile?.cardRouterUrl;
  if (stored) return checkRouterUrl(stored, "The stored card router URL");
  throw new Error(`${ROUTER_VARIABLE} is not set, and no router URL is stored. It is the webhook URL of the Workflows "teams-cli card router" flow. Set it in the environment, or run: teams-cli card setup --allow-write`);
}

/** The result of one POST to the router. `uncertain` means the card may still post: a network error, a 408, or a 5xx. */
export type RouterResult = { status?: number; body: string; uncertain: boolean };

/** Posts the payload once. It never retries, and nothing it returns holds the URL. */
export async function postToRouter(url: URL, payload: unknown): Promise<RouterResult> {
  let response: Response;
  try {
    response = await fetch(url, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
      redirect: "error",
      signal: AbortSignal.timeout(TIMEOUT_MS),
    });
  } catch {
    return { body: "", uncertain: true };
  }
  const text = await response.text().catch(() => "");
  const body = redact(text).replace(/\s+/g, " ").trim().slice(0, 200);
  return { status: response.status, body, uncertain: response.status === 408 || response.status >= 500 };
}
