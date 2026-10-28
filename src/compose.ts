import { createHash } from "node:crypto";
import { codeBlockHtml } from "./highlight.ts";
import { escapeHtml } from "./html.ts";
import { emojifyText } from "./emoji.ts";
import { decodeEntities, htmlToText } from "./message.ts";
import { badgeColor, badgeTextColor, highlightColor, textColor } from "./palette.ts";
import { mentionTokens, PLACEHOLDER_END, PLACEHOLDER_START, resolveMentionTokens, type MentionScope, type User } from "./users.ts";

export const MENTION_TYPE = "http://schema.skype.com/Mention";
const PLACEHOLDER = new RegExp(`${PLACEHOLDER_START}(\\d+)${PLACEHOLDER_END}`, "g");
export const LINK_TYPE = "http://schema.skype.com/HyperLink";
export const MAX_BODY_BYTES = 100_000;

export type MentionEntry = { "@type": string; itemid: string; mri: string; mentionType: "person" | "everyone" | "channel"; displayName: string };

export { escapeHtml };

/** The browser shape for one mention. The span text is the full display name. */
export function mentionHtml(itemid: string, displayName: string): string {
  return `<readonly class="skipProofing" itemtype="${MENTION_TYPE}" spellcheck="false"><span itemtype="${MENTION_TYPE}" itemscope itemid="${itemid}">${escapeHtml(displayName)}</span></readonly>`;
}

/** One `properties.links` entry. `previewenabled` asks Teams for a link preview card. */
export type LinkEntry = { "@type": string; itemid: string; url: string; previewenabled: boolean; preview?: unknown };

type Collector = { next: number; entries: MentionEntry[]; links: LinkEntry[]; rich?: boolean; open?: string[] };

/** The anchor id is a UUID-shaped hash of the URL, so one message always has the same content. */
function linkId(url: string): string {
  const hex = createHash("sha256").update(url).digest("hex");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}`;
}

function linkHtml(url: string, label: string, collector: Collector): string {
  collector.links.push({ "@type": LINK_TYPE, itemid: String(collector.links.length), url, previewenabled: true });
  const href = escapeHtml(url);
  return `<a href="${href}" itemtype="${LINK_TYPE}" rel="noreferrer noopener" title="${href}" target="_blank" itemid="${linkId(url)}">${escapeHtml(label)}</a>`;
}

// One of: a Slack link `<url>` or `<url|label>`, a bare URL, a Markdown link `[label](url)`, `*bold*` or `**bold**`, a color tag `[red]`, `[hl:red]`, or `[badge:red]`,
// the closing tag `[/]`, `` `inline code` ``, `_italic_`, `~strike~`, or `\[red]` for a literal tag. Bold opens after a space
// or an opening bracket and closes before a space or punctuation, so `2 * 3 * 4`, `snake_case_name` and a `* ` bullet stay text.
const RICH_INLINE = /<(https?:\/\/[^\s|>]+)(?:\|([^>\n]+))?>|(https?:\/\/[^\s<>]*[^\s<>.,;:!?)\]'"])|\[([^\]\n]+)\]\((https?:\/\/[^\s()<>]+)\)|(?<![\w*])(\*{1,2})(?=\S)([^*\n]*[^*\s])\6(?![\w*])|\[(hl:|badge:)?([a-z][a-z-]*)\]|(\[\/\])|`([^`\n]+)`|(?<![\w_])_(?=\S)([^_\n]*[^_\s])_(?![\w_])|(?<![\w~])~(?![/\s~])([^~\n]*[^~\s])~(?![\w~/])|\\(\[(?:hl:|badge:)?[a-z][a-z-]*\])|(<u>)|(<\/u>)/g;

/** Closes a badge. A badge inside a badge is flattened: its tag adds nothing, and its `[/]` pops an empty entry. */
const BADGE_CLOSE = "&nbsp;</b></span>";

/** Rich text between mentions: links and bold. Emoji become the Teams emoji markup, except in code. Everything else is escaped. */
function richText(text: string, collector: Collector): string {
  let html = "";
  let last = 0;
  for (const match of text.matchAll(RICH_INLINE)) {
    html += emojifyText(text.slice(last, match.index));
    if (match[1]) html += linkHtml(match[1], match[2] ?? match[1], collector);
    else if (match[3]) html += linkHtml(match[3], match[3], collector);
    else if (match[5]) html += linkHtml(match[5], match[4]!, collector);
    else if (match[7]) html += `<strong>${emojifyText(match[7])}</strong>`;
    else if (match[10]) {
      const close = collector.open?.pop();
      html += close ?? escapeHtml(match[0]);
    } else if (match[11]) html += `<code>${escapeHtml(match[11])}</code>`;
    else if (match[12]) html += `<i>${emojifyText(match[12])}</i>`;
    else if (match[13]) html += `<s>${emojifyText(match[13])}</s>`;
    else if (match[14]) html += escapeHtml(match[14]);
    else if (match[15] || (match[9] === "u" && !match[8])) { html += "<u>"; (collector.open ??= []).push("</u>"); }
    else if (match[16]) {
      if (collector.open?.at(-1) === "</u>") html += collector.open.pop(); else html += escapeHtml(match[0]);
    }
    else {
      if (match[8] === "badge:") {
        const bg = badgeColor(match[9]!);
        if (!bg) html += escapeHtml(match[0]);
        else if (collector.open?.includes(BADGE_CLOSE)) (collector.open ??= []).push("");
        else { html += `<span style="background-color:${bg};color:${badgeTextColor(bg)}"><b>&nbsp;`; (collector.open ??= []).push(BADGE_CLOSE); }
        last = match.index + match[0].length;
        continue;
      }
      const hex = match[8] ? highlightColor(match[9]!) : textColor(match[9]!);
      if (hex) { html += `<span style="${match[8] ? "background-color" : "color"}:${hex}">`; (collector.open ??= []).push("</span>"); }
      else html += escapeHtml(match[0]);
    }
    last = match.index + match[0].length;
  }
  return html + emojifyText(text.slice(last));
}

// The links of the plain format: a Slack link `<url>` or `<url|label>`, a Markdown link `[label](url)`, or a bare URL.
const PLAIN_LINK = /<(https?:\/\/[^\s|>]+)(?:\|([^>\n]+))?>|\[([^\]\n]+)\]\((https?:\/\/[^\s()<>]+)\)|(https?:\/\/[^\s<>]*[^\s<>.,;:!?)\]'"])/g;

/** Plain text between mentions: links become anchors, emoji become the Teams emoji markup, and everything else is escaped. */
function plainText(text: string, collector: Collector): string {
  let html = "";
  let last = 0;
  for (const match of text.matchAll(PLAIN_LINK)) {
    html += emojifyText(text.slice(last, match.index));
    if (match[1]) html += linkHtml(match[1], match[2] ?? match[1], collector);
    else if (match[4]) html += linkHtml(match[4], match[3]!, collector);
    else html += linkHtml(match[5]!, match[5]!, collector);
    last = match.index + match[0].length;
  }
  return html + emojifyText(text.slice(last));
}

/** Escapes one line and swaps each placeholder for a mention. Item ids count across the whole message. */
function composeInline(line: string, mentions: User[], collector: Collector): string {
  const text = (value: string) => collector.rich ? richText(value, collector) : plainText(value, collector);
  collector.open = [];
  let html = "";
  let last = 0;
  for (const match of line.matchAll(PLACEHOLDER)) {
    html += text(line.slice(last, match.index));
    const user = mentions[Number(match[1])]!;
    const itemid = String(collector.next++);
    collector.entries.push({ "@type": MENTION_TYPE, itemid, mri: user.mri, mentionType: user.mentionType ?? "person", displayName: user.displayName });
    html += mentionHtml(itemid, user.displayName);
    last = match.index + match[0].length;
  }
  html += text(line.slice(last));
  // A color tag that was never closed ends with its line.
  while (collector.open?.length) html += collector.open.pop();
  return html;
}

/** Plain format: each line is a paragraph, and links are anchors. An empty line is `&nbsp;` (unverified, see README). */
export function composePlain(text: string, mentions: User[]): { html: string; entries: MentionEntry[]; links: LinkEntry[] } {
  const collector: Collector = { next: 0, entries: [], links: [] };
  const html = text.split("\n")
    .map(line => line === "" ? "<p>&nbsp;</p>" : `<p>${composeInline(line, mentions, collector)}</p>`)
    .join("");
  return { html, entries: collector.entries, links: collector.links };
}

export type CodeBlock = { label: string; code: string };

const CODE_TOKEN = /^\uE002(\d+)\uE003$/;

/**
 * Replaces each fenced block with one marker line and appends the block to `blocks`. Mentions are resolved
 * after this, so `@{...}` inside code (a PowerShell hashtable) stays code. A fence that never closes runs to the end.
 */
export function extractFences(text: string, blocks: CodeBlock[] = []): { text: string; blocks: CodeBlock[] } {
  const out: string[] = [];
  const lines = text.split("\n");
  for (let i = 0; i < lines.length; i++) {
    const open = lines[i]!.match(/^```([^`]*)$/);
    if (!open) { out.push(lines[i]!); continue; }
    const code: string[] = [];
    for (i++; i < lines.length && !/^```\s*$/.test(lines[i]!); i++) code.push(lines[i]!);
    blocks.push({ label: open[1]!.trim(), code: code.join("\n") });
    out.push(`\uE002${blocks.length - 1}\uE003`);
  }
  return { text: out.join("\n"), blocks };
}

/**
 * Rich format. Consecutive `- ` or `* ` lines become one bullet list, and consecutive `1. ` lines one numbered list.
 * `> ` lines make a quote. `#` to `######` start headings, with `####` and deeper as the third level. A fenced block becomes a Teams code block with syntax colors.
 * Inside a line, `*bold*`, `**bold**`, Slack `<url>` and `<url|label>`, Markdown `[label](url)`, bare URLs, and color tags such as `[red]text[/]`,
 * `[hl:yellow]text[/]`, and `[badge:green]text[/]` become the Teams shapes. Other non-empty lines are paragraphs.
 */
export function composeRich(text: string, mentions: User[], codeBlocks: CodeBlock[] = []): { html: string; entries: MentionEntry[]; links: LinkEntry[] } {
  const collector: Collector = { next: 0, entries: [], links: [], rich: true };
  const blocks = [...codeBlocks];
  let html = "";
  let open: "ul" | "ol" | undefined;
  const lines = extractFences(text, blocks).text.split("\n");
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index]!;
    const bullet = line.match(/^[-*]\s+(.*)$/);
    const number = line.match(/^\d+[.)]\s+(.*)$/);
    const kind = bullet ? "ul" : number ? "ol" : undefined;
    if (open && open !== kind) { html += `</${open}>`; open = undefined; }
    if (kind) {
      if (!open) { html += `<${kind}>`; open = kind; }
      html += `<li>${composeInline((bullet ?? number)![1]!, mentions, collector)}</li>`;
      continue;
    }
    const code = line.match(CODE_TOKEN);
    if (code) { const block = blocks[Number(code[1])]!; html += codeBlockHtml(block.label, block.code); continue; }
    if (HORIZONTAL_RULE.test(line)) { html += "<hr>"; continue; }
    const alignments = tableAlignments(line, lines[index + 1]);
    if (alignments) {
      const rows: string[][] = [splitRow(line)];
      for (index += 2; index < lines.length && lines[index]!.includes("|"); index++) rows.push(splitRow(lines[index]!));
      index--;
      html += tableHtml(rows, alignments, cell => composeInline(cell, mentions, collector));
      continue;
    }
    if (/^>\s?/.test(line)) {
      // Consecutive quote lines are one block. Teams puts the lines in one paragraph, joined by line breaks.
      const quoted: string[] = [];
      for (; index < lines.length && /^>\s?/.test(lines[index]!); index++) quoted.push(composeInline(lines[index]!.replace(/^>\s?/, ""), mentions, collector));
      index--;
      html += `<blockquote><p>${quoted.join("<br>")}</p></blockquote>`;
      continue;
    }
    const heading = line.match(/^(#{1,6})\s+(.*)$/);
    if (heading) {
      // Teams has three heading levels. Four hashes and more give the third, as the Teams composer does.
      const level = Math.min(heading[1]!.length, 3);
      html += `<h${level}>${composeInline(heading[2]!, mentions, collector)}</h${level}>`;
      continue;
    }
    if (line.trim() !== "") html += `<p>${composeInline(line, mentions, collector)}</p>`;
  }
  if (open) html += `</${open}>`;
  return { html, entries: collector.entries, links: collector.links };
}

/** Three or more dashes, stars or underscores alone on a line. Teams stores a bare `<hr>`. */
const HORIZONTAL_RULE = /^\s*(?:-{3,}|\*{3,}|_{3,})\s*$/;

/** Splits a pipe row into cells. `\|` is a pipe inside a cell, and the outer pipes are optional. */
function splitRow(line: string): string[] {
  const cells = line.trim().replace(/^\|/, "").replace(/(?<!\\)\|$/, "").split(/(?<!\\)\|/);
  return cells.map(cell => cell.trim().replace(/\\\|/g, "|"));
}

type Align = "left" | "center" | "right" | undefined;

/** A table starts at a header row with a pipe, when the next line is a separator row with the same number of cells. */
function tableAlignments(header: string, separator: string | undefined): Align[] | undefined {
  if (!header.includes("|") || separator === undefined || !separator.includes("-")) return undefined;
  const marks = splitRow(separator);
  if (!marks.every(mark => /^:?-+:?$/.test(mark)) || marks.length !== splitRow(header).length) return undefined;
  return marks.map(mark => mark.startsWith(":") ? (mark.endsWith(":") ? "center" : "left") : mark.endsWith(":") ? "right" : undefined);
}

/**
 * The Teams table shape that rendered in the web client: a `thead` with `th` cells, a `tbody`, a `p` in every cell,
 * and the `align` attribute. A style on the cell or the `p` is ignored by Teams, so the CLI does not use one.
 */
function tableHtml(rows: string[][], alignments: Align[], inline: (cell: string) => string): string {
  const cell = (tag: "th" | "td", text: string | undefined, column: number) =>
    `<${tag}${alignments[column] ? ` align="${alignments[column]}"` : ""}><p>${text ? inline(text) : "&nbsp;"}</p></${tag}>`;
  const row = (tag: "th" | "td", cells: string[]) => `<tr>${alignments.map((_, column) => cell(tag, cells[column], column)).join("")}</tr>`;
  return `<table><thead>${row("th", rows[0]!)}</thead><tbody>${rows.slice(1).map(cells => row("td", cells)).join("")}</tbody></table>`;
}

/** Tags whose content is not message text. A `@{Name}` inside them stays as written. */
const LITERAL_TAGS = new Set(["code", "pre", "script", "style"]);
/** Stands in for the `@` of an escaped `@@{` while the tokens are resolved. */
const ESCAPED_AT = "\uE002";

/**
 * Finds the parts of an HTML string that are message text. Tags, attribute values, comments, and the
 * content of `<code>`, `<pre>`, `<script>`, and `<style>` are not text. A tag that is never closed
 * hides the rest of the string, so a stray `<` cannot make the CLI rewrite markup.
 */
export function htmlTextSpans(html: string): Array<{ start: number; end: number }> {
  const spans: Array<{ start: number; end: number }> = [];
  let literalDepth = 0;
  let textStart = 0;
  const flush = (end: number) => { if (!literalDepth && end > textStart) spans.push({ start: textStart, end }); };
  let i = 0;
  while (i < html.length) {
    if (html[i] !== "<") { i++; continue; }
    if (html.startsWith("<!--", i)) {
      flush(i);
      const close = html.indexOf("-->", i + 4);
      i = close < 0 ? html.length : close + 3;
      textStart = i;
      continue;
    }
    // A `<` that cannot start a tag is text.
    if (!/[A-Za-z/!?]/.test(html[i + 1] ?? "")) { i++; continue; }
    flush(i);
    let end = i + 1;
    let quote = "";
    while (end < html.length) {
      const char = html[end]!;
      if (quote) { if (char === quote) quote = ""; } else if (char === "\"" || char === "'") quote = char;
      else if (char === ">") break;
      end++;
    }
    const tag = /^<(\/?)([A-Za-z][^\s/>]*)/.exec(html.slice(i, end));
    if (tag && LITERAL_TAGS.has(tag[2]!.toLowerCase())) {
      if (tag[1]) literalDepth = Math.max(0, literalDepth - 1);
      else if (html[end - 1] !== "/") literalDepth++;
    }
    i = end + 1;
    textStart = i;
  }
  flush(html.length);
  return spans;
}

/**
 * Resolves `@{Name}` in the text of an HTML message. Each token becomes the same mention markup text
 * mode writes, with item ids in document order. Write `@@{` for a literal `@{` in text. Everything
 * else in the HTML is kept byte for byte.
 */
export async function composeHtml(raw: string, scope?: MentionScope): Promise<{ content: string; visible: string; warnText: string; entries: MentionEntry[] }> {
  const users: User[] = [];
  let marked = "";
  let last = 0;
  for (const span of htmlTextSpans(raw)) {
    marked += raw.slice(last, span.start);
    last = span.end;
    const guarded = raw.slice(span.start, span.end).replace(/@@\{/g, `${ESCAPED_AT}{`);
    // The text is HTML, so a name such as `Doe &amp; Co` must be decoded before the lookup.
    let decoded = "";
    let from = 0;
    for (const token of mentionTokens(guarded)) {
      decoded += `${guarded.slice(from, token.start)}@{${decodeEntities(token.query)}}`;
      from = token.end;
    }
    const resolved = await resolveMentionTokens(decoded + guarded.slice(from), scope);
    const offset = users.length;
    users.push(...resolved.mentions);
    // The span is already HTML text, so emoji become markup and nothing else is escaped again.
    marked += emojifyText(resolved.text, value => value).replace(PLACEHOLDER, (_, index: string) => `${PLACEHOLDER_START}${offset + Number(index)}${PLACEHOLDER_END}`).replaceAll(ESCAPED_AT, "@");
  }
  marked += raw.slice(last);

  const entries: MentionEntry[] = users.map((user, index) => ({ "@type": MENTION_TYPE, itemid: String(index), mri: user.mri, mentionType: user.mentionType ?? "person", displayName: user.displayName }));
  const content = marked.replace(PLACEHOLDER, (_, index: string) => mentionHtml(String(index), users[Number(index)]!.displayName));
  // The warning must not see the mentions, as in text mode, so it reads the text before they become @Name.
  const warnText = htmlToText(marked);
  const visible = warnText.replace(PLACEHOLDER, (_, index: string) => `@${users[Number(index)]!.displayName}`);
  return { content, visible, warnText, entries };
}

/** Finds plain `@Word` text that will not notify anyone. Mention placeholders never match. */
export function unresolvedMentionWarning(visible: string): string | undefined {
  const found = new Set<string>();
  for (const match of visible.matchAll(/(?:^|[\s([{])@([A-Za-z][\w'-]*)/g)) found.add(`@${match[1]}`);
  if (!found.size) return undefined;
  return `Message contains ${[...found].join(", ")} as plain text, not a Teams mention. It will not notify anyone. Did you mean to write @{Name} so it resolves to a real mention?`;
}

/** Refuses a request body above 100,000 UTF-8 bytes. The service's own limit is 102,400 bytes in total. */
export function checkBodySize(body: unknown): void {
  const bytes = Buffer.byteLength(JSON.stringify(body), "utf8");
  if (bytes > MAX_BODY_BYTES) {
    throw new Error(`Message is ${bytes.toLocaleString("en-US")} bytes, above the ${MAX_BODY_BYTES.toLocaleString("en-US")} byte limit. Teams allows 102,400 bytes in total.`);
  }
}

/** Builds the send request body in the browser shape. */
export function buildSendBody(input: {
  conversationId: string;
  content: string;
  mentions: MentionEntry[];
  links?: LinkEntry[];
  mri: string;
  name: string;
  now: string;
  clientMessageId: string;
  threadTs?: string;
  channelPost?: boolean;
  /** Object ids of uploaded inline images. */
  amsReferences?: string[];
  /** `properties.files` entries, from `upload.ts`. */
  files?: unknown[];
}) {
  return {
    id: "-1",
    type: "Message",
    conversationid: input.conversationId,
    conversationLink: `blah/${input.conversationId}${input.threadTs ? `;messageid=${input.threadTs}` : ""}`,
    from: input.mri,
    fromUserId: input.mri,
    composetime: input.now,
    originalarrivaltime: input.now,
    content: input.content,
    messagetype: "RichText/Html",
    contenttype: "Text",
    imdisplayname: input.name,
    clientmessageid: input.clientMessageId,
    callId: "",
    state: 0,
    version: "0",
    amsreferences: input.amsReferences ?? [],
    properties: {
      importance: "",
      subject: "",
      title: "",
      cards: "[]",
      links: input.links?.length ? JSON.stringify(input.links) : "[]",
      mentions: input.mentions.length ? JSON.stringify(input.mentions) : "[]",
      onbehalfof: null,
      files: input.files?.length ? JSON.stringify(input.files) : "[]",
      policyViolation: null,
      formatVariant: "TEAMS",
    },
    ...(input.threadTs ? { replyChainId: input.threadTs } : input.channelPost ? { postType: "Standard" } : {}),
    crossPostChannels: [],
  };
}

/**
 * Wraps a `buildSendBody()` result for `POST /users/ME/drafts`. The shape is from the 2026-10-19
 * capture of the web client: the inner message keeps its normal fields and gains `draftId`,
 * `draftDetails`, `threadtype`, and `innerThreadId`.
 */
export function buildScheduledBody(message: ReturnType<typeof buildSendBody>, sendAt: Date, threadId = message.conversationid) {
  return {
    draftDetails: { sendAt: String(sendAt.getTime()) },
    draftType: "ScheduledDraft",
    innerThreadId: threadId,
    message: {
      ...message,
      properties: { ...message.properties, draftId: message.clientmessageid },
      draftDetails: { sendAt: sendAt.toISOString() },
      threadtype: "streamofdrafts",
      innerThreadId: threadId,
    },
  };
}

/** A 19 digit decimal string, as the web client sends for clientmessageid. */
export function newClientMessageId(): string {
  const digits = Array.from({ length: 18 }, () => Math.floor(Math.random() * 10)).join("");
  return `${1 + Math.floor(Math.random() * 9)}${digits}`;
}
