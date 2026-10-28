import { DEFAULT_SEARCH_COUNT, MAX_SEARCH_COUNT } from "../limits.ts";
import { randomUUID } from "node:crypto";
import { findChannel, loadConversations, labelFor, resolveConversation } from "../conversations.ts";
import { channelTarget } from "../upload.ts";
import { formatMessageRef, isChannelId, isConversationId, isEmail } from "../refs.ts";
import { identity } from "../tokens.ts";
import { localTimeZone, request } from "../transport.ts";
import { searchPeople, type User } from "../users.ts";

export type SearchOptions = {
  json?: boolean; count?: number; from?: string; after?: string; before?: string; on?: string;
  in?: string; files?: boolean; type?: string;
};

/** What the query builders take: the resolved parts, not the raw flags. */
export type QueryParts = { text: string; email?: string; thread?: string; after?: string; before?: string; on?: string; type?: string; folder?: string };

const PAGE_SIZE = 25;
const SOFT_DELETED = "NOT (isClientSoftDeleted:TRUE)";
const NOT_TOPIC = "NOT (Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String:(topic OR space))";
const EXT = "Extension_SkypeSpaces_ConversationPost_Extension_";

function dateParts(raw: string): [number, number, number] {
  const time = Date.parse(raw);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw) || !Number.isFinite(time) || new Date(time).toISOString().slice(0, 10) !== raw) {
    throw new Error("Search dates must use YYYY-MM-DD.");
  }
  const [year, month, day] = raw.split("-").map(Number) as [number, number, number];
  return [year, month, day];
}

/** The UTC instant of local midnight at the start of a YYYY-MM-DD day, plus `days` whole days. */
export function localMidnight(raw: string, days = 0): string {
  const [year, month, day] = dateParts(raw);
  return new Date(year, month - 1, day + days).toISOString();
}

/** The sent-time bounds of the date flags: `from` is inclusive, `to` is exclusive. Both are UTC ISO strings. */
export function dateRange(options: { after?: string; before?: string; on?: string }): { from?: string; to?: string } {
  if (options.on !== undefined) {
    if (options.after !== undefined || options.before !== undefined) throw new Error("--on cannot be used with --after or --before.");
    return { from: localMidnight(options.on), to: localMidnight(options.on, 1) };
  }
  const from = options.after === undefined ? undefined : localMidnight(options.after);
  const to = options.before === undefined ? undefined : localMidnight(options.before);
  if (from !== undefined && to !== undefined && from >= to) throw new Error("--after must be an earlier day than --before.");
  return { ...(from ? { from } : {}), ...(to ? { to } : {}) };
}

function dateClause(field: string, range: { from?: string; to?: string }): string[] {
  return [...(range.from ? [`${field} >= ${range.from}`] : []), ...(range.to ? [`${field} < ${range.to}`] : [])];
}

function words(parts: QueryParts): string {
  return [parts.email, parts.text.trim()].filter(Boolean).join(" ");
}

/**
 * The Messages-tab query string. A channel id uses `clientthreadid:<id> AND <words>`. Any other conversation uses the chat
 * form, which also hides topic changes and deleted messages (captures `15-search-14`, `16-search-85`, `16-search-105`).
 */
export function messageQuery(parts: QueryParts): string {
  const clauses: string[] = [];
  const filtered = !!(parts.thread || parts.email || parts.after || parts.before || parts.on);
  const channel = !!parts.thread && isChannelId(parts.thread);
  if (parts.thread) clauses.push(`clientthreadid:${parts.thread}`);
  if (parts.thread && !channel) clauses.push(NOT_TOPIC);
  clauses.push(...dateClause("sent", dateRange(parts)));
  // Deleted messages stay out of every search, as the All tab and the chat filter do.
  clauses.push(SOFT_DELETED);
  const free = words(parts);
  if (!clauses.length) return free;
  const soft = clauses.includes(SOFT_DELETED);
  return `${clauses.join(" AND ")}${free ? `${soft ? " " : " AND "}${free}` : ""}`;
}

/** The Files-tab query string (captures `15-search-23`, `16-search-67`). */
export function fileQuery(parts: QueryParts): string {
  const clauses = dateClause("LastModifiedTime", dateRange(parts));
  // The channel annotation alone only ranks the channel's files first. ParentLink keeps the folder and its sub-folders, and
  // a library root, whose files list `<root>/Forms/AllItems.aspx` as their parent (checked live against the folders).
  if (parts.folder) clauses.unshift(`(ParentLink:"${parts.folder}" OR ParentLink:"${parts.folder}/*")`);
  if (parts.type) {
    const type = fileType(parts.type);
    clauses.push(`(filetype:${type} OR fileExtension:${type})`);
  }
  const free = words(parts);
  return `${clauses.join(" AND ")}${clauses.length && free ? " AND " : ""}${free}`;
}

/** A `--type` value: an extension such as `pdf` or `.docx`, or `loop` or `fluid`. */
export function fileType(raw: string): string {
  const type = raw.trim().replace(/^\./, "").toLowerCase();
  if (!/^[a-z0-9]{1,16}$/.test(type)) throw new Error("--type needs a file extension such as pdf or docx, or loop or fluid.");
  return type;
}

/** The Messages-tab request, copied from captures `16-search-72` and `16-search-85`. */
export function searchBody(query: string, from: number, size: number, text = query): unknown {
  return {
    entityRequests: [{
      entityType: "Message", contentSources: ["Teams"],
      fields: [`${EXT}FromSkypeInternalId_String`, `${EXT}FileData_String`, `${EXT}ThreadType_String`, `${EXT}SkypeGroupId_String`,
        `${EXT}SenderTenantId_String`, `${EXT}ImageSrc_String`, `${EXT}ParentMessageId_String`, `${EXT}AmsReferences_StringArray`],
      propertySet: "Optimized", query: { queryString: query, displayQueryString: text }, from, size, topResultsCount: 5,
    }],
    QueryAlterationOptions: { EnableAlteration: true, EnableSuggestion: true, SupportedRecourseDisplayTypes: ["Suggestion"] },
    cvid: randomUUID(), logicalId: randomUUID(),
    scenario: { Dimensions: [{ DimensionName: "QueryType", DimensionValue: "Messages" },
      { DimensionName: "FormFactor", DimensionValue: "general.web.reactSearch" }], Name: "powerbar" },
    timezone: localTimeZone(),
  };
}

const FILE_FIELDS = ["DefaultEncodingUrl", "FileName", "FileType", "HitHighlightedSummary", "LastModifiedTime", "LinkingUrl", "ModifiedBy",
  "OriginalPath", "Path", "Title", "ServerRedirectedPreviewUrl", "SpWebUrl", "ChannelGroupId", "FileExtension", "VisualizationAccessURL",
  "LastSharedWithMailboxOwnerByDisplayName", "LastSharedWithMailboxOwnerDateTime", "FileContextType", "ClassicAttachmentVisualizationUrl",
  "UniqueId", "SiteName", "Size", "Annotations", "ListId", "SiteId", "WebId", "ParentLink", "ContentClass"];
const FILE_REFINER = "or(andnot(IsDocument:true,Title:or(OneNote_DeletedPages,OneNote_RecycleBin),SecondaryFileExtension:onetoc2,FileExtension:vtt,ContentClass:ExternalLink,and(ContentClass:STS_List_DocumentLibrary,SiteTemplateId:21),FileType:or(aspx,htm,html,mhtml),and(ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118*,PromotedState:2)),ContentTypeId:or(0x010100F3754F12A9B6490D9622A01FE9D8F012,0x0120D520A808*),SecondaryFileExtension:or(wmv,avi,mpg,asf,mp4,ogg,ogv,webm,mov),FileType:or(ai,bmp,dib,dst,emb,eps,gif,ico,jpeg,jpg,odg,png,rle,svg,tiff,webp,wmf,wpd))";

/** The Files-tab request, copied from capture `16-search-70`. */
/** The annotation that scopes a file search to one channel's files folder (capture `17-search-files-in-channel`). */
export type ChannelAnnotation = { type: "Channel"; threadId: string; text: "dummyText"; groupId: string; path: string; tenantId: string };

export function fileSearchBody(query: string, from: number, size: number, text = query, channel?: ChannelAnnotation): unknown {
  return {
    entityRequests: [{
      // A channel scope searches SharePoint only, as the web client does.
      entityType: "File", contentSources: channel ? ["SharePoint"] : ["OneDriveBusiness", "SharePoint", "Exchange"], fields: FILE_FIELDS,
      query: { queryString: query, displayQueryString: text, ...(channel ? { queryAnnotations: [channel] } : {}) }, size, from,
      EnableQueryUnderstanding: false, EnableSpeller: false,
      HitHighlight: { HitHighlightedProperties: ["HitHighlightedSummary"], SummaryLength: 200 },
      IdFormat: "EwsId", ParserType: "None", propertySet: "Optimized", ResultsMerge: { Type: "Interleaved" },
      Sort: [{ Field: "PersonalScore", SortDirection: "Desc" }],
      extendedQueries: [{ SearchProvider: "SharePoint", Query: { SourceId: "8413CD39-2156-4E00-B54D-11EFD9ABDB89", EnableQueryRules: false,
        TrimDuplicates: false, BypassResultTypes: true, ProcessBestBets: false, ProcessPersonalFavorites: false, EnableInterleaving: false,
        EnableMultiGeo: true, RankingModelId: "ABBAABBA-AAAA-AAAA-CCCC-000000000426", Culture: 1033 } }],
      RefiningQueries: [{ RefinerString: FILE_REFINER }],
      EnableResultAnnotations: true, AnnotationsCount: 1, MaxOverallAggregationBucketCount: 5,
    }],
    QueryAlterationOptions: { EnableSuggestion: true },
    cvid: randomUUID(), logicalId: randomUUID(),
    scenario: { Dimensions: [{ DimensionName: "QueryType", DimensionValue: "Files" },
      { DimensionName: "FormFactor", DimensionValue: "general.cdlworker.reactSearch" }], Name: "powerbar" },
    textDecorations: "Off", timezone: localTimeZone(),
  };
}

/** The local `YYYY-MM-DD HH:MM` of an ISO time, or the input when it does not parse. */
export function localStamp(iso: unknown): string {
  const time = typeof iso === "string" ? new Date(iso) : undefined;
  if (!time || Number.isNaN(time.getTime())) return typeof iso === "string" ? iso : "";
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${time.getFullYear()}-${pad(time.getMonth() + 1)}-${pad(time.getDate())} ${pad(time.getHours())}:${pad(time.getMinutes())}`;
}

/** The file names in a `FileData` extension: a JSON string of `{type, rawJson}` items, where `rawJson` is a JSON string of files. */
export function fileNames(fileData: unknown): { has: boolean; names: string[] } {
  const parse = (value: unknown): unknown => { try { return typeof value === "string" ? JSON.parse(value) : value; } catch { return undefined; } };
  const list = parse(fileData);
  if (!Array.isArray(list)) return { has: false, names: [] };
  const items: any[] = [];
  for (const entry of list) {
    const inner = parse(entry?.rawJson);
    if (Array.isArray(inner)) items.push(...inner);
    else if (entry && !("rawJson" in entry)) items.push(entry);
    else if (inner && typeof inner === "object") items.push(inner);
  }
  const names = items.flatMap(item => {
    const name = [item?.fileName, item?.title, item?.name, item?.FileName].find(value => typeof value === "string" && value);
    return name ? [String(name)] : [];
  });
  return { has: items.length > 0, names };
}

/** Message refs use the source conversation and timestamp, never the search result's own ID. */
export function normalizeHit(hit: any) {
  const source = hit?.Source ?? {};
  const extensions = source.Extensions ?? {};
  const conversation = [source.ClientThreadId, source.ClientConversationId, source.ConversationId?.Id]
    .find(value => typeof value === "string" && isConversationId(value));
  const ts = String(source.InternetMessageId ?? "");
  const ref = typeof conversation === "string" && /^\d{10,}$/.test(ts) ? formatMessageRef(conversation, ts) : undefined;
  const parent = String(extensions.SkypeSpaces_ConversationPost_Extension_ParentMessageId ?? "");
  const files = fileNames(extensions.SkypeSpaces_ConversationPost_Extension_FileData);
  const ams = extensions.SkypeSpaces_ConversationPost_Extension_AmsReferences;
  const image = !!extensions.SkypeSpaces_ConversationPost_Extension_ImageSrc || (Array.isArray(ams) && ams.length > 0);
  return {
    ref,
    conversation_id: typeof conversation === "string" ? conversation : undefined,
    sent: source.DateTimeSent,
    sent_local: localStamp(source.DateTimeSent),
    sender: source.From?.EmailAddress?.Name ?? source.From?.EmailAddress?.Address ?? "",
    text: String(source.Preview ?? hit?.HitHighlightedSummary ?? "").replace(/<[^>]*>/g, ""),
    team: source.TeamName, channel: source.ChannelName,
    has_files: files.has,
    ...(files.has ? { files: files.names } : {}),
    has_image: image,
    reply_in_thread: typeof conversation === "string" && /^\d{10,}$/.test(parent) && parent !== ts ? formatMessageRef(conversation, parent) : undefined,
    mentions_you: source.MentionsPreview?.IsMentioned === true,
  };
}

type Hit = ReturnType<typeof normalizeHit>;

/** The markers that follow a message hit's id in text output. */
export function hitMarkers(hit: Hit): string[] {
  return [
    ...(hit.has_files ? [hit.files?.length ? `[files: ${hit.files.join(", ")}]` : "[files]"] : []),
    ...(hit.has_image ? ["[image]"] : []),
    ...(hit.mentions_you ? ["[mentions you]"] : []),
    ...(hit.reply_in_thread ? [`reply in thread ${hit.reply_in_thread}`] : []),
  ];
}

export function normalizeFile(hit: any) {
  const source = hit?.Source ?? {};
  const text = (value: unknown) => (typeof value === "string" && value ? value : undefined);
  return {
    name: text(source.FileName) ?? text(source.Title) ?? "",
    type: text(source.FileType) ?? text(source.FileExtension),
    size: typeof source.Size === "number" ? source.Size : undefined,
    site: text(source.SiteName) ?? text(source.SpWebUrl),
    folder: text(source.ParentLink),
    modified: text(source.LastModifiedTime),
    modified_local: localStamp(source.LastModifiedTime),
    modified_by: text(source.ModifiedBy),
    link: text(source.LinkingUrl) ?? text(source.DefaultEncodingUrl) ?? text(source.Path),
    path: text(source.DefaultEncodingUrl) ?? text(source.Path),
    unique_id: text(source.UniqueId)?.replace(/[{}]/g, ""),
  };
}

/** Resolves `--from` to one email. An email is used as typed. A name goes through the `users find` lookup and must match one person. */
export async function resolveSender(input: string): Promise<string> {
  const trimmed = input.trim().replace(/^@/, "");
  if (isEmail(trimmed)) return trimmed;
  const matches = await searchPeople(trimmed);
  const needle = trimmed.toLowerCase();
  const exact = matches.filter(user => [user.displayName, user.realName, user.mail, user.upn].some(value => value?.toLowerCase() === needle));
  const candidates: User[] = exact.length ? exact : matches;
  const label = (user: User) => `${user.displayName}${user.upn || user.mail ? ` <${user.upn ?? user.mail}>` : ""}`;
  if (candidates.length === 0) throw new Error(`--from "${input}" matches nobody. Check the name, or pass an email.`);
  if (candidates.length > 1) {
    throw new Error(`--from "${input}" matches ${candidates.length} people: ${candidates.slice(0, 8).map(label).join("; ")}. Pass the email of one.`);
  }
  const email = candidates[0]!.upn ?? candidates[0]!.mail;
  if (!email) throw new Error(`${candidates[0]!.displayName} has no email in the directory. Pass an email to --from.`);
  return email;
}

/** `--in` with `--files`: the channel's files folder, as Teams scopes a Files search. Teams offers no scope for a chat's files. */
export async function channelScope(input: string): Promise<ChannelAnnotation> {
  const conversation = await resolveConversation(input);
  if (!isChannelId(conversation.id)) throw new Error(`--in with --files takes a channel. ${conversation.label} is a chat, and Teams does not scope a file search to a chat.`);
  const target = await channelTarget(conversation.id);
  let { team } = await findChannel(conversation.id);
  if (!team.groupId) ({ team } = await findChannel(conversation.id, { fresh: true }));
  if (!team.groupId) throw new Error(`Teams does not list the group of the team ${team.name}, so the file search cannot be scoped to ${target.label}.`);
  return { type: "Channel", threadId: conversation.id, text: "dummyText", groupId: team.groupId, path: `https://${target.host}${target.folder}/`, tenantId: (await identity()).tenantId };
}

export async function search(text: string, options: SearchOptions): Promise<void> {
  const count = options.count ?? DEFAULT_SEARCH_COUNT;
  if (!Number.isInteger(count) || count < 1 || count > MAX_SEARCH_COUNT) throw new Error(`--count must be a whole number from 1 to ${MAX_SEARCH_COUNT}.`);
  if (options.type && !options.files) throw new Error("--type works only with --files.");
  if (options.files ? !text.trim() && !options.from && !options.type : !text.trim() && !options.from && !options.in) {
    throw new Error(options.files ? "Give text, --from, or --type to search for." : "Give text, --from, or --in to search for.");
  }
  dateRange(options);
  const parts: QueryParts = { text, after: options.after, before: options.before, on: options.on, type: options.type };
  if (options.from) parts.email = await resolveSender(options.from);
  let channel: ChannelAnnotation | undefined;
  if (options.in && options.files) {
    channel = await channelScope(options.in);
    parts.folder = channel.path.replace(/\/+$/, "");
  }
  else if (options.in) parts.thread = (await resolveConversation(options.in)).id;
  const query = options.files ? fileQuery(parts) : messageQuery(parts);
  const entity = options.files ? "File" : "Message";
  const cache = await loadConversations((await identity()).tenantId);
  const hits: any[] = [];
  const seen = new Set<string>();
  const keyOf = options.files
    ? (item: any, hit: ReturnType<typeof normalizeFile>) => `${hit.unique_id ?? hit.path ?? item.Id ?? ""}|${hit.modified ?? ""}`
    : (item: any, hit: Hit) => hit.ref ?? `${item.Id ?? ""}|${hit.sent ?? ""}|${hit.text}`;
  const normalize = (options.files ? normalizeFile : normalizeHit) as (item: any) => any;
  let offset = 0;
  let total: number | undefined;
  let more = false;
  for (;;) {
    // Keep the page size stable. A one-hit final page can otherwise contain only a duplicate.
    const body = await request("substrate", "POST", "/searchservice/api/v2/query", {
      body: options.files ? fileSearchBody(query, offset, PAGE_SIZE, text, channel) : searchBody(query, offset, PAGE_SIZE, text),
    });
    const entitySets: any[] = Array.isArray(body?.EntitySets) ? body.EntitySets : [];
    const set = entitySets.find(item => item?.EntityType === entity) ?? entitySets[0];
    const result = (Array.isArray(set?.ResultSets) ? set.ResultSets : [])[0];
    const results = Array.isArray(result?.Results) ? result.Results : [];
    if (typeof result?.Total === "number") total = result.Total;
    more = result?.MoreResultsAvailable === true;
    let added = 0;
    for (const [index, item] of results.entries()) {
      const hit = normalize(item);
      const key = keyOf(item, hit);
      if (seen.has(key)) continue;
      seen.add(key); hits.push(hit); added++;
      if (hits.length >= count) { if (index < results.length - 1) more = true; break; }
    }
    offset += results.length;
    if (hits.length >= count || !more || !results.length || !added) break;
  }
  if (options.files) {
    if (options.json) console.log(JSON.stringify({ query, total, more, files: hits }, null, 2));
    else for (const file of hits) {
      console.log(file.name);
      const place = file.folder ?? file.site;
      if (place) console.log(`  in: ${place}`);
      console.log(`  modified: ${file.modified_local}${file.modified_by ? ` by ${file.modified_by}` : ""}`);
      if (file.link) console.log(`  link: ${file.link}`);
    }
    return;
  }
  const messages = (hits as Hit[]).map(hit => ({
    ...hit,
    conversation: (hit.conversation_id ? labelFor(cache, hit.conversation_id) : undefined)
      ?? (hit.team && hit.channel ? `${hit.team}/${hit.channel}` : undefined) ?? hit.conversation_id,
  }));
  if (options.json) console.log(JSON.stringify({ query, total, more, messages }, null, 2));
  else for (const message of messages) {
    console.log(`[${message.sent_local}] ${message.sender}${message.conversation ? ` in ${message.conversation}` : ""}\n  ${message.text}`);
    console.log(`  id: ${message.ref ?? "no message ref"}${hitMarkers(message).map(marker => `  ${marker}`).join("")}`);
  }
}
