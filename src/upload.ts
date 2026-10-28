import type { FileAccess, FileRole } from "./settings.ts";
import { createHash, randomUUID } from "node:crypto";
import { createReadStream } from "node:fs";
import { stat } from "node:fs/promises";
import { basename, extname } from "node:path";
import { findChannel, loadConversations, saveConversations } from "./conversations.ts";
import { encodeConversation } from "./refs.ts";
import { parseSharepointHost, personalSite, sharepointHost, tenantSharepointHosts } from "./sharepoint.ts";
import { hasToken, identity } from "./tokens.ts";
import { request, TeamsError } from "./transport.ts";

import { MAX_FILES, MAX_UPLOAD_BYTES, mb } from "./limits.ts";

export { MAX_FILES, MAX_UPLOAD_BYTES };
/** A channel file above this size goes up in pieces of this size. The web client sent its whole 62 KB capture in one call. */
export const CHUNK_BYTES = 10 * 1024 * 1024;
export const CHAT_FOLDER = "Microsoft Teams Chat Files";
export const FILE_TYPE = "http://schema.skype.com/File";

export type LoadedFile = { path: string; name: string; size: number; digest: string };

/** What a send tells the user when it fails after a file went up: the files that are still on SharePoint. */
export class UploadFailure extends Error {
  constructor(message: string, readonly leftovers: string[], readonly uncertain = false) {
    super(message);
    this.name = "UploadFailure";
  }
}

const BAD_CHARS = /["*:<>?/\\|\u0000-\u001f]/;

/** Refuses a name that SharePoint would reject, before any upload. */
export function checkUploadName(name: string): void {
  if (BAD_CHARS.test(name)) throw new Error(`${name} cannot be uploaded. SharePoint does not allow " * : < > ? / \\ | in a file name. Rename the file.`);
  if (name !== name.trim() || name.endsWith(".")) throw new Error(`${name} cannot be uploaded. SharePoint does not allow a name that starts or ends with a space, or ends with a dot. Rename the file.`);
  if (name.startsWith("~$") || name.includes("_vti_")) throw new Error(`${name} cannot be uploaded. SharePoint does not allow a name that starts with ~$ or contains _vti_. Rename the file.`);
  if (name.length > 255) throw new Error(`${name} cannot be uploaded. The name is longer than 255 characters.`);
}

function sha256File(path: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const hash = createHash("sha256");
    createReadStream(path).on("data", chunk => hash.update(chunk)).on("error", reject).on("end", () => resolve(hash.digest("hex")));
  });
}

/** Reads one file's size and SHA-256 without holding its bytes. Refuses a missing file, a folder, an empty file, and a file over 250 MB. */
export async function loadFile(path: string): Promise<LoadedFile> {
  let info;
  try {
    info = await stat(path);
  } catch {
    throw new Error(`Cannot read the file: ${path}`);
  }
  if (!info.isFile()) throw new Error(`${path} is not a file.`);
  if (info.size === 0) throw new Error(`${path} is empty.`);
  if (info.size > MAX_UPLOAD_BYTES) throw new Error(`${path} is ${(info.size / 1024 / 1024).toFixed(1)} MB. The limit is ${mb(MAX_UPLOAD_BYTES)}.`);
  const name = basename(path);
  checkUploadName(name);
  try {
    return { path, name, size: info.size, digest: await sha256File(path) };
  } catch {
    throw new Error(`Cannot read the file: ${path}`);
  }
}

/** Checks a list of files as a set: the count, and no two with one name. */
export function checkFileSet(files: LoadedFile[]): void {
  if (files.length > MAX_FILES) throw new Error(`A message can hold ${MAX_FILES} files at most.`);
  const names = new Set<string>();
  for (const file of files) {
    if (names.has(file.name.toLowerCase())) throw new Error(`Two files are named ${file.name}. Rename one of them.`);
    names.add(file.name.toLowerCase());
  }
}

/** A SharePoint REST string value: single quotes doubled, then the whole value URL-encoded, as the web client sends it. */
function odata(value: string): string {
  return encodeURIComponent(`'${value.replace(/'/g, "''")}'`);
}

/** Time limit for one upload request: a minute, and four more seconds for each megabyte. */
function uploadTimeout(bytes: number): number {
  return 60_000 + Math.ceil(bytes / 1_048_576) * 4_000;
}

function extensionOf(name: string): string {
  return extname(name).slice(1).toLowerCase();
}

// ---------------------------------------------------------------------------------------------------------------
// Chats and Notes to self: OneDrive upload, then the message.
// ---------------------------------------------------------------------------------------------------------------

/** What the personal-site upload returns, reduced to the fields that the message needs. */
export type ChatUpload = {
  name: string;
  /** The drive item id, such as `01ABC...`. */
  itemId: string;
  /** The list item GUID. It is the `itemid` of the message entry and the id in the invite address. */
  uniqueId: string;
  fileUrl: string;
  siteUrl: string;
  driveId: string;
  listId: string;
  siteId: string;
  webId: string;
};

function text(value: unknown): string {
  return typeof value === "string" ? value : "";
}

/** Reads the upload answer of capture 375. Every field is checked, because the message is built from them. */
export function parseChatUpload(item: any): ChatUpload {
  const ids = item?.sharepointIds;
  const upload: ChatUpload = {
    name: text(item?.name),
    itemId: text(item?.id),
    uniqueId: text(ids?.listItemUniqueId),
    fileUrl: text(item?.webDavUrl),
    siteUrl: text(ids?.siteUrl).replace(/\/+$/, ""),
    driveId: text(item?.parentReference?.driveId),
    listId: text(ids?.listId),
    siteId: text(ids?.siteId),
    webId: text(ids?.webId),
  };
  const missing = Object.entries(upload).filter(([, value]) => !value).map(([key]) => key);
  if (missing.length) throw new TeamsError(`SharePoint did not return the file's ${missing.join(", ")}.`);
  return upload;
}

/** Uploads one file to `Microsoft Teams Chat Files` on your personal site. A name that exists is renamed by SharePoint. */
export async function uploadToOneDrive(site: { host: string; path: string }, file: LoadedFile): Promise<ChatUpload> {
  const path = `${site.path}/_api/v2.0/drive/root:/${encodeURIComponent(CHAT_FOLDER)}/${encodeURIComponent(file.name)}:/content?@name.conflictBehavior=rename&$select=*,sharepointIds,webDavUrl`;
  const item = await request("sharepoint", "PUT", path, {
    sharepointHost: site.host,
    bytes: Bun.file(file.path),
    contentType: "application/octet-stream",
    headers: { scenariotype: "AUO", scenario: "ShareUploadFile", prefer: "manualRedirect,pacToken=N" },
    timeoutMs: uploadTimeout(file.size),
  });
  return parseChatUpload(item);
}

/** The `properties.files` entry for a OneDrive file, in the shape of capture 397. The file is not shared by link, so the share fields stay null. */
export function chatFileEntry(upload: ChatUpload) {
  const type = extensionOf(upload.name);
  const base = `${upload.siteUrl}/`;
  return {
    itemid: upload.uniqueId,
    fileName: upload.name,
    fileType: type,
    fileInfo: { itemId: upload.itemId, fileUrl: upload.fileUrl, siteUrl: base, serverRelativeUrl: "", shareUrl: null, shareId: null },
    fileChicletState: { serviceName: "p2p", state: "active" },
    "@type": FILE_TYPE,
    version: 2,
    id: upload.uniqueId,
    baseUrl: base,
    objectUrl: upload.fileUrl,
    openUrl: null,
    type,
    title: upload.name,
    state: "active",
    chicletBreadcrumbs: null,
    providerData: "",
    botFileProperties: {},
    isUploadError: null,
    progressComplete: null,
    permissionScope: null,
    filePreview: { previewUrl: "", previewHeight: 0, previewWidth: 0 },
    sharepointIds: { driveId: upload.driveId, listId: upload.listId, listItemUniqueId: upload.uniqueId, siteId: upload.siteId, siteUrl: upload.siteUrl, webId: upload.webId },
    publication: null,
    site: null,
  };
}

/** Stands in for an upload answer in a preview and in the send-once key. The real values exist only after the upload. */
export function placeholderChatUpload(file: LoadedFile, index: number, site: { host: string; path: string }): ChatUpload {
  const n = index + 1;
  return {
    name: file.name,
    itemId: `<file-${n}-item-id>`,
    uniqueId: `<file-${n}-guid>`,
    fileUrl: `https://${site.host}${site.path}/Documents/${CHAT_FOLDER}/${encodeURIComponent(file.name)}`,
    siteUrl: `https://${site.host}${site.path}`,
    driveId: "<drive-id>",
    listId: "<list-id>",
    siteId: "<site-id>",
    webId: "<web-id>",
  };
}

// ---------------------------------------------------------------------------------------------------------------
// The invite: who can open a OneDrive file that a chat message names.
// ---------------------------------------------------------------------------------------------------------------

export type InviteSet = {
  recipients: Array<{ objectId: string; mri: string; name: string }>;
  skipped: Array<{ mri: string; reason: string }>;
};

/** The member ids of a chat, from the thread. */
export async function chatMembers(conversationId: string): Promise<string[]> {
  const thread: any = await request("chat", "GET", `/threads/${encodeConversation(conversationId)}`, { query: { view: "msnp24Equivalent" } });
  return (Array.isArray(thread?.members) ? thread.members : []).map((member: any) => text(member?.id)).filter(Boolean);
}

/**
 * Chooses who to invite, as the web client did in capture 07: every member except you and any member that has no
 * directory user. The check is a read. `fetchShortProfile` answers a member that exists with a user principal name,
 * and a deleted or disabled account with `accountEnabled: false` and no name or address (proven 2026-10-19: the one
 * member that capture 07 skipped came back that way). A failed read stops the send, so no one is skipped by a guess.
 */
export async function inviteRecipients(members: string[], selfMri: string): Promise<InviteSet> {
  const others = [...new Set(members)].filter(mri => mri !== selfMri);
  const skipped: InviteSet["skipped"] = others.filter(mri => !mri.startsWith("8:orgid:")).map(mri => ({ mri, reason: "not a directory user" }));
  const people = others.filter(mri => mri.startsWith("8:orgid:"));
  if (!people.length) return { recipients: [], skipped };
  if (!(await hasToken("spaces"))) throw new Error("Choosing who to invite needs a spaces token. Run: teams-cli auth login");
  const profiles: any[] = [];
  for (let start = 0; start < people.length; start += 100) {
    const json: any = await request("spaces", "POST", "/users/fetchShortProfile", {
      query: { isMailAddress: "false", enableGuest: "true", skypeTeamsInfo: "true", canBeSmtpAddress: "false", includeIBBarredUsers: "true", includeDisabledAccounts: "true" },
      body: people.slice(start, start + 100),
    });
    profiles.push(...(Array.isArray(json) ? json : Array.isArray(json?.value) ? json.value : []));
  }
  const recipients: InviteSet["recipients"] = [];
  for (const mri of people) {
    const profile = profiles.find(item => item?.mri === mri);
    const resolved = profile && text(profile.userPrincipalName) && profile.accountEnabled !== false;
    const objectId = text(profile?.objectId) || mri.slice("8:orgid:".length);
    if (!resolved) skipped.push({ mri, reason: profile ? "account disabled or deleted" : "no directory entry" });
    else recipients.push({ objectId, mri, name: text(profile.displayName) || text(profile.userPrincipalName) });
  }
  return { recipients, skipped };
}

/**
 * Gives the invited members edit access to one uploaded file. It asks for no link, so no link is created.
 * Without the default link, SharePoint needs `requireSignIn` or `sendInvitation` (400 "RequireSignIn and SendInvitation
 * cannot both be false", 2026-10-19). Sign-in is required and no invitation email is sent, so nobody is notified.
 */
export async function inviteToFile(
  site: { host: string; path: string },
  uniqueId: string,
  recipients: InviteSet["recipients"],
  sharing: { access: Exclude<FileAccess, "private">; role: FileRole } = { access: "people", role: "edit" },
): Promise<{ shareUrl: string; shareId: string } | undefined> {
  const roles = [sharing.role === "view" ? "read" : "write"];
  // `link` is the web client's request exactly (capture 07): no sign-in or invitation fields, and the default link.
  const body = sharing.access === "link"
    ? { recipients: recipients.map(item => ({ objectId: item.objectId })), roles }
    : { recipients: recipients.map(item => ({ objectId: item.objectId })), roles, requireSignIn: true, sendInvitation: false };
  const response: any = await request("sharepoint", "POST", `${site.path}/_api/v2.0/sites/root/items/${uniqueId}/driveItem/invite`, {
    sharepointHost: site.host,
    body,
    headers: { scenario: "TeamsFiles", prefer: sharing.access === "link" ? "ReturnSpecificErrorCode,getDefaultLink" : "ReturnSpecificErrorCode" },
  });
  // The message names the link the way the web client did: its address and the id of its permission.
  const link = Array.isArray(response?.value) ? response.value.find((item: any) => typeof item?.link?.webUrl === "string") : undefined;
  return link ? { shareUrl: link.link.webUrl, shareId: String(link.id) } : undefined;
}

// ---------------------------------------------------------------------------------------------------------------
// Channels: the files folder of the channel, on the team's site.
// ---------------------------------------------------------------------------------------------------------------

export type ChannelTarget = {
  host: string;
  /** The site path, such as `/sites/engineering`. */
  sitePath: string;
  siteUrl: string;
  /** The folder as a server-relative path, such as `/sites/engineering/Shared Documents/Engineering Tickets`. */
  folder: string;
  label: string;
};

/** The site that holds a folder: the team site when the folder is inside it, otherwise the `/sites/<name>` or `/teams/<name>` at the start of the path. */
export function siteOfFolder(folder: string, teamSiteUrl?: string): string | undefined {
  let team: string | undefined;
  try { team = teamSiteUrl ? new URL(teamSiteUrl).pathname.replace(/\/+$/, "") : undefined; } catch { team = undefined; }
  if (team && (folder === team || folder.startsWith(`${team}/`))) return team;
  return /^\/(?:sites|teams)\/[^/]+/.exec(folder)?.[0];
}

/**
 * Finds where a channel keeps its files. The folder is the `filesRelativePath` that Teams lists for the channel (the cached team
 * list). When the list has none, the channel thread's `channelDocsFolderRelativeUrl` is read, and a path found that way is cached.
 * Teams records neither until a channel's Files tab has been opened, so the send is refused then instead of guessing a folder.
 * A channel renamed after it was created keeps its first folder name, and the recorded path says which.
 */
export async function channelTarget(channelId: string): Promise<ChannelTarget> {
  const who = await identity();
  let { team, channel } = await findChannel(channelId);
  // A cache from before these fields were stored, or a folder that Teams recorded later, is read again once.
  if (!team.siteUrl || !channel.filesPath) ({ team, channel } = await findChannel(channelId, { fresh: true }));
  const label = `${team.name}/${channel.name}`;
  if (channel.isShared) throw new Error(`${label} is a shared channel. Its files live on a site that this CLI has not verified, so it cannot upload there.`);
  let folder = channel.filesPath;
  if (!folder) {
    const thread: any = await request("chat", "GET", `/threads/${encodeConversation(channelId)}`, { query: { view: "msnp24Equivalent" } });
    const found = text(thread?.properties?.channelDocsFolderRelativeUrl);
    if (found) {
      folder = found;
      const cache = await loadConversations(who.tenantId);
      const cached = cache.teams.find(item => item.id === team.id)?.channels.find(item => item.id === channelId);
      if (cached) {
        cached.filesPath = found;
        await saveConversations(who.tenantId, cache);
      }
    }
  }
  if (!folder) throw new Error(`Teams has no files folder recorded for ${label}. Open the channel's Files tab in Teams once, then try again.`);
  folder = folder.replace(/\/+$/, "");
  const sitePath = siteOfFolder(folder, team.siteUrl);
  if (!sitePath) throw new Error(`The files folder of ${label} (${folder}) is not on a /sites or /teams site, so the CLI cannot tell which site holds it.`);
  const host = parseSharepointHost(team.siteUrl);
  if (!host) throw new Error(`Teams does not name the SharePoint site of the team ${team.name}, so the CLI cannot tell which host holds the folder.`);
  if (!tenantSharepointHosts(await sharepointHost()).includes(host)) {
    throw new Error(`The files of ${label} are on ${host}, which is not a SharePoint host of your organization. The CLI will not send your token there.`);
  }
  return { host, sitePath, siteUrl: `https://${host}${sitePath}`, folder, label };
}

/** Reads the folder from SharePoint. A site or folder that is not there stops the send. */
export async function checkChannelFolder(target: ChannelTarget): Promise<void> {
  let folder: any;
  try {
    folder = await request("sharepoint", "GET", `${target.sitePath}/_api/web/GetFolderByServerRelativePath(decodedurl=@v)?@v=${odata(target.folder)}&$select=Exists,ServerRelativeUrl`, { sharepointHost: target.host });
  } catch (error) {
    if (error instanceof TeamsError && (error.status === 400 || error.status === 404)) {
      throw new Error(`SharePoint does not have the folder ${target.folder} on ${target.siteUrl}. Nothing was uploaded.`);
    }
    throw error;
  }
  if (folder?.Exists !== true) throw new Error(`The folder ${target.folder} does not exist on ${target.siteUrl}. Nothing was uploaded.`);
}

/** True when the folder already holds a file with this name. The upload would not replace it, so the send stops. */
export async function channelFileExists(target: ChannelTarget, name: string): Promise<boolean> {
  try {
    const file: any = await request("sharepoint", "GET", `${target.sitePath}/_api/web/GetFileByServerRelativePath(decodedurl=@v)?@v=${odata(`${target.folder}/${name}`)}&$select=Exists,Name`, { sharepointHost: target.host });
    return file?.Exists === true;
  } catch (error) {
    if (error instanceof TeamsError && error.status === 404) return false;
    throw error;
  }
}

export type ChannelUpload = { name: string; uniqueId: string; serverRelativeUrl: string; host: string; siteUrl: string };

const UPLOAD_HEADERS = { scenariotype: "AUO", scenario: "ShareUploadFile" };

/**
 * Uploads one file into the channel folder in the three calls of capture 11: `addStub`, `StartUpload`, `FinishUpload`.
 * A file of 10 MiB or less goes whole in `StartUpload`. A larger file goes in 10 MiB pieces: the first in `StartUpload`,
 * the rest in `ContinueUpload` with the offset so far, then an empty `FinishUpload` (capture 12, a 24,930,161 byte file).
 */
export async function uploadToChannel(target: ChannelTarget, file: LoadedFile): Promise<ChannelUpload> {
  const query = `@folder=${odata(target.folder)}&@file=${odata(file.name)}`;
  const options = { sharepointHost: target.host };
  const stub: any = await request("sharepoint", "POST", `${target.sitePath}/_api/web/GetFolderByServerRelativeUrl(@v)/Files/addStub(@fn)?@v=${odata(target.folder)}&@fn=${odata(file.name)}`, {
    ...options, headers: { ...UPLOAD_HEADERS, prefer: "manualRedirect" },
  });
  const uniqueId = text(stub?.UniqueId);
  if (!uniqueId) throw new TeamsError("SharePoint did not return the id of the new file.");
  const base = `${target.sitePath}/_api/web/GetFileById('${uniqueId}')`;
  const uploadId = randomUUID();
  const blob = Bun.file(file.path);
  let offset = 0;
  for (let first = true; offset < file.size; first = false) {
    const end = Math.min(file.size, offset + CHUNK_BYTES);
    const call = first ? `StartUpload(uploadId=guid'${uploadId}')` : `ContinueUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`;
    const reply: any = await request("sharepoint", "POST", `${base}/${call}?${query}`, {
      ...options, bytes: blob.slice(offset, end), contentType: "application/octet-stream", headers: UPLOAD_HEADERS, timeoutMs: uploadTimeout(end - offset),
    });
    if (Number(reply?.value) !== end) throw new TeamsError(`SharePoint took ${reply?.value ?? "an unknown number of"} bytes of ${file.name}, not ${end}.`);
    offset = end;
  }
  const done: any = await request("sharepoint", "POST", `${base}/FinishUpload(uploadId=guid'${uploadId}',fileOffset=${file.size})?${query}`, {
    ...options, headers: { ...UPLOAD_HEADERS, prefer: "pacToken=N" },
  });
  const serverRelativeUrl = text(done?.ServerRelativeUrl);
  if (!serverRelativeUrl || text(done?.UniqueId) !== uniqueId) throw new TeamsError(`SharePoint did not confirm the upload of ${file.name}.`);
  return { name: text(done?.Name) || file.name, uniqueId, serverRelativeUrl, host: target.host, siteUrl: target.siteUrl };
}

/** The `properties.files` entry for a team-site file, in the shape of capture 11 (request 4722), without the image preview. */
export function channelFileEntry(upload: ChannelUpload) {
  const type = extensionOf(upload.name);
  const fileUrl = `https://${upload.host}${upload.serverRelativeUrl}`;
  return {
    itemid: upload.uniqueId,
    fileName: upload.name,
    fileType: type,
    fileInfo: { itemId: null, fileUrl, siteUrl: upload.siteUrl, serverRelativeUrl: "", shareUrl: null, shareId: null },
    fileChicletState: { serviceName: "teams", state: "active" },
    "@type": FILE_TYPE,
    version: 2,
    id: upload.uniqueId,
    baseUrl: upload.siteUrl,
    objectUrl: fileUrl,
    openUrl: null,
    type,
    title: upload.name,
    state: "active",
    chicletBreadcrumbs: null,
    providerData: "",
    botFileProperties: {},
    isUploadError: null,
    progressComplete: null,
    permissionScope: null,
    filePreview: { previewUrl: "", previewHeight: 0, previewWidth: 0 },
    sharepointIds: null,
    publication: null,
    site: null,
  };
}

export function placeholderChannelUpload(file: LoadedFile, index: number, target: ChannelTarget): ChannelUpload {
  return { name: file.name, uniqueId: `<file-${index + 1}-guid>`, serverRelativeUrl: `${target.folder}/${file.name}`, host: target.host, siteUrl: target.siteUrl };
}

// ---------------------------------------------------------------------------------------------------------------
// One plan for a preview and a send.
// ---------------------------------------------------------------------------------------------------------------

export type UploadPlan =
  | { route: "onedrive"; site: { host: string; path: string }; invite?: InviteSet; sharing: { access: FileAccess; role: FileRole } }
  | { route: "channel"; target: ChannelTarget; conflicts: string[] };

/**
 * Works out where files go and who can open them. It only reads. A chat gets an invite list when it has members other than
 * you. Notes to self and channels get none. A channel plan names the files that already exist in the folder.
 */
export async function planUploads(
  conversation: { id: string; kind: "notes" | "stream" | "channel" | "oneOnOne" | "group"; pending?: { mri: string } },
  files: LoadedFile[],
  selfMri: string,
  sharing: { access: FileAccess; role: FileRole } = { access: "link", role: "edit" },
): Promise<UploadPlan> {
  if (conversation.kind === "channel") {
    const target = await channelTarget(conversation.id);
    await checkChannelFolder(target);
    const conflicts: string[] = [];
    for (const file of files) if (await channelFileExists(target, file.name)) conflicts.push(file.name);
    return { route: "channel", target, conflicts };
  }
  const site = await personalSite();
  if (conversation.kind === "notes" || conversation.kind === "stream" || sharing.access === "private") return { route: "onedrive", site, sharing };
  const members = conversation.pending ? [selfMri, conversation.pending.mri] : await chatMembers(conversation.id);
  const invite = await inviteRecipients(members, selfMri);
  return { route: "onedrive", site, sharing, ...(invite.recipients.length ? { invite } : {}) };
}

/** Where a file will go, in words. */
export function describeDestination(plan: UploadPlan, file: LoadedFile): string {
  return plan.route === "channel"
    ? `${plan.target.siteUrl}${plan.target.folder.slice(plan.target.sitePath.length)}/${file.name}`
    : `OneDrive ${plan.site.host}${plan.site.path}/Documents/${CHAT_FOLDER}/${file.name}`;
}

/**
 * Uploads every file, then invites the recipients to each OneDrive file. Stops at the first failure and throws an
 * `UploadFailure` that lists the files already on SharePoint, so they can be removed by hand. No message is sent here.
 */
export async function runUploads(plan: UploadPlan, files: LoadedFile[]): Promise<unknown[]> {
  const entries: unknown[] = [];
  const leftovers: string[] = [];
  const fail = (error: unknown, current?: LoadedFile): never => {
    const reason = error instanceof Error ? error.message : String(error);
    const uncertain = error instanceof TeamsError && error.uncertain;
    // An upload that did not confirm may still have landed.
    const maybe = current ? [`${describeDestination(plan, current)} (may exist)`] : [];
    throw new UploadFailure(`${reason}`, [...leftovers, ...maybe], uncertain);
  };
  const uploaded: ChatUpload[] = [];
  for (const file of files) {
    try {
      if (plan.route === "channel") {
        const upload = await uploadToChannel(plan.target, file);
        leftovers.push(`https://${upload.host}${upload.serverRelativeUrl}`);
        entries.push(channelFileEntry(upload));
      } else {
        const upload = await uploadToOneDrive(plan.site, file);
        leftovers.push(upload.fileUrl);
        uploaded.push(upload);
        entries.push(chatFileEntry(upload));
      }
    } catch (error) {
      fail(error, file);
    }
  }
  if (plan.route === "onedrive" && plan.invite && plan.sharing.access !== "private") {
    const sharing = { access: plan.sharing.access, role: plan.sharing.role };
    for (const [index, upload] of uploaded.entries()) {
      try {
        const link = await inviteToFile(plan.site, upload.uniqueId, plan.invite.recipients, sharing);
        if (link) Object.assign((entries[index] as any).fileInfo, link);
      } catch (error) {
        fail(new Error(`The invite for ${upload.name} failed: ${error instanceof Error ? error.message : String(error)}`));
      }
    }
  }
  return entries;
}
