import { mkdir, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { encodeConversation, parseMessageRef } from "../refs.ts";
import { paceWrite } from "../pace.ts";
import { parseSharepointHost, sharepointHost, shareId, tenantSharepointHosts } from "../sharepoint.ts";
import { identity } from "../tokens.ts";
import { downloadImage, imageExtension, inlineImagesOf } from "../image.ts";
import { request, TeamsError } from "../transport.ts";

/** The most bytes that `files download` accepts for one file. */
import { MAX_DOWNLOAD_BYTES as MAX_FILE_BYTES } from "../limits.ts";

/** The share lookup that the web client sends before it reads a file (capture 384). */
const SHARE_EXPAND = "site,driveItem(select=@microsoft.graph.downloadUrl,file,id,webDavUrl,sharepointIds,eTag,publication,viewpoint/sharing,mainlink)";

export type FileEntry = { name: string; url: string; siteUrl?: string };

/** Makes a file name safe to write inside one directory. */
export function safeFileName(name: string): string {
  const clean = name.replace(/[\\/\u0000-\u001f]/g, "_").trim();
  return clean === "" || clean === "." || clean === ".." ? "file" : clean;
}

/** Reads `properties.files` of a message. The service stores it as a JSON string. */
export function filesOf(message: any): FileEntry[] {
  let list = message?.properties?.files;
  if (typeof list === "string") {
    try { list = JSON.parse(list); } catch { throw new Error("The message has a files property that cannot be read."); }
  }
  if (!Array.isArray(list)) return [];
  return list.flatMap((file: any) => {
    const url = file?.fileInfo?.fileUrl ?? file?.objectUrl;
    if (typeof url !== "string" || !url) return [];
    const name = typeof file.fileName === "string" && file.fileName ? file.fileName : typeof file.title === "string" ? file.title : "file";
    const siteUrl = file?.fileInfo?.siteUrl;
    return [{ name: safeFileName(name), url, ...(typeof siteUrl === "string" && siteUrl ? { siteUrl } : {}) }];
  });
}

/** A SharePoint list attached to a message: `https://<host>/sites/<site>/Lists/<name>`. It has rows, not bytes. */
export function listAddress(fileUrl: string): { site: string; path: string; name: string } | undefined {
  const match = /^(https:\/\/[^/]+(\/sites\/[^/]+)?)(\/Lists\/([^/?#]+))\/?$/.exec(fileUrl);
  if (!match) return undefined;
  const path = decodeURIComponent(`${match[2] ?? ""}${match[3]}`);
  return { site: match[1]!, path, name: decodeURIComponent(match[4]!) };
}

/** Most rows that one list export reads. SharePoint pages at 5000. */
const MAX_LIST_ROWS = 100_000;

/** Reads every row of a list. Fields that SharePoint adds for its own use (`OData*`) are dropped. */
async function listRows(host: string, address: { site: string; path: string }): Promise<unknown[]> {
  const rows: unknown[] = [];
  const base = `${new URL(address.site).pathname}/_api/web/GetList('${address.path.replace(/'/g, "''")}')/items`;
  let next: string | undefined;
  do {
    const page: any = next
      ? await request("sharepoint", "GET", "", { fullUrl: next, sharepointHost: host, headers: { accept: "application/json;odata=nometadata" } })
      : await request("sharepoint", "GET", base, { query: { $top: "5000" }, sharepointHost: host, headers: { accept: "application/json;odata=nometadata" } });
    for (const row of page?.value ?? []) rows.push(Object.fromEntries(Object.entries(row).filter(([key]) => !key.startsWith("OData"))));
    next = page?.["odata.nextLink"] ?? page?.["@odata.nextLink"];
  } while (next && rows.length < MAX_LIST_ROWS);
  return rows;
}

/** A file named by its SharePoint link, as `search --files` prints it. The query (`?d=...`) and fragment are dropped. */
export function fileFromLink(link: string): FileEntry {
  let url: URL;
  try { url = new URL(link); } catch { throw new Error(`Not a link: ${link}`); }
  url.search = "";
  url.hash = "";
  const last = url.pathname.split("/").filter(Boolean).pop() ?? "";
  let name = last;
  try { name = decodeURIComponent(last); } catch { /* Keep the name as written. */ }
  return { name: safeFileName(name), url: url.toString() };
}

/** `target` is a message ref, whose files are saved, or a SharePoint file link from `search --files`. */
export async function filesDownload(target: string, options: { out?: string; json?: boolean }): Promise<void> {
  let files: FileEntry[];
  // Inline images are read before anything is written, because the file type, and so the name, comes from the bytes.
  const images: Array<{ name: string; bytes: Buffer }> = [];
  if (/^https?:\/\//i.test(target)) files = [fileFromLink(target)];
  else {
    const parsed = parseMessageRef(target);
    if (!parsed) throw new Error("Give a message ref (<conversation-id>#<message-id>) or a SharePoint file link (https://<host>.sharepoint.com/...).");
    const raw: any = await request("chat", "GET", `/users/ME/conversations/${encodeConversation(parsed.conversationId)}/messages/${parsed.messageId}`);
    const message = Array.isArray(raw?.messages) ? raw.messages[0] : raw;
    files = filesOf(message);
    const inline = inlineImagesOf(message);
    if (!files.length && !inline.length) throw new Error("That message has no files and no inline images.");
    for (const [index, image] of inline.entries()) {
      const bytes = await downloadImage(image.id);
      images.push({ name: `${parsed.messageId}-image-${index + 1}.${imageExtension(bytes, image.scope)}`, bytes });
    }
  }
  const names = new Set<string>();
  for (const file of files) {
    if (names.has(file.name)) throw new Error(`The message has two files named ${file.name}. Nothing was downloaded.`);
    names.add(file.name);
  }

  // A file can be on a team site, on anyone's personal site, or on another organization's SharePoint. The token goes
  // only to a *.sharepoint.com host, because the address comes from the message and anyone in the chat can write one.
  const hosts = files.length ? tenantSharepointHosts(await sharepointHost()) : [];
  const hostOf = new Map<FileEntry, string>();
  for (const file of files) {
    let url: URL;
    try { url = new URL(file.url); } catch { throw new Error(`The file ${file.name} has no usable address.`); }
    const host = url.protocol === "https:" ? parseSharepointHost(url.host) : undefined;
    if (!host || host !== url.host.toLowerCase()) {
      throw new Error(`The file ${file.name} is on ${url.host}. This command reads only files on SharePoint (*.sharepoint.com), so your token never goes elsewhere.`);
    }
    hostOf.set(file, host);
  }
  const outside = (host: string) => !hosts.includes(host);

  const dir = options.out ?? ".";
  // A list is saved as its rows, so its file name ends in .json.
  const targetName = (file: FileEntry) => listAddress(file.url) ? `${safeFileName(listAddress(file.url)!.name)}.json` : file.name;
  for (const name of [...files.map(targetName), ...images.map(image => image.name)]) {
    if (await stat(join(dir, name)).then(() => true, () => false)) {
      throw new Error(`${join(dir, name)} already exists. Nothing was downloaded. Choose another --out directory.`);
    }
  }

  const saved: Array<{ name: string; path: string; bytes: number }> = [];
  await mkdir(dir, { recursive: true });
  for (const image of images) {
    const path = join(dir, image.name);
    await writeFile(path, image.bytes, { mode: 0o600, flag: "wx" });
    saved.push({ name: image.name, path, bytes: image.bytes.length });
  }
  for (const file of files) {
    const fileHost = hostOf.get(file)!;
    const list = listAddress(file.url);
    if (list) {
      const rows = await listRows(fileHost, list);
      const path = join(dir, targetName(file));
      const text = `${JSON.stringify(rows, null, 2)}\n`;
      await writeFile(path, text, { mode: 0o600, flag: "wx" });
      saved.push({ name: targetName(file), path, bytes: Buffer.byteLength(text) });
      continue;
    }
    let share: any;
    try {
      share = await request("sharepoint", "GET", `/_api/v2.0/shares/${shareId(file.url)}/`, { query: { expand: SHARE_EXPAND }, sharepointHost: fileHost });
    } catch (error) {
      // Your token is for your own organization. Another organization's site refuses it.
      if (outside(fileHost) && error instanceof TeamsError && (error.status === 401 || error.status === 403)) {
        throw new Error(`${file.name} is on ${fileHost}, which belongs to another organization, and it refused your token (${error.status}). Open it in Teams.`);
      }
      throw error;
    }
    const item = share?.driveItem;
    const link = item?.["@content.downloadUrl"] ?? item?.["@microsoft.graph.downloadUrl"];
    // A Loop component or a link to a folder has no bytes to save.
    if (typeof link !== "string") throw new Error(`SharePoint gave no download address for ${file.name}. It may be a Loop component, a folder, or a link.`);
    const bytes: Buffer = await request("sharepoint", "GET", "", { fullUrl: link, binary: true, maxBytes: MAX_FILE_BYTES, sharepointHost: fileHost });
    const path = join(dir, file.name);
    await writeFile(path, bytes, { mode: 0o600, flag: "wx" });
    saved.push({ name: file.name, path, bytes: bytes.length });
  }
  if (options.json) console.log(JSON.stringify({ ok: true, files: saved }, null, 2));
  else for (const file of saved) console.log(`saved ${file.path} (${file.bytes} bytes)`);
}

export type FilesDeleteOptions = { json?: boolean; allowWrite?: boolean; noWait?: boolean };

type FileStatus = "recycle" | "already gone" | "refused";

type DeletePlan = {
  name: string;
  host: string;
  /** The file address on its site: `<site>/_api/web/GetFileByServerRelativePath(decodedurl='...')`. */
  address: string;
  path: string;
  status: FileStatus;
  size?: number;
  uploader?: string;
  uploaderEmail?: string;
  reason?: string;
};

const JSON_ACCEPT = { accept: "application/json;odata=nometadata" };

/** The path of a file on its site, as a quoted OData string: `'` is doubled, and each segment is percent-encoded. */
export function serverRelativeAddress(sitePath: string, filePath: string): string {
  const quoted = filePath.replace(/'/g, "''").split("/").map(segment => encodeURIComponent(segment)).join("/");
  return `${sitePath}/_api/web/GetFileByServerRelativePath(decodedurl='${quoted}')`;
}

function isNotFound(error: unknown): boolean {
  return error instanceof TeamsError && error.status === 404;
}

/**
 * Moves the files of one of your messages to the SharePoint recycle bin. Preview by default. Only `--allow-write`
 * recycles. It never hard-deletes, and it does not change the Teams message.
 */
export async function filesDelete(ref: string, options: FilesDeleteOptions): Promise<void> {
  const parsed = parseMessageRef(ref);
  if (!parsed) throw new Error("Give a message ref: <conversation-id>#<message-id>.");
  const raw: any = await request("chat", "GET", `/users/ME/conversations/${encodeConversation(parsed.conversationId)}/messages/${parsed.messageId}`);
  const message = Array.isArray(raw?.messages) ? raw.messages[0] : raw;
  const files = filesOf(message);
  if (!files.length) {
    if (message?.properties?.deletetime) {
      throw new Error("That message is deleted, so Teams no longer lists its files. They stay on SharePoint. Run `files delete` before `delete`. Nothing was changed.");
    }
    throw new Error("That message has no files.");
  }

  // The token goes only to a *.sharepoint.com host, because the address comes from the message.
  const plans: DeletePlan[] = [];
  const who = await identity();
  for (const file of files) {
    let url: URL;
    try { url = new URL(file.url); } catch { throw new Error(`The file ${file.name} has no usable address.`); }
    const host = url.protocol === "https:" ? parseSharepointHost(url.host) : undefined;
    if (!host || host !== url.host.toLowerCase()) {
      throw new Error(`The file ${file.name} is on ${url.host}. This command acts only on files on SharePoint (*.sharepoint.com), so your token never goes elsewhere. Nothing was changed.`);
    }
    if (listAddress(file.url)) throw new Error(`${file.name} is a SharePoint list, not a file. This command does not delete lists. Nothing was changed.`);
    let sitePath: string;
    try { sitePath = new URL(file.siteUrl ?? "").pathname.replace(/\/+$/, ""); } catch { throw new Error(`The file ${file.name} names no site, so its path is unknown. Nothing was changed.`); }
    const path = decodeURIComponent(url.pathname);
    const address = serverRelativeAddress(sitePath, path);
    const plan: DeletePlan = { name: file.name, host, address, path, status: "recycle" };
    try {
      const info: any = await request("sharepoint", "GET", address, { query: { $select: "Name,Length,TimeCreated" }, sharepointHost: host, headers: JSON_ACCEPT });
      plan.size = Number(info?.Length);
      const author: any = await request("sharepoint", "GET", `${address}/Author`, { query: { $select: "Title,Email" }, sharepointHost: host, headers: JSON_ACCEPT });
      plan.uploader = typeof author?.Title === "string" ? author.Title : undefined;
      plan.uploaderEmail = typeof author?.Email === "string" ? author.Email : undefined;
      if (!plan.uploaderEmail || plan.uploaderEmail.toLowerCase() !== who.upn.toLowerCase()) {
        plan.status = "refused";
        plan.reason = `uploaded by ${plan.uploader ?? "someone else"}${plan.uploaderEmail ? ` (${plan.uploaderEmail})` : ""}, not by you`;
      }
    } catch (error) {
      if (isNotFound(error)) plan.status = "already gone";
      else if (error instanceof TeamsError && (error.status === 401 || error.status === 403)) {
        throw new Error(`${file.name} is on ${host}, and it refused your token (${error.status}). Nothing was changed.`);
      } else throw error;
    }
    plans.push(plan);
  }

  const toRecycle = plans.filter(plan => plan.status === "recycle");
  const refused = plans.filter(plan => plan.status === "refused");
  const NOTE = "The Teams message is not changed. Use `delete` for the message, after this.";
  const BIN = "Recycled files can be restored from the site's recycle bin.";

  if (!options.allowWrite) {
    if (options.json) {
      console.log(JSON.stringify({
        ok: true,
        dry_run: true,
        effect: "file",
        method: "POST",
        action: "recycle",
        count: toRecycle.length,
        files: plans.map(plan => ({ name: plan.name, host: plan.host, path: plan.path, size: plan.size, uploader: plan.uploader, status: plan.status, ...(plan.reason ? { reason: plan.reason } : {}) })),
        note: `${NOTE} ${BIN}`,
      }, null, 2));
      return;
    }
    console.log(`Dry run (effect: file): would move ${toRecycle.length} file(s) to the SharePoint recycle bin:`);
    for (const plan of plans) {
      const detail = plan.status === "recycle" ? `${plan.size} bytes, uploaded by ${plan.uploader ?? "unknown"}, recycle` : plan.status === "refused" ? `refused: ${plan.reason}` : "already gone";
      console.log(`  ${plan.name}  ${detail}`);
    }
    console.log(NOTE);
    console.log(BIN);
    return;
  }

  // A file of another person stops the whole run before any file is recycled.
  if (refused.length) throw new Error(`${refused.map(plan => `${plan.name} was ${plan.reason}`).join("; ")}. This command recycles only your own files. Nothing was recycled.`);

  const results: Array<{ name: string; status: "recycled" | "already gone" | "failed" | "uncertain"; recycleBinItemId?: string; error?: string }> = [];
  for (const plan of plans) {
    if (plan.status === "already gone") {
      results.push({ name: plan.name, status: "already gone" });
      if (!options.json) console.log(`already gone ${plan.name}`);
      continue;
    }
    try {
      await paceWrite(who.tenantId, { noWait: options.noWait, send: false });
      const reply: any = await request("sharepoint", "POST", `${plan.address}/recycle()`, { sharepointHost: plan.host, headers: JSON_ACCEPT });
      const id = typeof reply?.value === "string" ? reply.value : undefined;
      results.push({ name: plan.name, status: "recycled", ...(id ? { recycleBinItemId: id } : {}) });
      if (!options.json) console.log(`recycled ${plan.name}${id ? ` (recycle bin item ${id})` : ""}`);
    } catch (error) {
      if (isNotFound(error)) {
        results.push({ name: plan.name, status: "already gone" });
        if (!options.json) console.log(`already gone ${plan.name}`);
        continue;
      }
      const uncertain = error instanceof TeamsError && error.uncertain;
      const text = error instanceof Error ? error.message : String(error);
      results.push({ name: plan.name, status: uncertain ? "uncertain" : "failed", error: text });
      if (!options.json) console.log(`${uncertain ? "uncertain" : "failed"} ${plan.name}: ${text}${uncertain ? " It may be recycled. Run the command again: a file that is gone is reported as already gone." : ""}`);
    }
  }
  const bad = results.filter(item => item.status === "failed" || item.status === "uncertain").length;
  if (options.json) console.log(JSON.stringify({ ok: bad === 0, files: results, note: `${NOTE} ${BIN}` }, null, 2));
  else {
    console.log(`Recycled ${results.filter(item => item.status === "recycled").length}, already gone ${results.filter(item => item.status === "already gone").length}, failed ${bad}.`);
    console.log(NOTE);
    console.log(BIN);
  }
  if (bad) process.exitCode = 1;
}
